from fastapi import FastAPI, APIRouter, HTTPException, UploadFile, File
from fastapi.responses import StreamingResponse
from dotenv import load_dotenv
from starlette.middleware.cors import CORSMiddleware
import os
import logging
from pathlib import Path
from pydantic import BaseModel, Field, ConfigDict
from typing import List, Optional
import uuid
from datetime import datetime, timezone
from enum import Enum
import pandas as pd
import io
import re
from contextlib import asynccontextmanager

ROOT_DIR = Path(__file__).parent
load_dotenv(ROOT_DIR / '.env')

# In-memory storage (reemplaza MongoDB)
cruces_storage: List[dict] = []

@asynccontextmanager
async def lifespan(_: FastAPI):
    """Arranque del backend sin datos precargados"""
    logger.info(f"Backend iniciado en modo vacío con {len(cruces_storage)} cruces")
    yield


app = FastAPI(lifespan=lifespan)

api_router = APIRouter(prefix="/api")


# Enums
class EstadoCruce(str, Enum):
    PENDIENTE = "Pendiente"
    HECHO = "Hecho"
    POR_HACER = "Por hacer"


class TipoServicio(str, Enum):
    REG_PT = "REG PT"
    REG_PARTES = "REG PARTES"
    LTL_1ERO = "LTL 1ERO."
    LTL_2DO = "LTL 2DO."
    LTL_3ERO = "LTL 3ERO."
    BRISTOL_1ERO = "BRISTOL 1ERO."
    BRISTOL_2DO = "BRISTOL 2DO."
    URGENCIA = "URGENCIA"


# Models
class CruceBase(BaseModel):
    numero_unidad: str
    placas: str
    hora_estimada: str  # formato HH:MM
    tipo_servicio: str
    operador: str
    ubicacion: str = "PARQUE INDUSTRIAL FINSA"
    orden: int = 0
    sello_aduana: Optional[str] = None
    fecha: Optional[str] = None


class CruceCreate(CruceBase):
    estado: str = EstadoCruce.PENDIENTE


class CruceUpdate(BaseModel):
    numero_unidad: Optional[str] = None
    placas: Optional[str] = None
    hora_estimada: Optional[str] = None
    tipo_servicio: Optional[str] = None
    operador: Optional[str] = None
    ubicacion: Optional[str] = None
    orden: Optional[int] = None
    estado: Optional[str] = None
    sello_aduana: Optional[str] = None
    fecha: Optional[str] = None


class Cruce(CruceBase):
    model_config = ConfigDict(extra="ignore")
    
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    estado: str = EstadoCruce.PENDIENTE
    created_at: str = Field(default_factory=lambda: datetime.now(timezone.utc).isoformat())
    updated_at: str = Field(default_factory=lambda: datetime.now(timezone.utc).isoformat())


# Routes
@api_router.get("/")
async def root():
    return {"message": "API de Cruces Logísticos FROMEX - Modo Local"}


@api_router.get("/cruces", response_model=List[Cruce])
async def get_cruces(
    estado: Optional[str] = None, 
    tipo_servicio: Optional[str] = None,
    search: Optional[str] = None
):
    """Obtener todos los cruces, opcionalmente filtrados"""
    result = cruces_storage.copy()
    
    if estado:
        result = [c for c in result if c.get("estado") == estado]
    
    if tipo_servicio:
        result = [c for c in result if c.get("tipo_servicio") == tipo_servicio]
    
    if search:
        search_lower = search.lower()
        result = [c for c in result if 
                  search_lower in c.get("numero_unidad", "").lower() or
                  search_lower in c.get("placas", "").lower() or
                  search_lower in c.get("operador", "").lower() or
                  search_lower in str(c.get("sello_aduana", "")).lower()]
    
    # Ordenar por orden y hora
    result.sort(key=lambda x: (x.get("orden", 0), x.get("hora_estimada", "")))
    return result


@api_router.get("/cruces/pendientes", response_model=List[Cruce])
async def get_cruces_pendientes():
    """Obtener solo los cruces pendientes o por hacer"""
    result = [c for c in cruces_storage if c.get("estado") in [EstadoCruce.PENDIENTE, EstadoCruce.POR_HACER]]
    result.sort(key=lambda x: (x.get("orden", 0), x.get("hora_estimada", "")))
    return result


@api_router.get("/cruces/export")
async def export_cruces(estado: Optional[str] = None):
    """Exportar cruces a Excel"""
    result = cruces_storage.copy()
    if estado:
        result = [c for c in result if c.get("estado") == estado]
    
    result.sort(key=lambda x: (x.get("orden", 0), x.get("hora_estimada", "")))
    
    df = pd.DataFrame(result)
    
    column_mapping = {
        'numero_unidad': 'NO. DE UNIDAD',
        'placas': '# DE PLACAS / # DE CAJA',
        'hora_estimada': 'HORA ESTIMADA DE LLEGADA',
        'tipo_servicio': 'CAMION / CAJA PARA',
        'operador': 'NOMBRE DEL OPERADOR',
        'sello_aduana': '# DE SELLO / ADUANA',
        'fecha': 'FECHA',
        'estado': 'ESTADO',
        'ubicacion': 'UBICACION'
    }
    
    rename_cols = {k: v for k, v in column_mapping.items() if k in df.columns}
    df = df.rename(columns=rename_cols)
    
    cols_to_drop = ['id', 'created_at', 'updated_at', 'orden']
    df = df.drop(columns=[c for c in cols_to_drop if c in df.columns], errors='ignore')
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Cruces')
    output.seek(0)
    
    filename = f"cruces_fromex_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@api_router.get("/health")
async def health_check():
    """Endpoint para que el DriverDisplay verifique que el servidor está activo"""
    return {
        "status": "ok", 
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "storage_count": len(cruces_storage)
    }

@api_router.get("/cruces/{cruce_id}", response_model=Cruce)
async def get_cruce(cruce_id: str):
    """Obtener un cruce por ID"""
    for cruce in cruces_storage:
        if cruce.get("id") == cruce_id:
            return cruce
    raise HTTPException(status_code=404, detail="Cruce no encontrado")


@api_router.post("/cruces", response_model=Cruce)
async def create_cruce(cruce_data: CruceCreate):
    """Crear un nuevo cruce"""
    cruce = Cruce(**cruce_data.model_dump())
    doc = cruce.model_dump()
    cruces_storage.append(doc)
    return cruce


@api_router.post("/cruces/import-excel")
async def import_excel(file: UploadFile = File(...)):
    """Importar cruces desde un archivo Excel - Solo datos de transportistas"""
    global cruces_storage

    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="El archivo debe ser un Excel (.xlsx o .xls)")
    
    try:
        contents = await file.read()
        df_raw = pd.read_excel(io.BytesIO(contents), header=None)
        
        def normalizar(texto):
            return re.sub(r'[^a-z0-9]', '', str(texto).lower())

        header_row_idx = None
        for idx, row in df_raw.iterrows():
            row_limpia = [normalizar(cell) for cell in row if pd.notna(cell)]
            
            # Buscar fila de encabezados de datos (transportistas)
            if 'nodeunidad' in row_limpia or 'unidad' in row_limpia or 'nounidad' in row_limpia:
                header_row_idx = idx
                break
        
        if header_row_idx is None:
            raise HTTPException(
                status_code=400, 
                detail="No se encontró la tabla de transportistas. Verifica que exista una columna 'NO. DE UNIDAD' o similar"
            )

        df = df_raw.iloc[header_row_idx + 1:].copy()
        df.columns = [normalizar(col) for col in df_raw.iloc[header_row_idx]]

        cruces_a_importar = []
        imported_count = 0

        for index, row in df.iterrows():
            u_val = str(row.get('nodeunidad', row.get('nounidad', row.get('unidad', '')))).strip().lower()
            
            # Saltar filas vacías o de texto decorativo
            if u_val in ['', 'nan', 'none', 'total', 'firmas', 'fecha', 'observaciones']:
                continue
            
            # Verificar si es un número válido (unidad)
            try:
                int(float(u_val.split('.')[0]))
            except:
                continue

            try:
                numero_unidad = u_val.split('.')[0]
                
                # Mapeo flexible de columnas
                placas = str(row.get('deplacasdecaja', row.get('placas', row.get('decaja', '')))).strip().upper()
                hora = parse_hora(row.get('horaestimadadellegada', row.get('hora', row.get('horaestimada', ''))))
                tipo = str(row.get('camioncajapara', row.get('tipo', row.get('servicio', 'REG PT')))).strip().upper()
                operador = str(row.get('nombredeloperador', row.get('operador', row.get('nombre', '')))).strip().upper()
                sello = str(row.get('desello', row.get('sello', row.get('aduananlco', '')))).strip() if pd.notna(row.get('desello', row.get('sello', row.get('aduananlco')))) else None
                
                # Limpiar valores 'nan'
                if placas.lower() == 'nan':
                    placas = ''
                if operador.lower() == 'nan':
                    operador = ''
                if sello and sello.lower() == 'nan':
                    sello = None
                
                cruce_data = {
                    "numero_unidad": numero_unidad,
                    "placas": placas,
                    "hora_estimada": hora,
                    "tipo_servicio": tipo if tipo and tipo.lower() != 'nan' else "REG PT",
                    "operador": operador,
                    "ubicacion": "PARQUE INDUSTRIAL FINSA",
                    "orden": len(cruces_a_importar) + 1,
                    "estado": EstadoCruce.PENDIENTE,
                    "sello_aduana": sello,
                    "fecha": datetime.now(timezone.utc).strftime('%Y-%m-%d')
                }

                cruce_validado = CruceCreate(**cruce_data)
                cruces_a_importar.append(Cruce(**cruce_validado.model_dump()).model_dump())
                imported_count += 1

            except Exception as e:
                print(f"Error en fila {index}: {e}")
                continue

        if not cruces_a_importar:
            raise HTTPException(status_code=400, detail="No se encontraron registros válidos de transportistas en el archivo")

        # Agregar a almacenamiento en memoria
        cruces_storage.extend(cruces_a_importar)
        
        return {
            "message": f"Éxito: Se importaron {imported_count} transportistas",
            "count": imported_count
        }

    except HTTPException as he:
        raise he
    except Exception as e:
        print(f"FATAL ERROR: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error interno al procesar el archivo: {str(e)}")


def parse_hora(value) -> str:
    """Parsear hora desde Excel"""
    if pd.isna(value) or value == '':
        return "00:00"
    
    if isinstance(value, (int, float)):
        total_minutes = int(round(value * 24 * 60))
        hours = (total_minutes // 60) % 24
        minutes = total_minutes % 60
        return f"{hours:02d}:{minutes:02d}"
    
    text = str(value).strip()
    match = re.match(r'^(\d{1,2}):(\d{2})', text)
    if match:
        return f"{int(match.group(1)):02d}:{match.group(2)}"
    
    return text if text else "00:00"


@api_router.put("/cruces/{cruce_id}", response_model=Cruce)
async def update_cruce(cruce_id: str, cruce_data: CruceUpdate):
    """Actualizar un cruce existente"""
    for i, cruce in enumerate(cruces_storage):
        if cruce.get("id") == cruce_id:
            update_data = {k: v for k, v in cruce_data.model_dump().items() if v is not None}
            update_data["updated_at"] = datetime.now(timezone.utc).isoformat()
            cruces_storage[i].update(update_data)
            return cruces_storage[i]
    
    raise HTTPException(status_code=404, detail="Cruce no encontrado")


@api_router.patch("/cruces/{cruce_id}/completar", response_model=Cruce)
async def completar_cruce(cruce_id: str):
    """Marcar un cruce como completado"""
    for i, cruce in enumerate(cruces_storage):
        if cruce.get("id") == cruce_id:
            cruces_storage[i]["estado"] = EstadoCruce.HECHO
            cruces_storage[i]["updated_at"] = datetime.now(timezone.utc).isoformat()
            return cruces_storage[i]
    
    raise HTTPException(status_code=404, detail="Cruce no encontrado")


@api_router.delete("/cruces/{cruce_id}")
async def delete_cruce(cruce_id: str):
    """Eliminar un cruce"""
    global cruces_storage
    original_len = len(cruces_storage)
    cruces_storage = [c for c in cruces_storage if c.get("id") != cruce_id]
    
    if len(cruces_storage) == original_len:
        raise HTTPException(status_code=404, detail="Cruce no encontrado")
    
    return {"message": "Cruce eliminado exitosamente"}


@api_router.delete("/cruces")
async def delete_all_cruces():
    """Eliminar todos los cruces"""
    global cruces_storage
    count = len(cruces_storage)
    cruces_storage = []
    return {"message": f"Se eliminaron {count} cruces"}


@api_router.post("/cruces/sync")
async def sync_cruces(cruces: List[dict]):
    """Sincronizar cruces desde el frontend (localStorage)"""
    global cruces_storage
    cruces_storage = cruces
    return {"message": f"Sincronizados {len(cruces)} cruces", "count": len(cruces)}


# Include the router in the main app
app.include_router(api_router)

app.add_middleware(
    CORSMiddleware,
    allow_credentials=True,
    allow_origins=os.environ.get('CORS_ORIGINS', '*').split(','),
    allow_methods=["*"],
    allow_headers=["*"],
)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(
        app,
        host=os.environ.get("HOST", "127.0.0.1"),
        port=int(os.environ.get("PORT", "8001")),
    )
