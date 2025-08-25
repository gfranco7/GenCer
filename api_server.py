from fastapi import FastAPI, HTTPException, UploadFile, File, Form, Depends
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import List, Optional, Dict, Any
import pandas as pd
import io
import json
from datetime import datetime
from auth.auth_manager import AuthManager
from one_drive.OD_manager import *
from one_drive.OD_manager import OneDriveManager

app = FastAPI(
    title="OneDrive Manager API",
    description="API para gestionar archivos y carpetas en OneDrive (Datacampus)",
    version="1.0.0"
)

# Configurar CORS para desarrollo
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # En producción, especificar dominios exactos
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Instancia global del manager (en producción usar dependency injection)
od_manager = None

class ItemResponse(BaseModel):
    id: str
    name: str
    type: str
    size: int
    created_datetime: str
    modified_datetime: str

class FolderContentsResponse(BaseModel):
    current_folder_id: str
    current_path: List[str]
    items: List[ItemResponse]
    total_items: int
    folders_count: int
    files_count: int

class CreateFileRequest(BaseModel):
    filename: str
    folder_id: Optional[str] = None
    data: Optional[Dict[str, List[Any]]] = None

class CreateFolderRequest(BaseModel):
    folder_name: str
    parent_folder_id: Optional[str] = None
    
class UpdateExcelRequest(BaseModel):
    file_id: str
    data: Dict[str, List[Any]]

class AuthResponse(BaseModel):
    status: str 
    message: str
    expires_in: Optional[int] = None

# Dependency para verificar autenticación
async def get_manager():
    global od_manager
    if od_manager is None:
        raise HTTPException(status_code=401, detail="No autenticado. Llama primero a /auth/login")
    return od_manager

# Endpoints de autenticación
@app.post("/auth/login", response_model=AuthResponse)
async def login():
    global od_manager
    try:
        # Obtener token primero antes de crear OneDriveManager
        print(" Iniciando proceso de autenticación...")
        auth = AuthManager()
        token = auth.get_token()
        
        print(" Token obtenido, creando OneDriveManager...")
        od_manager = OneDriveManager(token)  #  Pasar el token al constructor
        
        print(" Inicializando datacampus...")
        od_manager.initialize_datacampus()

        print(" Autenticación completada exitosamente")
        return AuthResponse(
            status="success",
            message="Autenticación exitosa",
            expires_in=token.get("expires_in", 3600) if isinstance(token, dict) else 3600
        )
    except Exception as e:
        print(f" Error en autenticación: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error en autenticación: {str(e)}")


@app.get("/auth/status", response_model=AuthResponse)
async def auth_status():
    global od_manager
    if od_manager is None or od_manager.token is None:
        raise HTTPException(status_code=401, detail="No autenticado")
    
    return AuthResponse(status="authenticated", message="Usuario autenticado")

# Endpoints de navegación y listado
@app.get("/folders", response_model=FolderContentsResponse)
async def list_folder_contents(
    folder_id: Optional[str] = None,
    manager: OneDriveManager = Depends(get_manager)
):
    try:
        items = manager.list_folder_contents(folder_id)
        
        items_response = []
        for item in items:
            items_response.append(ItemResponse(
                id=item.id,
                name=item.name,
                type=item.type,
                size=item.size,
                created_datetime=item.created_datetime,
                modified_datetime=item.modified_datetime
            ))
        
        folders_count = sum(1 for item in items if item.type == 'folder')
        files_count = sum(1 for item in items if item.type == 'file')
        
        return FolderContentsResponse(
            current_folder_id=folder_id or manager.datacampus_root_id,
            current_path=["datacampus"],  # TODO: implementar tracking de path
            items=items_response,
            total_items=len(items),
            folders_count=folders_count,
            files_count=files_count
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error al listar carpeta: {str(e)}")

@app.get("/items/{item_id}")
async def get_item_info(
    item_id: str,
    manager: OneDriveManager = Depends(get_manager)
):
    """Obtener información detallada de un elemento"""
    try:
        item_info = manager.get_item_info(item_id)
        return item_info
    except Exception as e:
        raise HTTPException(status_code=404, detail=f"Elemento no encontrado: {str(e)}")

@app.get("/search/{item_name}")
async def search_item(
    item_name: str,
    folder_id: Optional[str] = None,
    manager: OneDriveManager = Depends(get_manager)
):
    """Buscar un elemento por nombre"""
    try:
        item = manager.find_item_by_name(item_name, folder_id)
        if item:
            return ItemResponse(
                id=item.id,
                name=item.name,
                type=item.type,
                size=item.size,
                created_datetime=item.created_datetime,
                modified_datetime=item.modified_datetime
            )
        else:
            raise HTTPException(status_code=404, detail=f"Elemento '{item_name}' no encontrado")
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error en búsqueda: {str(e)}")

# Endpoints CRUD para archivos
@app.post("/files/excel")
async def create_excel_file(
    request: CreateFileRequest,
    manager: OneDriveManager = Depends(get_manager)
):
    try:
        folder_id = request.folder_id or manager.datacampus_root_id
        
        # Convertir datos a DataFrame si se proporcionan
        df = None
        if request.data:
            df = pd.DataFrame(request.data)
        
        result = manager.create_excel_file(folder_id, request.filename, df)
        return {"message": "Archivo creado exitosamente", "file_id": result["id"], "name": result["name"]}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error al crear archivo: {str(e)}")

@app.get("/files/{file_id}/download")
async def download_file(
    file_id: str,
    manager: OneDriveManager = Depends(get_manager)
):
    try:
        # Para archivos Excel, convertir a CSV para facilitar descarga
        df = manager.read_excel_file(file_id)
        
        # Crear buffer CSV
        csv_buffer = io.StringIO()
        df.to_csv(csv_buffer, index=False)
        csv_buffer.seek(0)
        
        # Obtener info del archivo para el nombre
        file_info = manager.get_item_info(file_id)
        filename = file_info['name'].replace('.xlsx', '.csv')
        
        return StreamingResponse(
            io.BytesIO(csv_buffer.getvalue().encode()),
            media_type="text/csv",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error al descargar archivo: {str(e)}")

@app.get("/files/{file_id}/content")
async def get_file_content(
    file_id: str,
    manager: OneDriveManager = Depends(get_manager)
):
    try:
        df = manager.read_excel_file(file_id)
        
        # Convertir DataFrame a formato JSON amigable para frontend
        content = {
            "columns": df.columns.tolist(),
            "data": df.values.tolist(),
            "shape": df.shape,
            "info": {
                "rows": len(df),
                "columns": len(df.columns),
                "column_types": df.dtypes.to_dict()
            }
        }
        return content
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error al leer archivo: {str(e)}")

@app.put("/files/{file_id}/content")
async def update_file_content(
    file_id: str,
    request: UpdateExcelRequest,
    manager: OneDriveManager = Depends(get_manager)
):
    try:
        # Convertir datos a DataFrame
        df = pd.DataFrame(request.data)
        
        result = manager.update_excel_file(file_id, df)
        return {"message": "Archivo actualizado exitosamente", "file_id": file_id}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error al actualizar archivo: {str(e)}")

@app.post("/files/upload")
async def upload_file(
    file: UploadFile = File(...),
    folder_id: Optional[str] = Form(None),
    manager: OneDriveManager = Depends(get_manager)
):
    try:
        if not file.filename.endswith('.xlsx'):
            raise HTTPException(status_code=400, detail="Solo se permiten archivos Excel (.xlsx)")
        
        folder_id = folder_id or manager.datacampus_root_id
        
        # Leer contenido del archivo
        content = await file.read()
        df = pd.read_excel(io.BytesIO(content))
        
        # Crear archivo en OneDrive
        result = manager.create_excel_file(folder_id, file.filename, df)
        return {"message": "Archivo subido exitosamente", "file_id": result["id"]}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error al subir archivo: {str(e)}")

# Endpoints CRUD para carpetas
@app.post("/folders")
async def create_folder(
    request: CreateFolderRequest,
    manager: OneDriveManager = Depends(get_manager)
):
    try:
        parent_folder_id = request.parent_folder_id or manager.datacampus_root_id
        result = manager.create_folder(parent_folder_id, request.folder_name)
        return {"message": "Carpeta creada exitosamente", "folder_id": result["id"], "name": result["name"]}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error al crear carpeta: {str(e)}")

# Endpoints de eliminación
@app.delete("/items/{item_id}")
async def delete_item(
    item_id: str,
    manager: OneDriveManager = Depends(get_manager)
):
    try:
        manager.delete_item(item_id)
        return {"message": "Elemento eliminado exitosamente"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error al eliminar elemento: {str(e)}")

# Endpoints utilitarios
@app.get("/")
async def root():
    """Endpoint raíz con información de la API"""
    return {
        "message": "OneDrive Manager API",
        "version": "1.0.0",
        "endpoints": {
            "auth": "/auth/login, /auth/status",
            "folders": "/folders, /folders (POST)",
            "files": "/files/excel (POST), /files/{id}/content, /files/{id}/download",
            "items": "/items/{id}, /items/{id} (DELETE)",
            "search": "/search/{name}",
            "docs": "/docs"
        }
    }

@app.get("/health")
async def health_check():
    return {
        "status": "healthy",
        "timestamp": datetime.now().isoformat(),
        "authenticated": od_manager is not None and od_manager.token is not None
    }

# Manejo de errores globales
@app.exception_handler(Exception)
async def global_exception_handler(request, exc):
    return HTTPException(
        status_code=500,
        detail=f"Error interno del servidor: {str(exc)}"
    )

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000, reload=True)



"""
if __name__ == "__main__":
    import uvicorn
    import os
    host = os.getenv("HOST", "0.0.0.0")
    port = int(os.getenv("PORT", 8000))
    uvicorn.run(app, host=host, port=port)
"""