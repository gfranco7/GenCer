import os
import asyncio
from fastapi import FastAPI, UploadFile, Form
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from core.certificados import generar_certificados
from onedrive_client import OneDriveClient

app = FastAPI()

# Montar carpeta static (para logo, estilos, etc.)
app.mount("/static", StaticFiles(directory="static"), name="static")

# Configuración OneDrive
ACCESS_TOKEN = os.getenv("ONEDRIVE_TOKEN")  # exporta tu token de Azure
ONEDRIVE_PARENT_ID = "ID_DE_LA_CARPETA_DATACAMPUS"  # carpeta compartida en OneDrive

@app.get("/", response_class=HTMLResponse)
async def index():
    with open("templates/index.html", "r", encoding="utf-8") as f:
        return f.read()

@app.post("/procesar", response_class=HTMLResponse)
async def procesar(excel_file: UploadFile):
    certificados_por_compania, output_dir = await generar_certificados(excel_file.file)

    if not certificados_por_compania:
        return HTMLResponse("<h2>No hay certificados por generar</h2>", status_code=400)

    # Cliente de OneDrive
    client = OneDriveClient(ACCESS_TOKEN)

    tasks = []
    for compania, archivos in certificados_por_compania.items():
        # Crear subcarpeta en OneDrive
        folder_task = asyncio.to_thread(client.create_folder, ONEDRIVE_PARENT_ID, compania)
        folder_id = await folder_task

        for archivo in archivos:
            file_path = os.path.join(output_dir, compania.replace(" ", "_"), archivo)
            # Subir cada archivo en paralelo
            tasks.append(asyncio.to_thread(client.upload_file, folder_id, file_path))

    await asyncio.gather(*tasks)

    # Renderizar página de éxito
    with open("templates/success.html", "r", encoding="utf-8") as f:
        html = f.read()
        html = html.replace("{{ ruta }}", "OneDrive/Datacampus")

    return HTMLResponse(content=html)
