import os
import platform
import subprocess
from datetime import datetime
from collections import defaultdict
from pathlib import Path

import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert  # solo Windows
import asyncio
import aiofiles

ON_WINDOWS = platform.system() == "Windows"

def get_downloads_folder():
    home = Path.home()
    downloads = home / "Downloads"
    if not downloads.exists():
        downloads = home
    return downloads

async def generar_certificados(excel_file) -> tuple[dict, str]:
    """
    Genera certificados a partir de un archivo Excel.
    Devuelve un diccionario con certificados agrupados por compañía y la ruta base.
    """
    df = pd.read_excel(excel_file)

    # Se valida que haya al menos un "no" en la columna certificado
    if not (df["certificado"].astype(str).str.lower() == "no").any():
        return {}, ""

    plantilla = DocxTemplate("plantilla.docx")

    fecha_hoy = datetime.today().strftime("%Y-%m-%d")
    downloads_folder = get_downloads_folder()
    output_dir = downloads_folder / f"certificados_{fecha_hoy}"
    os.makedirs(output_dir, exist_ok=True)

    certificados_por_compania = defaultdict(list)

    tasks = []

    for _, row in df.iterrows():
        if str(row["certificado"]).lower() == "no":
            contexto = {
                "NOMBRE": row["nombre"],
                "CEDULA": row["cedula"],
                "FECHA": row["fecha"].strftime("%d/%m/%Y") if not pd.isna(row["fecha"]) else "",
                "COMPANIA": row["compañia"]
            }

            compania_folder = output_dir / str(row["compañia"]).replace(" ", "_")
            os.makedirs(compania_folder, exist_ok=True)

            nombre_base = f"certificado_{row['nombre'].replace(' ', '_')}"
            docx_file = compania_folder / f"{nombre_base}.docx"
            pdf_file = compania_folder / f"{nombre_base}.pdf"

            # Renderizar docx
            plantilla.render(contexto)
            plantilla.save(docx_file)

            # Crear tarea asíncrona para convertir a PDF
            tasks.append(convert_to_pdf(docx_file, pdf_file))

            certificados_por_compania[row["compañia"]].append(f"{nombre_base}.pdf")

    # Ejecutar conversiones en paralelo
    await asyncio.gather(*tasks)

    return certificados_por_compania, str(output_dir)

async def convert_to_pdf(docx_file, pdf_file):
    """Convierte un DOCX a PDF usando docx2pdf (Windows) o LibreOffice (Linux/Mac)."""
    if ON_WINDOWS:
        # docx2pdf es síncrono ---> correr en un thread
        await asyncio.to_thread(convert, str(docx_file), str(pdf_file))
        os.remove(docx_file)
    else:
        # LibreOffice --> correr en proceso asíncrono
        process = await asyncio.create_subprocess_exec(
            "soffice", "--headless", "--convert-to", "pdf",
            "--outdir", str(pdf_file.parent), str(docx_file)
        )
        await process.communicate()
        os.remove(docx_file)


"""
GenCer (Repositorio GitHub):
-agents: data_campus_agent.py
-auth: auth_manager.py
-statics: logo, estilos, etc.
-templates: index.html, success.html
-venv
-certificados.py
-onedrive_client.py
-main.py
-plantilla.docx
-readme.md
-.env
-requirements.txt


"""