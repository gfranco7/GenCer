import io
import os
import platform
import subprocess
import tempfile
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
from collections import defaultdict
import asyncio
from docx2pdf import convert  # solo Windows

ON_WINDOWS = platform.system() == "Windows"

async def generar_certificados_desde_excel(excel_input) -> tuple[dict, pd.DataFrame]:
    """
    Genera certificados en memoria a partir de un Excel (ruta o dict).
    Devuelve:
        certificados_por_compania: {compañía: [{"filename": str, "content": bytes}, ...]}
        df: DataFrame actualizado con 'certificado' marcado en "si".
    """

    # === 1) Convertimos excel_input a DataFrame ===
    if isinstance(excel_input, (str, os.PathLike)):
        df = pd.read_excel(excel_input)
    elif isinstance(excel_input, dict) and "columns" in excel_input and "data" in excel_input:
        df = pd.DataFrame(excel_input["data"], columns=excel_input["columns"])
    else:
        raise ValueError("excel_input debe ser ruta a Excel o un dict con columnas y data")

    # Validar columna 'certificado'
    if "certificado" not in df.columns:
        raise ValueError("El archivo no contiene la columna 'certificado'")

    # Si no hay pendientes, retornar vacío
    if not (df["certificado"].astype(str).str.lower() == "no").any():
        return {}, df

    plantilla = DocxTemplate("plantilla.docx")
    certificados_por_compania = defaultdict(list)

    # === 2) Procesar cada fila pendiente ===
    for idx, row in df.iterrows():
        if str(row["certificado"]).lower() == "no":
            contexto = {
                "NOMBRE": row.get("nombre", ""),
                "CEDULA": row.get("cedula", ""),
                "FECHA": row["fecha"].strftime("%d/%m/%Y") if "fecha" in row and not pd.isna(row["fecha"]) else "",
                "COMPANIA": row.get("compañia", "Desconocida")
            }

            nombre_base = f"certificado_{contexto['NOMBRE'].replace(' ', '_')}.pdf"

            # Crear DOCX temporal
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
                plantilla.render(contexto)
                plantilla.save(tmp_docx.name)

                # Crear PDF temporal
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                    if ON_WINDOWS:
                        # docx2pdf
                        await asyncio.to_thread(convert, tmp_docx.name, tmp_pdf.name)
                    else:
                        # LibreOffice
                        process = await asyncio.create_subprocess_exec(
                            "soffice", "--headless", "--convert-to", "pdf",
                            "--outdir", os.path.dirname(tmp_pdf.name), tmp_docx.name
                        )
                        await process.communicate()

                    # Leer PDF en memoria
                    with open(tmp_pdf.name, "rb") as f:
                        pdf_bytes = f.read()

                    certificados_por_compania[contexto["COMPANIA"]].append({
                        "filename": nombre_base,
                        "content": pdf_bytes
                    })

                # Limpieza
                os.remove(tmp_docx.name)
                os.remove(tmp_pdf.name)

            # === 3) Marcar certificado como generado ===
            df.at[idx, "certificado"] = "si"

    return certificados_por_compania, df



"""
GenCer (Repositorio GitHub):
-agents: data_campus_agent.py
-auth: auth_manager.py
-core: certificados.py
-one_drive: OD_manager.py
-statics: logo, estilos, etc.
-templates: index.html, success.html
-tests: test_e2e.py
-venv
-certificados.py
-onedrive_client.py
-main.py
-plantilla.docx
-readme.md
-.env
-requirements.txt


"""



