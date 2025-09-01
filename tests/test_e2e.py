import os
import io
import logging
import traceback
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple

import pandas as pd

from agents.datacampus_agent import DatacampusAgent
from core.certificados import generar_certificados_desde_excel  # async: devuelve (certificados_por_compania, df_actualizado)


# ==============================================================
# Config de logging
# ==============================================================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('certificados_log.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class CertificadosProcessor:
    """
    Orquestador del flujo E2E. Mi objetivo: no dejar basura local.
    - Bajo el Excel desde OneDrive (vía mi API)
    - Genero TODOS los PDFs en memoria (bytes)
    - Subo cada PDF directo a OneDrive, organizado por compañía
    - Actualizo el Excel en memoria y lo sobrescribo en OneDrive

    Si algo falla aquí, el problema probablemente está en el agent
    (headers, token o endpoints), no en la lógica de negocio.
    """

    def __init__(self):
        self.agent = DatacampusAgent()

        # Carpeta en OneDrive donde viven los certificados por compañía.
        # Si cambia, lo apunto aquí. La dejo como env para no re-editar código.
        self.certificados_folder_id = os.getenv(
            "CERTIFICADOS_FOLDER_ID",
            "01WIY7HEN2WE6VD2WDPFG3UMFYFXUVC25K"  # valor por defecto actual
        )

        # Excel a procesar: necesito su fileId para poder reemplazarlo (ideal)
        self.excel_file_id = os.getenv("EXCEL_FILE_ID")  # obligatorio
        # Nombre visible del Excel (por si el agent no soporta update por id y toca subir por nombre)
        self.excel_file_name = os.getenv("EXCEL_FILE_NAME", "datos.xlsx")
        # Por si necesito subir por carpeta en vez de update por id
        self.excel_parent_folder_id = os.getenv("EXCEL_PARENT_FOLDER_ID", self.certificados_folder_id)

    # ----------------------------------------------------------
    # API público
    # ----------------------------------------------------------
    def ejecutar_flujo_completo(self) -> bool:
        logger.info("=== INICIANDO FLUJO DE CERTIFICADOS ===")
        try:
            # 1) Autenticación
            if not self._autenticar():
                return False

            # 2) Traer Excel desde OneDrive
            excel_dict = self._obtener_excel_como_dict()
            if not excel_dict:
                return False

            # 3) Generar certificados en memoria + DF actualizado
            certificados_por_compania, df_actualizado = self._generar(cert_input=excel_dict)

            # Si no hay nada para generar, igual actualizo Excel si fuera necesario y salgo
            if not certificados_por_compania:
                logger.info("No hay registros pendientes de certificar")
                # Aseguro consistencia: si el DF viene igual, no pasa nada
                self._subir_excel_actualizado(df_actualizado)
                logger.info("=== FLUJO COMPLETADO (sin pendientes) ===")
                return True

            # 4) Subir PDFs por compañía a OneDrive
            if not self._subir_certificados_por_empresa(certificados_por_compania):
                return False

            # 5) Subir Excel actualizado a OneDrive (reemplazo)
            if not self._subir_excel_actualizado(df_actualizado):
                return False

            logger.info("=== FLUJO COMPLETADO EXITOSAMENTE ===")
            return True

        except Exception as e:
            logger.error(f"Error en el flujo principal: {str(e)}")
            logger.error(traceback.format_exc())
            return False

    # ----------------------------------------------------------
    # Pasos internos
    # ----------------------------------------------------------
    def _autenticar(self) -> bool:
        """Me autentico una única vez. Si aquí revienta, ni intento el resto."""
        logger.info("1) Autenticando con Microsoft Graph...")
        try:
            if self.agent.autenticar():
                logger.info(" Autenticación exitosa")
                return True
            logger.error(" Error en autenticación")
            return False
        except Exception as e:
            logger.error(f" Error durante autenticación: {str(e)}")
            return False

    def _obtener_excel_como_dict(self) -> Optional[Dict[str, Any]]:
        """
        Descargo el Excel desde OneDrive vía mi API y lo convierto a un formato
        amigable para generar_certificados_desde_excel:
            {
              'columns': [...],
              'data': [[fila en orden de columns], ...]
            }
        """
        logger.info("2) Descargando Excel desde OneDrive...")
        try:
            if not self.excel_file_id:
                logger.error(" EXCEL_FILE_ID no está configurado en .env")
                return None

            # El agent debe devolverme un JSON amigable. Yo lo normalizo a DataFrame.
            excel_payload = self.agent.obtener_excel_como_json(self.excel_file_id)
            # Acepto dos formatos: {'data': [dict, dict, ...]} o {'columns': [...], 'data': [[...], ...]}
            df: pd.DataFrame
            if isinstance(excel_payload, dict) and 'data' in excel_payload and isinstance(excel_payload['data'], list):
                # Caso 1: lista de dicts
                if excel_payload.get('columns') is None or isinstance(excel_payload['data'][0], dict):
                    df = pd.DataFrame(excel_payload['data'])
                else:
                    # Caso 2: columnas + data tabular
                    df = pd.DataFrame(excel_payload['data'], columns=excel_payload['columns'])
            else:
                logger.error(" Respuesta del Excel no tiene el formato esperado")
                return None

            # Limpieza mínima: me aseguro de que existan columnas base
            for col in ("nombre", "cedula", "compañia", "certificado"):
                if col not in df.columns:
                    df[col] = ""

            # Log útil
            logger.info(f" Excel leído con {len(df)} filas")

            # Lo convierto al contrato esperado por el generador
            columns = df.columns.tolist()
            data = df.fillna("").astype(str).values.tolist()
            return {"columns": columns, "data": data}

        except Exception as e:
            logger.error(f" Error al obtener Excel: {str(e)}")
            return None

    def _generar(self, cert_input: Dict[str, Any]) -> Tuple[Dict[str, List[Dict[str, Any]]], pd.DataFrame]:
        """
        Llamo a la función async de generación usando asyncio.run porque este
        orquestador es síncrono. La función devuelve:
          - certificados_por_compania: {compañía: [{filename, content(bytes)}]}
          - df_actualizado: DataFrame con 'certificado' marcado en 'si' para los procesados
        """
        logger.info("3) Generando certificados PDF en memoria...")
        import asyncio
        try:
            certificados_por_compania, df_actualizado = asyncio.run(
                generar_certificados_desde_excel(cert_input)
            )
            total = sum(len(v) for v in certificados_por_compania.values())
            logger.info(f" Generados {total} certificados en memoria")
            return certificados_por_compania, df_actualizado
        except Exception as e:
            logger.error(f" Error al generar certificados: {str(e)}")
            raise

    def _subir_certificados_por_empresa(self, certificados: Dict[str, List[Dict[str, Any]]]) -> bool:
        """
        Subo cada PDF por compañía. No escribo nada a disco: el agent debe aceptar
        file-like o bytes. Si no, lo adapto aquí con BytesIO.
        """
        logger.info("4) Subiendo certificados por empresa a OneDrive...")
        try:
            for empresa, archivos in certificados.items():
                empresa_norm = self._normalizar_nombre_carpeta(empresa)
                logger.info(f"  Procesando empresa: {empresa_norm}")

                folder_empresa_id = self._crear_carpeta_empresa(empresa_norm)
                if not folder_empresa_id:
                    # Si no pude crear/ubicar la carpeta, sigo con las demás
                    continue

                for item in archivos:
                    filename: str = item["filename"]
                    content: bytes = item["content"]

                    # Armo un stream en memoria. Algunos endpoints requieren .name
                    bio = io.BytesIO(content)
                    setattr(bio, 'name', filename)

                    ok = False
                    # Reuso el contrato existente del agent si ya tengo subir_pdf(fileobj, folder_id, filename)
                    if hasattr(self.agent, 'subir_pdf'):
                        try:
                            ok = self.agent.subir_pdf(bio, folder_id=folder_empresa_id, filename=filename)
                        except TypeError:
                            # Por si la firma difiere, caigo a bytes
                            ok = False

                    if not ok and hasattr(self.agent, 'subir_pdf_bytes'):
                        ok = self.agent.subir_pdf_bytes(content, folder_id=folder_empresa_id, filename=filename)

                    if not ok and hasattr(self.agent, 'upload_file'):
                        # Plan C: API genérica
                        ok = self.agent.upload_file(content, path=f"/certificados/{empresa_norm}/{filename}")

                    if ok:
                        logger.info(f"     Subido: {filename}")
                    else:
                        logger.warning(f"     Error subiendo: {filename}")

            return True
        except Exception as e:
            logger.error(f" Error al subir certificados: {str(e)}")
            return False

    def _subir_excel_actualizado(self, df_actualizado: pd.DataFrame) -> bool:
        """
        Subo el Excel actualizado. Preferencia: actualizar por fileId para mantener versionado.
        Si el agent no lo soporta, lo subo por nombre a la carpeta definida.
        """
        logger.info("5) Actualizando Excel en OneDrive...")
        try:
            # Serializo a bytes una única vez
            buffer = io.BytesIO()
            df_actualizado.to_excel(buffer, index=False)
            content = buffer.getvalue()

            ok = False
            if self.excel_file_id and hasattr(self.agent, 'actualizar_archivo_por_id'):
                ok = self.agent.actualizar_archivo_por_id(self.excel_file_id, content)

            if not ok and hasattr(self.agent, 'subir_pdf'):
                # Reuso subir_pdf como carga binaria genérica (nombre y folder controlados por env)
                bio = io.BytesIO(content)
                setattr(bio, 'name', self.excel_file_name)
                ok = self.agent.subir_pdf(bio, folder_id=self.excel_parent_folder_id, filename=self.excel_file_name)

            if not ok and hasattr(self.agent, 'crear_reporte'):
                # Último recurso: endpoint que construye Excel desde columnas/series
                # Lo armo con el mismo esquema del DF para no cambiar contratos del agent
                cols = df_actualizado.columns.tolist()
                data = df_actualizado.fillna("").astype(str).values.tolist()
                payload = {col: [row[i] for row in data] for i, col in enumerate(cols)}
                ok = self.agent.crear_reporte(
                    folder_id=self.excel_parent_folder_id,
                    nombre_archivo=self.excel_file_name,
                    datos=payload
                )

            if ok:
                logger.info(" Excel actualizado correctamente")
                return True

            logger.error(" Error al actualizar Excel (ningún endpoint aceptó la subida)")
            return False
        except Exception as e:
            logger.error(f" Error al actualizar Excel: {str(e)}")
            return False

    # ----------------------------------------------------------
    # Utilidades
    # ----------------------------------------------------------
    def _crear_carpeta_empresa(self, empresa: str) -> Optional[str]:
        """Creo (o recupero) la carpeta de la empresa dentro de la raíz de certificados."""
        try:
            folder_id = None
            if hasattr(self.agent, 'crear_carpeta'):
                folder_id = self.agent.crear_carpeta(empresa, parent_folder_id=self.certificados_folder_id)
            elif hasattr(self.agent, 'ensure_folder'):
                folder_id = self.agent.ensure_folder(parent_id=self.certificados_folder_id, name=empresa)

            if folder_id:
                logger.info(f"     Carpeta creada/encontrada: {empresa}")
                return folder_id

            logger.warning(f"     No se pudo crear carpeta: {empresa}")
            return None
        except Exception as e:
            logger.error(f" Error al crear carpeta {empresa}: {str(e)}")
            return None

    @staticmethod
    def _normalizar_nombre_carpeta(nombre: str) -> str:
        """Limpio el nombre para que sea un nombre de carpeta válido en OneDrive."""
        limpio = ''.join(c for c in nombre if c.isalnum() or c in (' ', '-', '_')).strip()
        return limpio or 'Sin_Nombre'


# ==============================================================
# CLI
# ==============================================================

def main():
    processor = CertificadosProcessor()

    print(" Iniciando generación de certificados...")
    print(" Revisa el archivo 'certificados_log.log' para detalles")

    success = processor.ejecutar_flujo_completo()

    if success:
        print("\n Proceso completado exitosamente")
        print(" Los certificados fueron subidos por compañía y el Excel quedó actualizado en OneDrive")
    else:
        print("\n El proceso falló. Revisa el log para más detalles")
        print(" Archivo de log: certificados_log.log")


if __name__ == "__main__":
    main()
