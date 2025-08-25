# tests/test_e2e_mejorado.py
import os
import logging
import traceback
from pathlib import Path
from typing import Dict, List, Any
from agents.datacampus_agent import DatacampusAgent
from core.certificados import generar_certificados_desde_excel

# Configurar logging
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
    def __init__(self):
        self.agent = DatacampusAgent()
        self.certificados_folder_id = "01WIY7HEN2WE6VD2WDPFG3UMFYFXUVC25K"
        self.pdf_temp_folder = "temp_pdfs"
        
    def ejecutar_flujo_completo(self) -> bool:
        """
        Ejecuta el flujo completo de generación de certificados
        """
        logger.info("=== INICIANDO FLUJO DE CERTIFICADOS ===")
        
        try:
            # 1. Autenticación
            if not self._autenticar():
                return False
                
            # 2. Obtener datos del Excel
            excel_data = self._obtener_datos_excel()
            if not excel_data:
                return False
                
            # 3. Filtrar registros que necesitan certificado
            registros_pendientes = self._filtrar_registros_pendientes(excel_data)
            if not registros_pendientes:
                logger.info("No hay registros pendientes de certificar")
                return True
                
            # 4. Generar certificados PDF
            pdfs_generados = self._generar_certificados_pdf(registros_pendientes)
            if not pdfs_generados:
                return False
                
            # 5. Organizar PDFs por compañía y subir a OneDrive
            if not self._subir_certificados_por_empresa(pdfs_generados, registros_pendientes):
                return False
                
            # 6. Actualizar Excel marcando certificados como generados
            if not self._actualizar_excel_certificados(excel_data, registros_pendientes):
                return False
                
            # 7. Limpiar archivos temporales
            self._limpiar_archivos_temporales()
            
            logger.info("=== FLUJO COMPLETADO EXITOSAMENTE ===")
            return True
            
        except Exception as e:
            logger.error(f"Error en el flujo principal: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def _autenticar(self) -> bool:
        """Autentica con Microsoft Graph API"""
        logger.info("1) Autenticando con Microsoft Graph...")
        try:
            if self.agent.autenticar():
                logger.info(" Autenticación exitosa")
                return True
            else:
                logger.error(" Error en autenticación")
                return False
        except Exception as e:
            logger.error(f" Error durante autenticación: {str(e)}")
            return False
    
    def _obtener_datos_excel(self) -> Dict[str, Any]:
        """Obtiene los datos del Excel desde OneDrive"""
        logger.info("2) Descargando Excel desde OneDrive...")
        try:
            excel_id = os.getenv("EXCEL_FILE_ID")
            if not excel_id:
                logger.error(" EXCEL_FILE_ID no está configurado en .env")
                return None
                
            excel_data = self.agent.obtener_excel_como_json(excel_id)
            if excel_data and excel_data.get('data'):
                logger.info(f" Excel leído con {len(excel_data['data'])} filas")
                return excel_data
            else:
                logger.error(" Error al leer Excel o Excel vacío")
                return None
                
        except Exception as e:
            logger.error(f" Error al obtener Excel: {str(e)}")
            return None
    
    def _filtrar_registros_pendientes(self, excel_data: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Filtra registros que necesitan certificado (certificado='no')"""
        logger.info("3) Filtrando registros pendientes...")
        try:
            columns = excel_data['columns']
            data = excel_data['data']
            
            # Verificar que existe la columna 'certificado'
            if 'certificado' not in columns:
                logger.error(" No se encontró la columna 'certificado' en el Excel")
                return []
            
            cert_index = columns.index('certificado')
            comp_index = columns.index('compañia') if 'compañia' in columns else -1
            
            registros_pendientes = []
            for i, row in enumerate(data):
                if len(row) > cert_index and str(row[cert_index]).lower() in ['no', 'n', '0', 'false']:
                    registro = {}
                    for j, col in enumerate(columns):
                        registro[col] = row[j] if j < len(row) else ''
                    registro['_row_index'] = i
                    registros_pendientes.append(registro)
            
            logger.info(f" Encontrados {len(registros_pendientes)} registros pendientes")
            return registros_pendientes
            
        except Exception as e:
            logger.error(f" Error al filtrar registros: {str(e)}")
            return []
    
    def _generar_certificados_pdf(self, registros: List[Dict[str, Any]]) -> List[str]:
        """Genera los PDFs de certificados"""
        logger.info("4) Generando certificados PDF...")
        try:
            # Crear carpeta temporal
            Path(self.pdf_temp_folder).mkdir(exist_ok=True)
            
            # Convertir registros al formato esperado por generar_certificados_desde_excel
            excel_format = {
                'columns': list(registros[0].keys()) if registros else [],
                'data': [[reg.get(col, '') for col in registros[0].keys()] for reg in registros],
                'info': {'rows': len(registros)}
            }
            
            # Generar PDFs
            generar_certificados_desde_excel(excel_format, self.pdf_temp_folder)
            
            # Listar PDFs generados
            pdfs_generados = [f for f in os.listdir(self.pdf_temp_folder) if f.endswith('.pdf')]
            logger.info(f" Generados {len(pdfs_generados)} certificados PDF")
            return pdfs_generados
            
        except Exception as e:
            logger.error(f" Error al generar PDFs: {str(e)}")
            return []
    
    def _subir_certificados_por_empresa(self, pdfs: List[str], registros: List[Dict[str, Any]]) -> bool:
        """Sube certificados organizados por empresa a OneDrive"""
        logger.info("5) Subiendo certificados por empresa a OneDrive...")
        try:
            # Agrupar por compañía
            empresas = {}
            for registro in registros:
                empresa = registro.get('compañia', 'Sin_Empresa').strip()
                if not empresa:
                    empresa = 'Sin_Empresa'
                if empresa not in empresas:
                    empresas[empresa] = []
                empresas[empresa].append(registro)
            
            # Procesar cada empresa
            for empresa, regs_empresa in empresas.items():
                logger.info(f"  Procesando empresa: {empresa}")
                
                # Crear/obtener carpeta de empresa
                folder_empresa_id = self._crear_carpeta_empresa(empresa)
                if not folder_empresa_id:
                    continue
                
                # Subir PDFs de esta empresa
                for registro in regs_empresa:
                    pdf_filename = self._generar_nombre_pdf(registro)
                    pdf_path = os.path.join(self.pdf_temp_folder, pdf_filename)
                    
                    if os.path.exists(pdf_path):
                        with open(pdf_path, "rb") as f:
                            result = self.agent.subir_pdf(
                                f, 
                                folder_id=folder_empresa_id, 
                                filename=pdf_filename
                            )
                            if result:
                                logger.info(f"     Subido: {pdf_filename}")
                            else:
                                logger.warning(f"     Error subiendo: {pdf_filename}")
                    else:
                        logger.warning(f"     PDF no encontrado: {pdf_filename}")
            
            return True
            
        except Exception as e:
            logger.error(f" Error al subir certificados: {str(e)}")
            return False
    
    def _crear_carpeta_empresa(self, empresa: str) -> str:
        """Crea carpeta de empresa dentro de Certificados"""
        try:
            # Normalizar nombre de empresa para nombre de carpeta
            nombre_carpeta = "".join(c for c in empresa if c.isalnum() or c in (' ', '-', '_')).strip()
            if not nombre_carpeta:
                nombre_carpeta = "Sin_Nombre"
            
            # Crear carpeta (implementar en DatacampusAgent si no existe)
            folder_id = self.agent.crear_carpeta(
                nombre_carpeta, 
                parent_folder_id=self.certificados_folder_id
            )
            
            if folder_id:
                logger.info(f"     Carpeta creada/encontrada: {nombre_carpeta}")
                return folder_id
            else:
                logger.warning(f"     No se pudo crear carpeta: {nombre_carpeta}")
                return None
                
        except Exception as e:
            logger.error(f" Error al crear carpeta {empresa}: {str(e)}")
            return None
    
    def _generar_nombre_pdf(self, registro: Dict[str, Any]) -> str:
        """Genera nombre del archivo PDF basado en los datos del registro"""
        nombre = registro.get('nombre', 'Sin_Nombre').strip()
        cedula = registro.get('cedula', registro.get('identificacion', 'Sin_ID')).strip()
        
        # Limpiar caracteres especiales
        nombre_limpio = "".join(c for c in nombre if c.isalnum() or c in (' ', '-', '_')).strip()
        cedula_limpia = "".join(c for c in str(cedula) if c.isalnum())
        
        return f"Certificado_{nombre_limpio}_{cedula_limpia}.pdf"
    
    def _actualizar_excel_certificados(self, excel_data: Dict[str, Any], registros_procesados: List[Dict[str, Any]]) -> bool:
        """Actualiza el Excel marcando los certificados como generados"""
        logger.info("6) Actualizando Excel...")
        try:
            columns = excel_data['columns']
            data = excel_data['data'].copy()
            
            cert_index = columns.index('certificado')
            
            # Actualizar filas procesadas
            for registro in registros_procesados:
                row_index = registro['_row_index']
                if row_index < len(data):
                    data[row_index][cert_index] = 'si'
            
            # Crear datos para actualizar
            datos_actualizados = {
                col: [row[i] if i < len(row) else '' for row in data] 
                for i, col in enumerate(columns)
            }
            
            # Subir Excel actualizado
            result = self.agent.crear_reporte(
                folder_id=None,  # Misma ubicación del original
                nombre_archivo="excel_actualizado.xlsx",
                datos=datos_actualizados
            )
            
            if result:
                logger.info(" Excel actualizado correctamente")
                return True
            else:
                logger.error(" Error al actualizar Excel")
                return False
                
        except Exception as e:
            logger.error(f" Error al actualizar Excel: {str(e)}")
            return False
    
    def _limpiar_archivos_temporales(self):
        """Limpia archivos temporales"""
        try:
            if os.path.exists(self.pdf_temp_folder):
                import shutil
                shutil.rmtree(self.pdf_temp_folder)
                logger.info("🧹 Archivos temporales eliminados")
        except Exception as e:
            logger.warning(f" Error al limpiar temporales: {str(e)}")


def main():
    """Función principal para ejecutar desde terminal"""
    processor = CertificadosProcessor()
    
    print(" Iniciando generación de certificados...")
    print(" Revisa el archivo 'certificados_log.log' para detalles")
    
    success = processor.ejecutar_flujo_completo()
    
    if success:
        print("\n ¡Proceso completado exitosamente!")
        print(" Los certificados han sido organizados por empresa en OneDrive")
    else:
        print("\n El proceso falló. Revisa el log para más detalles")
        print(" Archivo de log: certificados_log.log")


if __name__ == "__main__":
    main()