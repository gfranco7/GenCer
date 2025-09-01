# config.py
import os
from dataclasses import dataclass
from dotenv import load_dotenv


load_dotenv()
@dataclass
class CertificadosConfig:
    """Configuración para el sistema de certificados"""
    
    # OneDrive/Microsoft Graph
    excel_file_id: str
    certificados_folder_id: str = "01WIY7HEN2WE6VD2WDPFG3UMFYFXUVC25K"
    
    # Archivos locales
    plantilla_path: str = "plantilla.docx"
    temp_folder: str = "temp_pdfs"
    log_file: str = "certificados_log.log"
    
    # Columnas esperadas en Excel
    columna_certificado: str = "certificado"
    columna_empresa: str = "compañia"
    columna_nombre: str = "nombre"
    columna_cedula: str = "cedula"
    
    # Valores para marcar certificados 
    valor_pendiente: list = None
    valor_completado: str = "si"
    
    def __post_init__(self):
        if self.valor_pendiente is None:
            self.valor_pendiente = ['no', 'n', '0', 'false', '']
    
    @classmethod
    def from_env(cls) -> 'CertificadosConfig':
        """Crea configuración desde variables de entorno"""
        excel_file_id = os.getenv("EXCEL_FILE_ID")
        if not excel_file_id:
            raise ValueError("EXCEL_FILE_ID debe estar definido en .env")
        
        return cls(
            excel_file_id=excel_file_id,
            certificados_folder_id=os.getenv("CERTIFICADOS_FOLDER_ID", cls.certificados_folder_id),
            plantilla_path=os.getenv("PLANTILLA_PATH", cls.plantilla_path),
            temp_folder=os.getenv("TEMP_FOLDER", cls.temp_folder),
            log_file=os.getenv("LOG_FILE", cls.log_file),
            columna_certificado=os.getenv("COLUMNA_CERTIFICADO", cls.columna_certificado),
            columna_empresa=os.getenv("COLUMNA_EMPRESA", cls.columna_empresa),
            columna_nombre=os.getenv("COLUMNA_NOMBRE", cls.columna_nombre),
            columna_cedula=os.getenv("COLUMNA_CEDULA", cls.columna_cedula),
            valor_completado=os.getenv("VALOR_COMPLETADO", cls.valor_completado)
        )
    
    def validar(self) -> bool:
        """Valida que la configuración sea correcta"""
        errores = []
        
        if not os.path.exists(self.plantilla_path):
            errores.append(f"Plantilla no encontrada: {self.plantilla_path}")
        
        if not self.excel_file_id:
            errores.append("ID del archivo Excel no configurado")
        
        if errores:
            print(" Errores de configuración:")
            for error in errores:
                print(f"  - {error}")
            return False
        
        return True