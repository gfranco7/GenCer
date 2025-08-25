# datacampus_agent_metodos_adicionales.py
# Métodos adicionales que podrían faltarte en DatacampusAgent

import requests
import logging
from typing import Dict, Any, Optional, BinaryIO
import json

logger = logging.getLogger(__name__)

class DatacampusAgentMethods:
    """
    Métodos adicionales para DatacampusAgent que podrías necesitar implementar
    Copia estos métodos a tu clase DatacampusAgent existente
    """
    
    def crear_carpeta(self, nombre_carpeta: str, parent_folder_id: str = None) -> Optional[str]:
        """
        Crea una carpeta en OneDrive o busca si ya existe
        
        Args:
            nombre_carpeta: Nombre de la carpeta a crear
            parent_folder_id: ID de la carpeta padre (None para raíz)
            
        Returns:
            ID de la carpeta creada/encontrada o None si falla
        """
        try:
            # Primero buscar si la carpeta ya existe
            folder_id = self._buscar_carpeta(nombre_carpeta, parent_folder_id)
            if folder_id:
                logger.info(f"Carpeta '{nombre_carpeta}' ya existe")
                return folder_id
            
            # Si no existe, crear la carpeta
            headers = {
                'Authorization': f'Bearer {self.token}',
                'Content-Type': 'application/json'
            }
            
            # Determinar endpoint
            if parent_folder_id:
                url = f"https://graph.microsoft.com/v1.0/me/drive/items/{parent_folder_id}/children"
            else:
                url = "https://graph.microsoft.com/v1.0/me/drive/root/children"
            
            data = {
                "name": nombre_carpeta,
                "folder": {},
                "@microsoft.graph.conflictBehavior": "rename"
            }
            
            response = requests.post(url, headers=headers, json=data)
            
            if response.status_code == 201:
                folder_info = response.json()
                folder_id = folder_info.get('id')
                logger.info(f"Carpeta '{nombre_carpeta}' creada exitosamente")
                return folder_id
            else:
                logger.error(f"Error al crear carpeta: {response.status_code} - {response.text}")
                return None
                
        except Exception as e:
            logger.error(f"Excepción al crear carpeta '{nombre_carpeta}': {str(e)}")
            return None
    
    def _buscar_carpeta(self, nombre_carpeta: str, parent_folder_id: str = None) -> Optional[str]:
        """
        Busca una carpeta por nombre en una ubicación específica
        
        Args:
            nombre_carpeta: Nombre de la carpeta a buscar
            parent_folder_id: ID de la carpeta padre donde buscar
            
        Returns:
            ID de la carpeta si la encuentra, None si no existe
        """
        try:
            headers = {
                'Authorization': f'Bearer {self.token}'
            }
            
            # Determinar endpoint para listar contenido
            if parent_folder_id:
                url = f"https://graph.microsoft.com/v1.0/me/drive/items/{parent_folder_id}/children"
            else:
                url = "https://graph.microsoft.com/v1.0/me/drive/root/children"
            
            response = requests.get(url, headers=headers)
            
            if response.status_code == 200:
                items = response.json().get('value', [])
                
                # Buscar carpeta por nombre
                for item in items:
                    if (item.get('name', '').lower() == nombre_carpeta.lower() and 
                        'folder' in item):
                        return item.get('id')
                
                return None
            else:
                logger.warning(f"Error al buscar carpeta: {response.status_code}")
                return None
                
        except Exception as e:
            logger.error(f"Error al buscar carpeta '{nombre_carpeta}': {str(e)}")
            return None
    
    def subir_pdf(self, archivo: BinaryIO, folder_id: str = None, filename: str = None) -> bool:
        """
        Sube un archivo PDF a OneDrive
        
        Args:
            archivo: Objeto de archivo binario (file handle)
            folder_id: ID de la carpeta destino (None para raíz)
            filename: Nombre del archivo
            
        Returns:
            True si se subió exitosamente, False si falló
        """
        try:
            if not filename:
                filename = "certificado.pdf"
            
            headers = {
                'Authorization': f'Bearer {self.token}',
                'Content-Type': 'application/pdf'
            }
            
            # Determinar endpoint
            if folder_id:
                url = f"https://graph.microsoft.com/v1.0/me/drive/items/{folder_id}:/{filename}:/content"
            else:
                url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{filename}:/content"
            
            # Leer contenido del archivo
            archivo.seek(0)
            file_content = archivo.read()
            
            response = requests.put(url, headers=headers, data=file_content)
            
            if response.status_code in [200, 201]:
                logger.info(f"Archivo '{filename}' subido exitosamente")
                return True
            else:
                logger.error(f"Error al subir '{filename}': {response.status_code} - {response.text}")
                return False
                
        except Exception as e:
            logger.error(f"Excepción al subir archivo '{filename}': {str(e)}")
            return False
    
    def obtener_excel_como_json(self, file_id: str) -> Optional[Dict[str, Any]]:
        """
        Descarga un archivo Excel de OneDrive y lo convierte a JSON
        
        Args:
            file_id: ID del archivo Excel en OneDrive
            
        Returns:
            Dict con estructura: {
                'columns': [lista de nombres de columnas],
                'data': [lista de filas como listas],
                'info': {'rows': número_de_filas}
            }
        """
        try:
            # Descargar archivo
            excel_content = self._descargar_archivo(file_id)
            if not excel_content:
                return None
            
            # Procesar con pandas/openpyxl
            import pandas as pd
            from io import BytesIO
            
            # Leer Excel
            df = pd.read_excel(BytesIO(excel_content))
            
            # Convertir a formato esperado
            result = {
                'columns': df.columns.tolist(),
                'data': df.values.tolist(),
                'info': {
                    'rows': len(df),
                    'columns': len(df.columns)
                }
            }
            
            logger.info(f"Excel procesado: {result['info']['rows']} filas, {result['info']['columns']} columnas")
            return result
            
        except Exception as e:
            logger.error(f"Error al procesar Excel: {str(e)}")
            return None
    
    def _descargar_archivo(self, file_id: str) -> Optional[bytes]:
        """
        Descarga contenido de un archivo por su ID
        
        Args:
            file_id: ID del archivo en OneDrive
            
        Returns:
            Contenido del archivo como bytes, None si falla
        """
        try:
            headers = {
                'Authorization': f'Bearer {self.token}'
            }
            
            url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/content"
            response = requests.get(url, headers=headers)
            
            if response.status_code == 200:
                return response.content
            else:
                logger.error(f"Error al descargar archivo: {response.status_code}")
                return None
                
        except Exception as e:
            logger.error(f"Error al descargar archivo: {str(e)}")
            return None
    
    def crear_reporte(self, folder_id: str = None, nombre_archivo: str = "reporte.xlsx", 
                     datos: Dict[str, list] = None) -> bool:
        """
        Crea y sube un archivo Excel con datos
        
        Args:
            folder_id: ID de la carpeta destino
            nombre_archivo: Nombre del archivo Excel
            datos: Dict donde keys son nombres de columnas y values son listas de datos
            
        Returns:
            True si se creó exitosamente, False si falló
        """
        try:
            if not datos:
                logger.error("No hay datos para crear el reporte")
                return False
            
            import pandas as pd
            from io import BytesIO
            
            # Crear DataFrame
            df = pd.DataFrame(datos)
            
            # Guardar en BytesIO
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Datos')
            
            excel_buffer.seek(0)
            
            # Subir archivo
            headers = {
                'Authorization': f'Bearer {self.token}',
                'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            }
            
            if folder_id:
                url = f"https://graph.microsoft.com/v1.0/me/drive/items/{folder_id}:/{nombre_archivo}:/content"
            else:
                url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{nombre_archivo}:/content"
            
            response = requests.put(url, headers=headers, data=excel_buffer.getvalue())
            
            if response.status_code in [200, 201]:
                logger.info(f"Reporte '{nombre_archivo}' creado exitosamente")
                return True
            else:
                logger.error(f"Error al crear reporte: {response.status_code} - {response.text}")
                return False
                
        except Exception as e:
            logger.error(f"Error al crear reporte: {str(e)}")
            return False
    
    def autenticar(self) -> bool:
        """
        Método base de autenticación (debes implementar según tu auth_manager)
        
        Returns:
            True si la autenticación fue exitosa
        """
        try:
            # Aquí deberías usar tu AuthManager existente
            # Este es solo un ejemplo de estructura
            
            from auth.auth_manager import AuthManager
            auth_manager = AuthManager()
            
            self.token = auth_manager.obtener_token()
            
            if self.token:
                logger.info("Autenticación exitosa")
                return True
            else:
                logger.error("Error en autenticación - token no obtenido")
                return False
                
        except Exception as e:
            logger.error(f"Error durante autenticación: {str(e)}")
            return False
    
    def validar_token(self) -> bool:
        """
        Valida si el token actual sigue siendo válido
        
        Returns:
            True si el token es válido
        """
        try:
            headers = {
                'Authorization': f'Bearer {self.token}'
            }
            
            # Hacer una llamada simple para validar el token
            url = "https://graph.microsoft.com/v1.0/me"
            response = requests.get(url, headers=headers)
            
            return response.status_code == 200
            
        except Exception as e:
            logger.error(f"Error al validar token: {str(e)}")
            return False

# Ejemplo de uso de estos métodos en tu DatacampusAgent existente:
"""
class DatacampusAgent(DatacampusAgentMethods):
    def __init__(self):
        self.token = None
        # ... resto de tu inicialización
"""