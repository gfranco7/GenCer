
import requests
import pandas as pd
import msal
import os
import dotenv
import webbrowser
import io
import json
from msal import PublicClientApplication, SerializableTokenCache
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass
from auth.auth_manager import AuthManager

# Inicializar y autenticar
dotenv.load_dotenv()

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = [
    'https://graph.microsoft.com/Files.ReadWrite.All',
    'https://graph.microsoft.com/Sites.ReadWrite.All',
    'https://graph.microsoft.com/User.Read'
]

@dataclass
class DriveItem:
    id: str
    name: str
    type: str 
    size: int = 0
    created_datetime: str = ""
    modified_datetime: str = ""

class OneDriveManager:

    def __init__(self, token=None):
        """Inicializar OneDriveManager con o sin token"""
        self.token = token
        self.authenticated = token is not None
        
        # Inicializar atributos que se usan en initialize_datacampus
        self.datacampus_drive_id = None
        self.datacampus_root_id = None
        
        if token:
            self._initialize_with_token(token)
    
    def _initialize_with_token(self, token):
        """Inicialización con token válido"""
        self.token = token
        self.authenticated = True

    def authenticate(self):
        """Método para autenticar después de la inicialización"""
        if not self.token:
            # Obtener token usando AuthManager
            auth = AuthManager()
            self.token = auth.get_token()
            self.authenticated = True
        return True

    def _make_request(self, method: str, url: str, **kwargs) -> requests.Response:
        """Realizar petición HTTP con manejo de errores"""
        if not self.token:
            raise Exception("No hay token de autenticación. Llama a authenticate() primero.")
            
        headers = kwargs.get('headers', {})
        headers['Authorization'] = f"Bearer {self.token['access_token']}"
        kwargs['headers'] = headers
        
        response = requests.request(method, url, **kwargs)
        
        if response.status_code == 401:
            print(" Token expirado, reautenticando...")
            self.authenticate()
            headers['Authorization'] = f"Bearer {self.token['access_token']}"
            response = requests.request(method, url, **kwargs)
            
        return response

    def initialize_datacampus(self) -> Tuple[str, str]:
        """Inicializar y encontrar la carpeta datacampus"""
        if self.datacampus_drive_id and self.datacampus_root_id:
            return self.datacampus_drive_id, self.datacampus_root_id
            
        url = 'https://graph.microsoft.com/v1.0/me/drive/sharedWithMe'
        response = self._make_request('GET', url)

        if response.status_code != 200:
            raise Exception(f"Error al obtener archivos compartidos: {response.status_code} - {response.text}")

        shared_items = response.json().get('value', [])
        
        for item in shared_items:
            if item['name'].lower() == "datacampus":
                remote_item = item.get('remoteItem', {})
                drive_id = remote_item['parentReference']['driveId']
                root_id = remote_item['id']
                
                self.datacampus_drive_id = drive_id
                self.datacampus_root_id = root_id
                
                print(f" Datacampus encontrado - Drive ID: {drive_id[:8]}...")
                return drive_id, root_id

        raise Exception(" No se encontró la carpeta 'datacampus' en elementos compartidos.")

    def list_folder_contents(self, folder_id: Optional[str] = None) -> List[DriveItem]:
        """Listar contenido de una carpeta"""
        if not folder_id:
            folder_id = self.datacampus_root_id
            
        url = f"https://graph.microsoft.com/v1.0/drives/{self.datacampus_drive_id}/items/{folder_id}/children"
        response = self._make_request('GET', url)

        if response.status_code != 200:
            raise Exception(f"Error al listar contenido: {response.status_code} - {response.text}")

        items = []
        for item_data in response.json().get('value', []):
            item_type = 'folder' if 'folder' in item_data else 'file'
            
            items.append(DriveItem(
                id=item_data['id'],
                name=item_data['name'],
                type=item_type,
                size=item_data.get('size', 0),
                created_datetime=item_data.get('createdDateTime', ''),
                modified_datetime=item_data.get('lastModifiedDateTime', '')
            ))
        
        return items

    def find_item_by_name(self, name: str, folder_id: Optional[str] = None) -> Optional[DriveItem]:
        """Buscar un elemento por nombre en una carpeta"""
        items = self.list_folder_contents(folder_id)
        
        for item in items:
            if item.name.lower() == name.lower():
                return item
        
        return None

    def create_excel_file(self, folder_id: str, filename: str, data: Optional[pd.DataFrame] = None) -> Dict:
        """Crear un archivo Excel en una carpeta"""
        if not filename.endswith('.xlsx'):
            filename += '.xlsx'
        
        # Excel por defecto
        if data is None:
            data = pd.DataFrame({
                'Columna1': ['Valor1', 'Valor2', 'Valor3'],
                'Columna2': [10, 20, 30],
                'Columna3': ['A', 'B', 'C']
            })
        
        excel_buffer = io.BytesIO()
        data.to_excel(excel_buffer, index=False, engine='openpyxl')
        excel_buffer.seek(0)

        url = f"https://graph.microsoft.com/v1.0/drives/{self.datacampus_drive_id}/items/{folder_id}:/{filename}:/content"
        headers = {
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }
        
        response = self._make_request('PUT', url, headers=headers, data=excel_buffer)

        if response.status_code in [200, 201]:
            print(f" Archivo '{filename}' creado exitosamente")
            return response.json()
        else:
            raise Exception(f"Error al crear archivo: {response.status_code} - {response.text}")

    def read_excel_file(self, file_id: str) -> pd.DataFrame:
        """Leer un archivo Excel desde OneDrive y retornar DataFrame"""
        url = f"https://graph.microsoft.com/v1.0/drives/{self.datacampus_drive_id}/items/{file_id}/content"
        print(f"DEBUG: Descargando archivo con ID {file_id} desde {url}")

        response = self._make_request('GET', url)

        print("DEBUG: Código de respuesta:", response.status_code)
        print("DEBUG: Content-Type:", response.headers.get("Content-Type"))
        print("DEBUG: Tamaño del archivo descargado:", len(response.content), "bytes")

        if response.status_code != 200:
            raise Exception(f"Error al descargar archivo: {response.status_code} - {response.text}")

        try:
            df = pd.read_excel(io.BytesIO(response.content), engine='openpyxl')
            print(" DEBUG: Excel leído correctamente. Shape:", df.shape)
            return df
        except Exception as e:
            # Guardar copia local para inspección
            with open("debug_excel.xlsx", "wb") as f:
                f.write(response.content)
            print(" DEBUG: Error leyendo Excel con pandas:", str(e))
            print(" DEBUG: Se guardó una copia como debug_excel.xlsx para inspección manual.")
            raise Exception(f"Error al leer archivo Excel: {e}")

    def update_excel_file(self, file_id: str, data: pd.DataFrame) -> Dict:
        """Actualizar un archivo Excel existente"""
        excel_buffer = io.BytesIO()
        data.to_excel(excel_buffer, index=False, engine='openpyxl')
        excel_buffer.seek(0)

        url = f"https://graph.microsoft.com/v1.0/drives/{self.datacampus_drive_id}/items/{file_id}/content"
        headers = {
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }
        
        response = self._make_request('PUT', url, headers=headers, data=excel_buffer)

        if response.status_code == 200:
            print(" Archivo actualizado exitosamente")
            return response.json()
        else:
            raise Exception(f"Error al actualizar archivo: {response.status_code} - {response.text}")

    def delete_item(self, item_id: str) -> None:
        """Eliminar un archivo o carpeta"""
        url = f"https://graph.microsoft.com/v1.0/drives/{self.datacampus_drive_id}/items/{item_id}"
        response = self._make_request('DELETE', url)

        if response.status_code == 204:
            print(" Elemento eliminado exitosamente")
        else:
            raise Exception(f"Error al eliminar elemento: {response.status_code} - {response.text}")

    def create_folder(self, parent_folder_id: str, folder_name: str) -> Dict:
        """Crear una nueva carpeta"""
        url = f"https://graph.microsoft.com/v1.0/drives/{self.datacampus_drive_id}/items/{parent_folder_id}/children"
        
        data = {
            "name": folder_name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "rename"
        }
        
        response = self._make_request('POST', url, json=data)

        if response.status_code == 201:
            print(f" Carpeta '{folder_name}' creada exitosamente")
            return response.json()
        else:
            raise Exception(f"Error al crear carpeta: {response.status_code} - {response.text}")

    def get_item_info(self, item_id: str) -> Dict:
        """Obtener información detallada de un elemento"""
        url = f"https://graph.microsoft.com/v1.0/drives/{self.datacampus_drive_id}/items/{item_id}"
        response = self._make_request('GET', url)

        if response.status_code == 200:
            return response.json()
        else:
            raise Exception(f"Error al obtener información: {response.status_code} - {response.text}")


def encontrar_carpeta_datacampus(token):
    """Función de compatibilidad"""
    manager = OneDriveManager(token)
    return manager.initialize_datacampus()