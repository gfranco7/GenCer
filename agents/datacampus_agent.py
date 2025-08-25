# datacampus_agent.py
import requests
from typing import Optional, Dict, Any
import json


class DatacampusAgent:
    def __init__(self, base_url: str = "http://localhost:8000"):
        self.base_url = base_url
        self.token_ok = False
        self.session = requests.Session() 

    def autenticar(self) -> bool:
        """Realiza autenticación inicial"""
        try:
            headers = {
                'Content-Type': 'application/json',
                'Accept': 'application/json'
            }
            response = self.session.post(
                f"{self.base_url}/auth/login", 
                headers=headers,
                json={} 
            )
            
            print(f"Status code: {response.status_code}")
            print(f"Response headers: {dict(response.headers)}")
            print(f"Response content: {response.text}")
            
            response.raise_for_status()
            self.token_ok = True
            return True
            
        except requests.exceptions.HTTPError as e:
            print(f"Error HTTP de autenticación: {e}")
            print(f"Response content: {e.response.text if e.response else 'No response'}")
            return False
        except Exception as e:
            print(f"Error de autenticación: {e}")
            return False
        
    def verificar_autenticacion(self) -> bool:
        """Verifica si la autenticación sigue válida"""
        try:
            response = self.session.get(f"{self.base_url}/auth/status")
            response.raise_for_status()
            return True
        except:
            return False

    def listar_contenido(self, folder_id: Optional[str] = None) -> Optional[Dict[str, Any]]:
        """Lista contenido de carpeta"""
        if not self.token_ok:
            print("Autenticacion fallida")
            return None
            
        try:
            params = {"folder_id": folder_id} if folder_id else {}
            response = self.session.get(f"{self.base_url}/folders", params=params)
            response.raise_for_status()
            return response.json()
        except Exception as e:
            print(f"Error al listar contenido: {e}")
            return None

    def obtener_excel_como_json(self, file_id: str) -> Optional[Dict[str, Any]]:
        """Obtiene el contenido de un archivo Excel como JSON"""
        if not self.token_ok:
            print("Autenticacion fallida")
            return None
            
        try:
            response = self.session.get(f"{self.base_url}/files/{file_id}/content")
            response.raise_for_status()
            return response.json()
        except Exception as e:
            print(f"Error al obtener contenido del archivo: {e}")
            return None

    def crear_reporte(self, folder_id: str, nombre_archivo: str, datos: Dict[str, list]) -> Optional[str]:
        """Crea un archivo Excel en la carpeta especificada con datos dados"""
        if not self.token_ok:
            print("Autenticacion fallida")
            return None
            
        try:
            payload = {
                "filename": nombre_archivo,
                "folder_id": folder_id,
                "data": datos
            }
            response = self.session.post(f"{self.base_url}/files/excel", json=payload)
            response.raise_for_status()
            result = response.json()
            print(f"Archivo creado: {result.get('file_id')}")
            return result.get("file_id")
        
        except Exception as e:
            print(f"Error al crear archivo: {e}")
            return None

    def eliminar_elemento(self, item_id: str) -> bool:  
        """Elimina un archivo o carpeta"""
        if not self.token_ok:
            print("Autenticacion fallida")
            return False
            
        try:
            response = self.session.delete(f"{self.base_url}/items/{item_id}")
            response.raise_for_status()
            print("Elemento eliminado correctamente")
            return True 
        except Exception as e:
            print(f"Error al eliminar elemento: {e}")
            return False

    def debug_servidor(self):
        """Método para diagnosticar problemas de conexión"""
        try:
            response = self.session.get(f"{self.base_url}/")
            print(f"Servidor accesible: {response.status_code}")
            
            response = self.session.get(f"{self.base_url}/docs")
            print(f"Documentación disponible: {response.status_code}")
            
        except Exception as e:
            print(f"Error al conectar con servidor: {e}")