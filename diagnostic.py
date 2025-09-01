#!/usr/bin/env python3
"""
diagnostico.py - Herramienta de diagnóstico para problemas de autenticación y permisos

Uso: python diagnostic.py
"""

import os
import requests
import logging
from agents.datacampus_agent import DatacampusAgent
from dotenv import load_dotenv
load_dotenv()
# Configurar logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class DiagnosticTool:
    def __init__(self):
        self.agent = DatacampusAgent()
        
    def ejecutar_diagnostico_completo(self):
        """Ejecuta un diagnóstico completo del sistema"""
        print("🔍 DIAGNÓSTICO DEL SISTEMA DE CERTIFICADOS")
        print("=" * 50)
        
        # 1. Verificar variables de entorno
        self.verificar_variables_entorno()
        
        # 2. Verificar autenticación
        if not self.verificar_autenticacion():
            return
        
        # 3. Verificar permisos del token
        self.verificar_permisos_token()
        
        # 4. Verificar acceso al archivo Excel
        self.verificar_acceso_excel()
        
        # 5. Verificar carpeta de certificados
        self.verificar_carpeta_certificados()
        
        print("\n✅ Diagnóstico completado. Revisa los resultados arriba.")
    
    def verificar_variables_entorno(self):
        """Verifica que todas las variables de entorno estén configuradas"""
        print("\n1️⃣ VERIFICANDO VARIABLES DE ENTORNO...")
        
        variables_requeridas = [
            'CLIENT_ID',
            'CLIENT_SECRET', 
            'TENANT_ID',
            'EXCEL_FILE_ID'
        ]
        
        variables_opcionales = [
            'CERTIFICADOS_FOLDER_ID'
        ]
        
        todas_ok = True
        
        for var in variables_requeridas:
            valor = os.getenv(var)
            if valor:
                print(f"   ✅ {var}: {'*' * min(len(valor), 20)}...")
            else:
                print(f"   ❌ {var}: NO CONFIGURADA")
                todas_ok = False
        
        for var in variables_opcionales:
            valor = os.getenv(var)
            if valor:
                print(f"   📋 {var}: {'*' * min(len(valor), 20)}...")
            else:
                print(f"   📋 {var}: usando valor por defecto")
        
        if todas_ok:
            print("   ✅ Todas las variables requeridas están configuradas")
        else:
            print("   ❌ Faltan variables requeridas en el archivo .env")
        
        return todas_ok
    
    def verificar_autenticacion(self):
        """Verifica que la autenticación funcione correctamente"""
        print("\n2️⃣ VERIFICANDO AUTENTICACIÓN...")
        
        try:
            if self.agent.autenticar():
                print("   ✅ Autenticación exitosa")
                
                # Verificar que el token sea válido
                if hasattr(self.agent, 'token') and self.agent.token:
                    print(f"   📋 Token obtenido (longitud: {len(self.agent.token)})")
                    return True
                else:
                    print("   ❌ Token no disponible después de autenticación")
                    return False
            else:
                print("   ❌ Error en autenticación")
                return False
                
        except Exception as e:
            print(f"   ❌ Excepción durante autenticación: {str(e)}")
            return False
    
    def verificar_permisos_token(self):
        """Verifica los permisos del token haciendo llamadas de prueba"""
        print("\n3️⃣ VERIFICANDO PERMISOS DEL TOKEN...")
        
        if not hasattr(self.agent, 'token') or not self.agent.token:
            print("   ❌ No hay token disponible para verificar")
            return
        
        headers = {
            'Authorization': f'Bearer {self.agent.token}'
        }
        
        # Test 1: Información del usuario
        try:
            response = requests.get('https://graph.microsoft.com/v1.0/me', headers=headers)
            if response.status_code == 200:
                user_info = response.json()
                print(f"   ✅ Acceso a perfil de usuario: {user_info.get('displayName', 'N/A')}")
            else:
                print(f"   ❌ Error al acceder perfil de usuario: {response.status_code}")
                print(f"      Response: {response.text[:200]}...")
        except Exception as e:
            print(f"   ❌ Excepción al verificar perfil: {str(e)}")
        
        # Test 2: Acceso a OneDrive
        try:
            response = requests.get('https://graph.microsoft.com/v1.0/me/drive', headers=headers)
            if response.status_code == 200:
                drive_info = response.json()
                print(f"   ✅ Acceso a OneDrive: {drive_info.get('driveType', 'N/A')}")
            else:
                print(f"   ❌ Error al acceder OneDrive: {response.status_code}")
                print(f"      Response: {response.text[:200]}...")
        except Exception as e:
            print(f"   ❌ Excepción al verificar OneDrive: {str(e)}")
        
        # Test 3: Listar archivos en raíz
        try:
            response = requests.get('https://graph.microsoft.com/v1.0/me/drive/root/children', headers=headers)
            if response.status_code == 200:
                items = response.json().get('value', [])
                print(f"   ✅ Acceso a archivos raíz: {len(items)} elementos encontrados")
            else:
                print(f"   ❌ Error al listar archivos raíz: {response.status_code}")
                print(f"      Response: {response.text[:200]}...")
        except Exception as e:
            print(f"   ❌ Excepción al listar archivos: {str(e)}")
    
    def verificar_acceso_excel(self):
        """Verifica el acceso específico al archivo Excel"""
        print("\n4️⃣ VERIFICANDO ACCESO AL ARCHIVO EXCEL...")
        
        excel_file_id = os.getenv('EXCEL_FILE_ID')
        if not excel_file_id:
            print("   ❌ EXCEL_FILE_ID no configurado")
            return
        
        print(f"   📋 Excel File ID: {excel_file_id[:30]}...")
        
        if not hasattr(self.agent, 'token') or not self.agent.token:
            print("   ❌ No hay token disponible")
            return
        
        headers = {
            'Authorization': f'Bearer {self.agent.token}'
        }
        
        # Test 1: Obtener información del archivo
        try:
            url = f'https://graph.microsoft.com/v1.0/me/drive/items/{excel_file_id}'
            response = requests.get(url, headers=headers)
            
            if response.status_code == 200:
                file_info = response.json()
                print(f"   ✅ Archivo encontrado: {file_info.get('name', 'N/A')}")
                print(f"   📋 Tamaño: {file_info.get('size', 'N/A')} bytes")
                print(f"   📋 Última modificación: {file_info.get('lastModifiedDateTime', 'N/A')}")
                
                # Verificar si es un archivo Excel
                mime_type = file_info.get('file', {}).get('mimeType', '')
                if 'excel' in mime_type or 'spreadsheet' in mime_type:
                    print(f"   ✅ Tipo de archivo correcto: Excel")
                else:
                    print(f"   ⚠️  Tipo de archivo: {mime_type}")
                    
            elif response.status_code == 404:
                print("   ❌ Archivo no encontrado (404)")
                print("   💡 Verifica que el EXCEL_FILE_ID sea correcto")
                
            elif response.status_code == 403:
                print("   ❌ Sin permisos para acceder al archivo (403)")
                print("   💡 Verifica los permisos de la aplicación en Azure")
                
            elif response.status_code == 401:
                print("   ❌ Token no autorizado para este archivo (401)")
                print("   💡 El token puede haber expirado o no tener permisos suficientes")
                
            else:
                print(f"   ❌ Error al acceder al archivo: {response.status_code}")
                print(f"      Response: {response.text[:300]}...")
                
        except Exception as e:
            print(f"   ❌ Excepción al verificar archivo Excel: {str(e)}")
        
        # Test 2: Intentar descargar el contenido
        try:
            url = f'https://graph.microsoft.com/v1.0/me/drive/items/{excel_file_id}/content'
            response = requests.get(url, headers=headers)
            
            if response.status_code == 200:
                content_length = len(response.content)
                print(f"   ✅ Contenido descargable: {content_length} bytes")
                
                # Verificar si parece ser un archivo Excel válido
                if response.content.startswith(b'PK'):
                    print("   ✅ Formato de archivo Excel válido")
                else:
                    print("   ⚠️  El contenido puede no ser un archivo Excel válido")
                    
            elif response.status_code == 401:
                print("   ❌ Error 401 al descargar contenido (este es tu problema principal)")
                print("   💡 Posibles causas:")
                print("      - Token expirado")
                print("      - Permisos insuficientes en la aplicación Azure")
                print("      - El archivo está en una ubicación restringida")
                
            else:
                print(f"   ❌ Error al descargar contenido: {response.status_code}")
                
        except Exception as e:
            print(f"   ❌ Excepción al descargar contenido: {str(e)}")
    
    def verificar_carpeta_certificados(self):
        """Verifica el acceso a la carpeta de certificados"""
        print("\n5️⃣ VERIFICANDO CARPETA DE CERTIFICADOS...")
        
        folder_id = os.getenv('CERTIFICADOS_FOLDER_ID', '01WIY7HEN2WE6VD2WDPFG3UMFYFXUVC25K')
        print(f"   📋 Carpeta ID: {folder_id}")
        
        if not hasattr(self.agent, 'token') or not self.agent.token:
            print("   ❌ No hay token disponible")
            return
        
        headers = {
            'Authorization': f'Bearer {self.agent.token}'
        }
        
        try:
            url = f'https://graph.microsoft.com/v1.0/me/drive/items/{folder_id}'
            response = requests.get(url, headers=headers)
            
            if response.status_code == 200:
                folder_info = response.json()
                print(f"   ✅ Carpeta encontrada: {folder_info.get('name', 'N/A')}")
                
                # Listar contenido de la carpeta
                url_children = f'https://graph.microsoft.com/v1.0/me/drive/items/{folder_id}/children'
                response_children = requests.get(url_children, headers=headers)
                
                if response_children.status_code == 200:
                    children = response_children.json().get('value', [])
                    print(f"   📋 Elementos en la carpeta: {len(children)}")
                    
                    # Mostrar algunos elementos
                    for i, child in enumerate(children[:5]):
                        item_type = "📁" if 'folder' in child else "📄"
                        print(f"      {item_type} {child.get('name', 'N/A')}")
                    
                    if len(children) > 5:
                        print(f"      ... y {len(children) - 5} elementos más")
                        
                else:
                    print(f"   ⚠️  No se pudo listar contenido: {response_children.status_code}")
                    
            else:
                print(f"   ❌ Error al acceder carpeta: {response.status_code}")
                
        except Exception as e:
            print(f"   ❌ Excepción al verificar carpeta: {str(e)}")
    
    def sugerir_soluciones(self):
        """Sugiere soluciones basadas en los problemas encontrados"""
        print("\n🔧 SUGERENCIAS DE SOLUCIÓN PARA ERROR 401:")
        print("=" * 50)
        
        print("\n1️⃣ VERIFICAR PERMISOS EN AZURE:")
        print("   - Ve a Azure Portal > App Registrations > Tu App")
        print("   - Sección 'API Permissions'")
        print("   - Asegúrate de tener estos permisos:")
        print("     • Files.ReadWrite.All (Application o Delegated)")
        print("     • Files.Read.All (Application o Delegated)")
        print("     • User.Read (Delegated)")
        print("   - Haz clic en 'Grant admin consent'")
        
        print("\n2️⃣ VERIFICAR CONFIGURACIÓN DEL TOKEN:")
        print("   - El token puede estar expirando")
        print("   - Verifica que uses el tenant correcto")
        print("   - Considera usar refresh tokens")
        
        print("\n3️⃣ VERIFICAR ID DEL ARCHIVO:")
        print("   - El EXCEL_FILE_ID puede ser incorrecto")
        print("   - Prueba obteniendo el ID directamente desde OneDrive")
        print("   - Verifica que el archivo esté en la cuenta correcta")
        
        print("\n4️⃣ VERIFICAR UBICACIÓN DEL ARCHIVO:")
        print("   - Si el archivo está en un SharePoint compartido")
        print("   - Puede necesitar permisos específicos de Sites")
        print("   - Considera mover el archivo a OneDrive personal")


def main():
    print("🚀 Iniciando diagnóstico del sistema...")
    
    diagnostic = DiagnosticTool()
    diagnostic.ejecutar_diagnostico_completo()
    diagnostic.sugerir_soluciones()
    
    print(f"\n📋 Para más detalles, revisa el log completo")


if __name__ == "__main__":
    main()