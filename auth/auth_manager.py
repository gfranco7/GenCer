# auth_manager.py
import os
import json
import webbrowser
from msal import PublicClientApplication, SerializableTokenCache
from typing import Optional
from dotenv import load_dotenv

load_dotenv()

# Configuraciones desde .env
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = [
    'https://graph.microsoft.com/Files.ReadWrite.All',
    'https://graph.microsoft.com/Sites.ReadWrite.All',
    'https://graph.microsoft.com/User.Read'
]

class AuthManager:
    def __init__(self, cache_path: str = 'token_cache.bin'):
        self.cache_path = cache_path
        self.cache = SerializableTokenCache()
        self.token = None

        self._load_cache()

        self.app = PublicClientApplication(
            client_id=CLIENT_ID,
            authority=AUTHORITY,
            token_cache=self.cache
        )

    def _load_cache(self):
        if os.path.exists(self.cache_path):
            with open(self.cache_path, 'rb') as f:
                self.cache.deserialize(f.read())

    def _save_cache(self):
        if self.cache.has_state_changed:
            with open(self.cache_path, 'wb') as f:
                f.write(self.cache.serialize().encode("utf-8"))

    def get_token(self, force_auth: bool = False) -> dict:
        if not force_auth:
            accounts = self.app.get_accounts()
            if accounts:
                result = self.app.acquire_token_silent(SCOPES, account=accounts[0])
                if result and "access_token" in result:
                    print("Token obtenido silenciosamente.")
                    self.token = result
                    return result

        print(" Iniciando autenticación interactiva...")
        flow = self.app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise Exception("Error al iniciar flujo de autenticación")

        print(f" Código: {flow['user_code']} | Visita: {flow['verification_uri']}")
        webbrowser.open(flow["verification_uri"])

        result = self.app.acquire_token_by_device_flow(flow)

        if "access_token" in result:
            print("Autenticación exitosa.")
            self.token = result
            self._save_cache()
            return result
        else:
            raise Exception(f" Error en autenticación: {result.get('error_description', 'Desconocido')}")
