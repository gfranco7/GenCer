import requests
import os

class OneDriveClient:
    def __init__(self, access_token):
        self.base_url = "https://graph.microsoft.com/v1.0/me/drive"
        self.headers = {"Authorization": f"Bearer {access_token}"}

    def create_folder(self, parent_id, folder_name):
        url = f"{self.base_url}/items/{parent_id}/children"
        data = {
            "name": folder_name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "rename"
        }
        r = requests.post(url, headers=self.headers, json=data)
        r.raise_for_status()
        return r.json()["id"]

    def upload_file(self, folder_id, file_path):
        file_name = os.path.basename(file_path)
        url = f"{self.base_url}/items/{folder_id}:/{file_name}:/content"
        with open(file_path, "rb") as f:
            r = requests.put(url, headers=self.headers, data=f)
        r.raise_for_status()
        return r.json()
