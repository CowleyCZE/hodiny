import dropbox
import os
from dropbox.files import WriteMode
from dropbox.exceptions import ApiError

class DropboxManager:
    def __init__(self, access_token):
        self.dbx = dropbox.Dropbox(access_token)

    def upload_file(self, local_path, dropbox_path):
        with open(local_path, 'rb') as f:
            try:
                self.dbx.files_upload(f.read(), dropbox_path, mode=WriteMode('overwrite'))
                print(f"Soubor {local_path} byl úspěšně nahrán na Dropbox jako {dropbox_path}")
            except ApiError as e:
                print(f'Chyba při nahrávání souboru na Dropbox: {str(e)}')

    def download_file(self, dropbox_path, local_path):
        try:
            _, response = self.dbx.files_download(dropbox_path)
            with open(local_path, 'wb') as f:
                f.write(response.content)
            print(f"Soubor {dropbox_path} byl úspěšně stažen z Dropboxu jako {local_path}")
        except ApiError as e:
            print(f'Chyba při stahování souboru z Dropboxu: {str(e)}')

    def read_json(self, dropbox_path):
        try:
            _, response = self.dbx.files_download(dropbox_path)
            return response.content
        except ApiError as e:
            print(f'Chyba při čtení JSON souboru z Dropboxu: {str(e)}')
            return None

    def write_json(self, dropbox_path, content):
        try:
            self.dbx.files_upload(content.encode(), dropbox_path, mode=WriteMode('overwrite'))
            print(f"JSON soubor byl úspěšně nahrán na Dropbox jako {dropbox_path}")
        except ApiError as e:
            print(f'Chyba při zápisu JSON souboru na Dropbox: {str(e)}')