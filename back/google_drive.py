import os
import io
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from PySide6.QtWidgets import QMessageBox, QInputDialog

from utils.config import SCOPES, TOKEN_FILE, CREDENTIALS_FILE

class GoogleDriveManager:
    def __init__(self, parent_window):
        self.service = None
        self.parent = parent_window

    def authenticate(self) -> bool:
        if self.service:
            QMessageBox.information(self.parent, "Авторизація", "Ви вже увійшли в Google.")
            return True
            
        try:
            creds = self._get_credentials()
            self.service = build('drive', 'v3', credentials=creds)
            QMessageBox.information(self.parent, "Успіх", "Авторизація Google пройшла успішно.")
            return True
        except FileNotFoundError as e:
             QMessageBox.critical(self.parent, "Помилка авторизації", f"{e}")
             return False
        except Exception as e:
            QMessageBox.critical(self.parent, "Помилка авторизації", f"Не вдалося авторизуватися: {e}")
            self.service = None
            return False

    def logout(self) -> None:
        self.service = None
        try:
            if os.path.exists(TOKEN_FILE):
                os.remove(TOKEN_FILE)
            QMessageBox.information(self.parent, "Вихід", "Ви успішно вийшли з акаунту Google.")
        except Exception as e:
            print(f"Не вдалося видалити {TOKEN_FILE}: {e}")
            QMessageBox.warning(self.parent, "Вихід", f"Не вдалося видалити файл сесії: {e}")

    def _get_credentials(self) -> Credentials:
        creds = None
        if os.path.exists(TOKEN_FILE):
            creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
        
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                if not os.path.exists(CREDENTIALS_FILE):
                    raise FileNotFoundError(f"Не знайдено {CREDENTIALS_FILE}. Завантажте його з Google Cloud Console.")
                flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
                creds = flow.run_local_server(port=0)
            with open(TOKEN_FILE, 'w') as token:
                token.write(creds.to_json())
        return creds

    def list_spreadsheets(self) -> tuple:
        if not self.service:
            return None, "Спочатку потрібно авторизуватися."
        
        try:
            query = ("(mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' "
                     "or mimeType='application/vnd.google-apps.spreadsheet')")
            results = self.service.files().list(
                q=query, pageSize=30, fields="files(id, name, mimeType)").execute()
            files = results.get('files', [])
            return files, None
        except HttpError as e:
            return None, f"Помилка Google API: {e}"
        except Exception as e:
            return None, f"Невідома помилка: {e}"


    def download_file(self, file_id: str, mime_type: str) -> tuple:
        if not self.service:
             return None, "Спочатку потрібно авторизуватися."
             
        try:
            request = None
            if mime_type == 'application/vnd.google-apps.spreadsheet':
                request = self.service.files().export_media(
                    fileId=file_id,
                    mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            else:
                request = self.service.files().get_media(fileId=file_id)
            
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while done is False:
                status, done = downloader.next_chunk()
            fh.seek(0)
            return fh, None
        except HttpError as e:
            return None, f"Помилка Google API при завантаженні: {e}"
        except Exception as e:
            return None, f"Невідома помилка при завантаженні: {e}"

    def upload_file(self, file_name: str, file_data_io: io.BytesIO) -> tuple:
         if not self.service:
             return None, None, "Спочатку потрібно авторизуватися."
         
         try:
            file_metadata = {'name': file_name}
            media = MediaIoBaseUpload(
                file_data_io,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                resumable=True
            )
            file = self.service.files().create(
                body=file_metadata, media_body=media, fields='id, webViewLink').execute()
            return file.get('id'), file.get('webViewLink'), None
         except HttpError as e:
            return None, None, f"Помилка Google API при збереженні: {e}"
         except Exception as e:
            return None, None, f"Невідома помилка при збереженні: {e}"