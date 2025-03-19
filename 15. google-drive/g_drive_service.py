import streamlit as st
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials
from google.oauth2 import service_account

class GoogleDriveService:
    def __init__(self):
        self._SCOPES = ['https://www.googleapis.com/auth/drive']

    def build(self):
        # Usar credenciais diretamente dos secrets do Streamlit
        creds_dict = st.secrets["connections"]["gdrive"]
        
        # Criar credenciais a partir do dicionário
        creds = service_account.Credentials.from_service_account_info(
            creds_dict, scopes=self._SCOPES
        )
        
        # Construir o serviço
        service = build('drive', 'v3', credentials=creds)
        
        return service