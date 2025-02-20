import subprocess
import sys
import os
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

def upload_files():
    # Escopos de acesso à API do Drive
    SCOPES = ['https://www.googleapis.com/auth/drive']
    SERVICE_ACCOUNT_FILE = r"C:\Users\leonardo.fragoso\Desktop\Projetos\Depot-Project\gdrive_credentials.json"
    
    # Cria as credenciais com o arquivo de serviço
    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    
    drive_service = build('drive', 'v3', credentials=credentials)
    
    # ID da pasta no Google Drive onde os arquivos serão enviados
    folder_id = "1ROPmQRq9Wy_Ugzi9rZQ2Vnt5mfnSZ0a5"
    
    downloads_folder = os.path.join(os.getcwd(), "downloads")
    # Procura arquivos Excel na pasta "downloads"
    for filename in os.listdir(downloads_folder):
        if filename.endswith(".xlsx"):
            filepath = os.path.join(downloads_folder, filename)
            file_metadata = {
                'name': filename,
                'parents': [folder_id]
            }
            media = MediaFileUpload(
                filepath,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            file = drive_service.files().create(
                body=file_metadata, media_body=media, fields='id'
            ).execute()
            print(f"Arquivo {filename} enviado para o Drive. ID: {file.get('id')}")

def main():
    # Executa os scripts na ordem desejada
    subprocess.run([sys.executable, "rbt.py"], check=True)
    subprocess.run([sys.executable, "multirio.py"], check=True)
    subprocess.run([sys.executable, "tecon.py"], check=True)
    
    # Envia as planilhas da pasta "downloads" para a pasta específica do Google Drive
    upload_files()

if __name__ == "__main__":
    main()
