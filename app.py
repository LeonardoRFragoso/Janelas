#!/usr/bin/env python3
import os
import time
import pandas as pd
from datetime import datetime
import subprocess
import requests

# Importa as funções dos módulos export.py e importacao.py
from export import run_export, verificar_dialogo
from importacao import run_import

# Importações necessárias para o Selenium
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait

# Configuração do Telegram
TELEGRAM_BOT_TOKEN = "7660740075:AAG0zy6T3QV6pdv2VOwRlxShb0UzVlNwCUk"  # Substitua pelo token do seu bot
TELEGRAM_CHAT_ID = "833732395"  # Substitua pelo Chat ID correto

# Configuração do Google Drive
GOOGLE_CREDENTIALS_FILE = r"C:\Users\leonardo.fragoso\Desktop\Projetos\Depot-Project\gdrive_credentials.json"
# IDs dos arquivos no Google Drive – estes arquivos já devem existir e serão atualizados (sobrescritos) a cada execução:
# ID da planilha "informacoes_janelas" extraído do link:
# https://docs.google.com/spreadsheets/d/1prMkez7J-wbWUGbZp-VLyfHtisSLi-XQ/edit?gid=404271548#gid=404271548
GOOGLE_DRIVE_FILE_ID_INFORMACOES = "1prMkez7J-wbWUGbZp-VLyfHtisSLi-XQ"
GOOGLE_DRIVE_FILE_ID_MULTIRIO = "1Eh58MkuHwyHpYCscMPD9X1r_83dbHV63"  # Link permanente informado para a planilha do multirio

def send_telegram_message(message):
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    data = {"chat_id": TELEGRAM_CHAT_ID, "text": message}
    try:
        response = requests.post(url, data=data)
        if response.status_code == 200:
            print("✅ Mensagem enviada ao Telegram com sucesso!")
        else:
            print("❌ Erro ao enviar mensagem ao Telegram:", response.text)
    except Exception as e:
        print("❌ Erro ao enviar mensagem ao Telegram:", e)

def update_excel_file_on_drive(file_path, file_id, credentials_file):
    """
    Atualiza (sobrescreve) um arquivo Excel (.xlsx) existente no Google Drive,
    mantendo o mesmo link (ID) para acesso.
    """
    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaFileUpload
    except ImportError:
        print("❌ As bibliotecas google-auth e google-api-python-client não estão instaladas. Por favor, instale-as.")
        raise

    # Define o escopo e cria as credenciais
    scopes = ['https://www.googleapis.com/auth/drive']
    creds = service_account.Credentials.from_service_account_file(credentials_file, scopes=scopes)
    service = build('drive', 'v3', credentials=creds)
    
    file_mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    media = MediaFileUpload(file_path, mimetype=file_mime_type, resumable=True)
    try:
        updated_file = service.files().update(
            fileId=file_id,
            media_body=media
        ).execute()
        if updated_file is None:
            raise ValueError("A resposta da API do Google Drive foi None.")
        file_id_updated = updated_file.get('id')
        if not file_id_updated:
            raise ValueError("ID do arquivo atualizado não encontrado na resposta da API do Google Drive.")
        print(f"✅ Arquivo '{file_path}' atualizado com sucesso no Google Drive.")
        return file_id_updated
    except Exception as e:
        print("❌ Erro ao atualizar o arquivo no Google Drive:", e)
        raise

def main(chrome_driver_path: str, usuario: str, senha: str):
    """
    Inicializa o WebDriver do Chrome, realiza o login, executa os processos
    de exportação e importação, chama o script multirio.py para extrair os dados de janelas MultiRio,
    atualiza os arquivos Excel e envia um resumo via Telegram.
    """
    # Configurar as opções do Chrome
    chrome_options = Options()
    # chrome_options.add_argument("--headless=new")
    # chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--log-level=3")
    chrome_options.add_argument("--silent")

    # Configurar o serviço do Chrome
    service = Service(chrome_driver_path)
    driver = webdriver.Chrome(service=service, options=chrome_options)
    wait = WebDriverWait(driver, 20)
    
    try:
        # 1) LOGIN E NAVEGAÇÃO NO MENU
        print("Iniciando login...")
        driver.get("https://portaldeservicos.riobrasilterminal.com/tosp/Workspace/load#/CentralCeR")
        campo = wait.until(lambda d: d.find_element(By.XPATH, '//*[@id="username"]'))
        campo.send_keys(usuario)
        driver.find_element(By.XPATH, '//*[@id="pass"]').send_keys(senha)
        driver.find_element(By.XPATH, '/html/body/div[1]/div/div/form/table/tbody/tr[2]/td[1]/button').click()
        wait.until(lambda d: d.find_element(By.XPATH, '//*[@id="bs-example-navbar-collapse-1"]/ul/li[2]/a/span'))
        print("✅ Login realizado com sucesso!")
        
        # Navegar no menu
        driver.find_element(By.XPATH, '//*[@id="bs-example-navbar-collapse-1"]/ul/li[2]/a/span').click()
        time.sleep(1)
        driver.find_element(By.XPATH, '//*[@id="bs-example-navbar-collapse-1"]/ul/li[2]/ul/li[1]/a/span').click()
        time.sleep(2)
        print("✅ Menu navegado com sucesso!")
        
        # Abrir a segunda aba para exportação
        export_url = "https://lookerstudio.google.com/u/0/reporting/55ec93e4-3114-46d5-9125-79d7191b1c0a/page/p_5jlrz7xapd"
        print("Abrindo segunda aba para exportação...")
        driver.execute_script(f"window.open('{export_url}', '_blank');")
        
        # Executa o loop de extração para Exportação
        print("Iniciando extração de exportação...")
        export_summary = run_export(driver, wait) or {}
        
        # Prepara a aba 2 para importação
        if len(driver.window_handles) < 2:
            driver.execute_script("window.open();")
        driver.switch_to.window(driver.window_handles[1])
        time.sleep(3)
        verificar_dialogo(driver, wait)
        
        # Executa o loop de extração para Importação
        print("Iniciando extração de importação...")
        import_summary = run_import(driver, wait) or {}
        
        print("✅ Processos de Exportação e Importação concluídos e dados salvos.")
        
        # Compila o resumo final com base na planilha "informacoes_janelas.xlsx"
        excel_file = "informacoes_janelas.xlsx"
        if os.path.exists(excel_file):
            df = pd.read_excel(excel_file)
            total_registros = len(df)
            registros_export = len(df[df["Tipo"] == "exportacao"])
            registros_import = len(df[df["Tipo"] == "importacao"])
        else:
            total_registros = registros_export = registros_import = 0
        
        # Atualiza a planilha 'informacoes_janelas.xlsx' no Google Drive (substituindo o conteúdo)
        if os.path.exists(excel_file):
            try:
                file_id_informacoes = update_excel_file_on_drive(
                    excel_file, 
                    GOOGLE_DRIVE_FILE_ID_INFORMACOES, 
                    GOOGLE_CREDENTIALS_FILE
                )
            except Exception as e:
                file_id_informacoes = None
        else:
            file_id_informacoes = None
        
        # 2) EXECUÇÃO DO SCRIPT multirio.py
        print("Iniciando extração do Multirio chamando o script multirio.py...")
        # Em ambientes Windows, se necessário, use "python" em vez de "python3"
        subprocess.run(["python3", "multirio.py"], check=True)
        print("Extração do Multirio concluída.")
        
        # Renomeia o arquivo gerado pelo multirio.py (supondo que seja "janelas_multirio_corrigido.xlsx")
        origem = "janelas_multirio_corrigido.xlsx"
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        novo_nome = f"Janelas_multirio_{timestamp}.xlsx"
        if os.path.exists(origem):
            os.rename(origem, novo_nome)
            print(f"Arquivo renomeado para {novo_nome}")
        else:
            novo_nome = None
            print("❌ Arquivo do Multirio não encontrado!")
        
        # Atualiza a planilha do Multirio no Google Drive (substituindo o conteúdo)
        if novo_nome and os.path.exists(novo_nome):
            try:
                file_id_multirio = update_excel_file_on_drive(
                    novo_nome, 
                    GOOGLE_DRIVE_FILE_ID_MULTIRIO, 
                    GOOGLE_CREDENTIALS_FILE
                )
            except Exception as e:
                file_id_multirio = None
        else:
            file_id_multirio = None
        
        # Monta a mensagem final de resumo com informações detalhadas
        final_summary = (
            "🚀 *Processo de Extração Concluído!*\n\n"
            "*Resumo Geral:*\n"
            f"  - Finalizado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n\n"
            "*Exportação:*\n"
            f"  - Registros processados: {export_summary.get('processed', 0)}\n"
            f"  - Duplicatas consecutivas: {export_summary.get('duplicates', 0)}\n\n"
            "*Importação:*\n"
            f"  - Registros processados: {import_summary.get('processed', 0)}\n"
            f"  - Duplicatas consecutivas: {import_summary.get('duplicates', 0)}\n\n"
            "*Planilha 'informacoes_janelas.xlsx':*\n"
            f"  - Total de registros: {total_registros}\n"
            f"  - Registros de Exportação: {registros_export}\n"
            f"  - Registros de Importação: {registros_import}\n"
            f"  - Atualizada no Google Drive (ID): {file_id_informacoes if file_id_informacoes else 'Falha na atualização'}\n\n"
            "*Janelas Multirio:*\n"
            f"  - Arquivo gerado: {novo_nome if novo_nome else 'Arquivo não encontrado'}\n"
            f"  - Atualizada no Google Drive (ID): {file_id_multirio if file_id_multirio else 'Falha na atualização'}\n\n"
            "✅ *Processo concluído com sucesso!*"
        )
        
        print(final_summary)
        send_telegram_message(final_summary)
        
    except Exception as e:
        print("❌ Erro inesperado:", e)
    finally:
        driver.quit()
        print("✅ Navegador fechado.")

if __name__ == '__main__':
    # Informações definidas diretamente no código (sem interação com o usuário)
    chrome_driver_path = r"C:\Users\leonardo.fragoso\Desktop\Projetos\Depot-Project\chromedriver.exe"
    usuario = "redex.gate03@itracker.com.br"
    senha = "123"
    
    main(chrome_driver_path, usuario, senha)
