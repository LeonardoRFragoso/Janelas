#!/usr/bin/env python3
import os
import time
import pandas as pd
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
import requests

# Importa as funções dos módulos export.py e importacao.py
from export import run_export, verificar_dialogo
from importacao import run_import

# Importa as classes necessárias para configuração do Chrome
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

# Configuração do Telegram
TELEGRAM_BOT_TOKEN = "7660740075:AAG0zy6T3QV6pdv2VOwRlxShb0UzVlNwCUk"  # Substitua pelo token do seu bot
TELEGRAM_CHAT_ID = "833732395"  # Substitua pelo Chat ID correto

# Configuração do Google Drive / Google Sheets
GOOGLE_CREDENTIALS_FILE = r"C:\Users\leona\OneDrive\Documentos\DEPOT-PROJECT\gdrive_credentials.json"
# Em vez de usar FOLDER_ID, definimos o ID da planilha que deve ser atualizada:
GOOGLE_SHEET_ID = "1prMkez7J-wbWUGbZp-VLyfHtisSLi-XQ"  # ID da planilha (Google Sheets)

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

def update_google_sheet(file_path, sheet_id, credentials_file):
    """
    Atualiza a planilha do Google Sheets existente com o conteúdo do arquivo Excel.
    Essa função utiliza a biblioteca gspread e gspread_dataframe para:
      - Abrir o Google Sheet pelo seu ID;
      - Limpar a primeira worksheet;
      - Escrever os dados do arquivo Excel na worksheet.
    """
    try:
        import gspread
        from gspread_dataframe import set_with_dataframe
    except ImportError:
        print("❌ As bibliotecas gspread e gspread_dataframe não estão instaladas. Por favor, instale-as.")
        raise

    try:
        gc = gspread.service_account(filename=credentials_file)
        sh = gc.open_by_key(sheet_id)
        worksheet = sh.sheet1  # Atualiza a primeira aba
        df = pd.read_excel(file_path)
        worksheet.clear()
        set_with_dataframe(worksheet, df)
        print("✅ Planilha atualizada com sucesso no Google Sheets.")
        return sheet_id
    except Exception as e:
        print("❌ Erro ao atualizar a planilha no Google Sheets:", e)
        raise

def main(chrome_driver_path: str, usuario: str, senha: str):
    """
    Inicializa o WebDriver do Chrome, realiza o login, executa os processos
    de exportação e importação (que extraem e salvam a coluna "DI / BOOKING / CTE"
    como primeira coluna na planilha) e, ao final, atualiza a planilha do Google Sheets
    e envia um resumo via Telegram.
    """
    # Configurar as opções do Chrome
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--log-level=3")
    chrome_options.add_argument("--silent")

    # Configurar o serviço do Chrome com o caminho fornecido
    service = Service(chrome_driver_path)
    driver = webdriver.Chrome(service=service, options=chrome_options)
    wait = WebDriverWait(driver, 20)
    
    try:
        # 1) LOGIN E MENU
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
        export_summary = run_export(driver, wait)
        
        # Prepara a aba 2 para importação
        if len(driver.window_handles) < 2:
            driver.execute_script("window.open();")
        driver.switch_to.window(driver.window_handles[1])
        time.sleep(3)
        verificar_dialogo(driver, wait)
        
        # Executa o loop de extração para Importação
        print("Iniciando extração de importação...")
        import_summary = run_import(driver, wait)
        
        print("✅ Processo concluído e dados salvos (Exportação e Importação).")
        
        # Leitura da planilha para compilar o resumo final
        if os.path.exists("informacoes_janelas.xlsx"):
            df = pd.read_excel("informacoes_janelas.xlsx")
            total_registros = len(df)
            registros_export = len(df[df["Tipo"] == "exportacao"])
            registros_import = len(df[df["Tipo"] == "importacao"])
        else:
            total_registros = registros_export = registros_import = 0
        
        # Monta o resumo da extração
        resumo = (
            "📊 Resumo da Extração:\n\n"
            "Exportação:\n"
            f"  - Registros processados: {export_summary.get('processed', 0)}\n"
            f"  - Duplicatas consecutivas: {export_summary.get('duplicates', 0)}\n\n"
            "Importação:\n"
            f"  - Registros processados: {import_summary.get('processed', 0)}\n"
            f"  - Duplicatas consecutivas: {import_summary.get('duplicates', 0)}\n\n"
            "Planilha 'informacoes_janelas.xlsx':\n"
            f"  - Total de registros: {total_registros}\n"
            f"  - Registros de Exportação: {registros_export}\n"
            f"  - Registros de Importação: {registros_import}\n"
        )
        
        # Em vez de criar um novo arquivo no Drive, atualizamos a planilha existente no Google Sheets:
        excel_file = "informacoes_janelas.xlsx"
        if os.path.exists(excel_file):
            try:
                sheet_id = update_google_sheet(excel_file, GOOGLE_SHEET_ID, GOOGLE_CREDENTIALS_FILE)
                resumo += f"\n📤 Planilha atualizada com sucesso no Google Sheets. ID: {sheet_id}\n"
            except Exception as e:
                resumo += f"\n❌ Erro ao atualizar a planilha no Google Sheets: {e}\n"
        else:
            resumo += "\n⚠️ Arquivo 'informacoes_janelas.xlsx' não encontrado. A planilha não foi atualizada no Google Sheets.\n"
        
        print(resumo)
        send_telegram_message(resumo)
        
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
