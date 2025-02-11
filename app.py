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
GOOGLE_CREDENTIALS_FILE = r"C:\Users\leonardo.fragoso\Desktop\Projetos\Depot-Project\gdrive_credentials.json"
# ID da planilha do Google Sheets que deverá ser atualizada (mantém a mesma URL)
GOOGLE_SHEET_ID = "1prMkez7J-wbWUGbZp-VLyfHtisSLi-XQ"  # ID da planilha

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
    Essa função utiliza as bibliotecas gspread e gspread_dataframe para:
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
        worksheet.clear()  # Remove os dados antigos
        set_with_dataframe(worksheet, df)  # Escreve os novos dados
        print("✅ Planilha atualizada com sucesso no Google Sheets.")
        return sheet_id
    except Exception as e:
        print("❌ Erro ao atualizar a planilha no Google Sheets:", e)
        raise

def main(chrome_driver_path: str, usuario: str, senha: str):
    """
    Inicializa o WebDriver do Chrome, realiza o login, executa os processos
    de exportação e importação e, ao final, atualiza a planilha do Google Sheets
    (substituindo os dados antigos) e envia um resumo via Telegram.
    """
    # Configurar as opções do Chrome
    chrome_options = Options()
    # Caso deseje executar em modo headless, descomente as linhas abaixo:
    # chrome_options.add_argument("--headless=new")
    # chrome_options.add_argument("--disable-gpu")
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
        
        # 2) ABRIR A ABA SECUNDÁRIA PARA EXTRAÇÃO (EXPORTAÇÃO/IMPORTAÇÃO)
        export_url = "https://lookerstudio.google.com/u/0/reporting/55ec93e4-3114-46d5-9125-79d7191b1c0a/page/p_5jlrz7xapd"
        print("Abrindo nova aba para extração (exportação/importação)...")
        driver.execute_script(f"window.open('{export_url}', '_blank');")
        # Define: janela[0] = portal principal; janela[1] = extração (export)
        driver.switch_to.window(driver.window_handles[1])
        time.sleep(3)  # Aguarda carregamento da aba
        
        # 3) EXTRAÇÃO DA EXPORTAÇÃO
        print("Iniciando extração de exportação...")
        export_summary = run_export(driver, wait)
        
        # 4) EXTRAÇÃO DA IMPORTAÇÃO
        # Certifica que a aba de extração (índice 1) está ativa para a extração de DI/BOOKING/CTE
        driver.switch_to.window(driver.window_handles[1])
        time.sleep(2)
        # Caso haja algum diálogo a ser tratado, a função abaixo já é chamada pelo módulo importação
        verificar_dialogo(driver, wait)
        print("Iniciando extração de importação...")
        import_summary = run_import(driver, wait)
        
        print("✅ Processos de Exportação e Importação concluídos.")
        
        # 5) LEITURA DA PLANILHA GERADA E COMPILAÇÃO DO RESUMO
        excel_file = "informacoes_janelas.xlsx"
        if os.path.exists(excel_file):
            df = pd.read_excel(excel_file)
            total_registros = len(df)
            registros_export = len(df[df["Tipo"] == "exportacao"])
            registros_import = len(df[df["Tipo"] == "importacao"])
        else:
            total_registros = registros_export = registros_import = 0
        
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
        
        # 6) ATUALIZAÇÃO DO GOOGLE SHEETS
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
    # Configurações definidas diretamente no código (poderão ser parametrizadas se necessário)
    chrome_driver_path = r"C:\Users\leonardo.fragoso\Desktop\Projetos\Depot-Project\chromedriver.exe"
    usuario = "redex.gate03@itracker.com.br"
    senha = "123"
    
    main(chrome_driver_path, usuario, senha)
