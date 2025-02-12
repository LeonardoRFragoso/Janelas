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

# Importa as classes necessárias para configuração do Chrome
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait

# Configuração do Telegram
TELEGRAM_BOT_TOKEN = "7660740075:AAG0zy6T3QV6pdv2VOwRlxShb0UzVlNwCUk"  # Substitua pelo token do seu bot
TELEGRAM_CHAT_ID = "833732395"  # Substitua pelo Chat ID correto

# Configuração do Google Sheets
GOOGLE_CREDENTIALS_FILE = r"/home/dev/Documentos/DEPOT-PROJECT/gdrive_credentials.json"
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
        worksheet.clear()  # Limpa os dados antigos
        set_with_dataframe(worksheet, df)  # Escreve os novos dados
        print("✅ Planilha atualizada com sucesso no Google Sheets.")
        return sheet_id
    except Exception as e:
        if "This operation is not supported for this document" in str(e):
            print(f"❌ Erro: O documento com ID '{sheet_id}' não suporta esta operação. "
                  "Verifique se o ID corresponde a uma planilha nativa do Google Sheets e se a conta de serviço tem permissão de edição.")
        else:
            print("❌ Erro ao atualizar a planilha no Google Sheets:", e)
        raise

def main(chrome_driver_path: str, usuario: str, senha: str):
    """
    Inicializa o WebDriver do Chrome, realiza o login, executa os processos
    de exportação e importação, e após isso chama o script multirio.py para
    extrair os dados de janelas MultiRio. Em seguida, renomeia o arquivo gerado,
    atualiza a planilha do Google Sheets e envia um resumo via Telegram.
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
    driver =  webdriver.Chrome(service=service, options=chrome_options)
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
        
        print("✅ Processos de Exportação e Importação concluídos e dados salvos.")
        
        # Leitura da planilha para compilar o resumo final
        excel_file = "informacoes_janelas.xlsx"
        if os.path.exists(excel_file):
            df = pd.read_excel(excel_file)
            total_registros = len(df)
            registros_export = len(df[df["Tipo"] == "exportacao"])
            registros_import = len(df[df["Tipo"] == "importacao"])
        else:
            total_registros = registros_export = registros_import = 0
        
        resumo = (
            " Resumo da Extração:\n\n"
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
        
        # Atualiza a planilha 'informacoes_janelas.xlsx' no Google Sheets
        if os.path.exists(excel_file):
            try:
                sheet_id = update_google_sheet(excel_file, GOOGLE_SHEET_ID, GOOGLE_CREDENTIALS_FILE)
                resumo += f"\n Planilha 'informacoes_janelas.xlsx' atualizada no Google Sheets. ID: {sheet_id}\n"
            except Exception as e:
                resumo += f"\n❌ Erro ao atualizar 'informacoes_janelas.xlsx' no Google Sheets: {e}\n"
        else:
            resumo += "\n⚠️ Arquivo 'informacoes_janelas.xlsx' não encontrado. A planilha não foi atualizada no Google Sheets.\n"
        
        # 2) EXECUÇÃO DO SCRIPT multirio.py
        print("Iniciando extração do Multirio chamando o script multirio.py...")
        # Chama o script multirio.py; certifique-se de que ele esteja no mesmo diretório ou informe o caminho correto.
        subprocess.run(["python3", "multirio.py"], check=True)
        print("Extração do Multirio concluída.")
        
        # Renomeia o arquivo gerado (supondo que o multirio.py salve em "janelas_multirio_corrigido.xlsx")
        origem = "janelas_multirio_corrigido.xlsx"
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        novo_nome = f"Janelas_multirio_{timestamp}.xlsx"
        if os.path.exists(origem):
            os.rename(origem, novo_nome)
            print(f"Arquivo renomeado para {novo_nome}")
        else:
            print("❌ Arquivo do Multirio não encontrado!")
        
        # Atualiza a planilha do Google Sheets com o arquivo do Multirio
        if os.path.exists(novo_nome):
            try:
                sheet_id_multirio = update_google_sheet(novo_nome, GOOGLE_SHEET_ID, GOOGLE_CREDENTIALS_FILE)
                resumo += f"\n Planilha 'Janelas_multirio' atualizada no Google Sheets. ID: {sheet_id_multirio}\n"
            except Exception as e:
                resumo += f"\n❌ Erro ao atualizar 'Janelas_multirio' no Google Sheets: {e}\n"
        else:
            resumo += "\n⚠️ Arquivo 'Janelas_multirio' não encontrado. A planilha não foi atualizada no Google Sheets.\n"
        
        print(resumo)
        send_telegram_message(resumo)
        
    except Exception as e:
        print("❌ Erro inesperado:", e)
    finally:
        driver.quit()
        print("✅ Navegador fechado.")

if __name__ == '__main__':
    # Informações definidas diretamente no código (sem interação com o usuário)
    chrome_driver_path = r"/home/dev/Documentos/DEPOT-PROJECT/chromedriver"
    usuario = "redex.gate03@itracker.com.br"
    senha = "123"
    
    main(chrome_driver_path, usuario, senha)
