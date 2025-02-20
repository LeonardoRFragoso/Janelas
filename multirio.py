import os
import time
import datetime
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

def combine_headers(header_rows):
    """
    Recebe uma lista de elementos <tr> (do <thead>) e retorna uma lista com os títulos
    finais de cada coluna, combinando os textos de cada nível (respeitando os atributos colspan).
    """
    first_row_cells = header_rows[0].find_elements(By.XPATH, "./th | ./td")
    total_cols = 0
    for cell in first_row_cells:
        colspan = cell.get_attribute("colspan")
        try:
            colspan = int(colspan) if colspan and colspan.isdigit() else 1
        except:
            colspan = 1
        total_cols += colspan

    headers = [[] for _ in range(total_cols)]
    for row in header_rows:
        cells = row.find_elements(By.XPATH, "./th | ./td")
        col_index = 0
        for cell in cells:
            text = cell.text.strip().replace("\n", " ")
            colspan = cell.get_attribute("colspan")
            try:
                colspan = int(colspan) if colspan and colspan.isdigit() else 1
            except:
                colspan = 1
            for i in range(colspan):
                if text:
                    headers[col_index + i].append(text)
            col_index += colspan

    final_headers = [" ".join(parts) for parts in headers]
    return final_headers

def main():
    # Configurações de download
    download_dir = os.path.join(os.getcwd(), "downloads")
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)

    chrome_options = Options()
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    # Utilize o webdriver_manager para obter a versão correta do ChromeDriver
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    driver.maximize_window()

    try:
        # Acessa o site
        url = "https://www.multiterminais.com.br/janelas-disponiveis"
        driver.get(url)

        wait = WebDriverWait(driver, 20)
        wait.until(EC.presence_of_element_located((By.ID, "tblJanelasMRIO")))

        # Dias a serem consultados: 0 (hoje), 1 (amanhã) e 2 (depois de amanhã)
        dias_offset = [0, 1, 2]
        dfs = []

        for offset in dias_offset:
            data_consulta = (datetime.datetime.now() + datetime.timedelta(days=offset)).strftime("%d/%m/%Y")
            
            if offset != 0:
                date_field = driver.find_element(By.XPATH, '//*[@id="CPH_Body_txtData"]')
                date_field.click()
                date_field.send_keys(Keys.CONTROL, "a")
                date_field.send_keys(Keys.DELETE)
                date_field.send_keys(data_consulta)
                date_field.send_keys(Keys.RETURN)
                
                filter_button = driver.find_element(By.XPATH, '//*[@id="CPH_Body_btnFiltrar"]')
                filter_button.click()
                
                wait.until(lambda d: d.find_element(By.XPATH, '//*[@id="CPH_Body_txtData"]').get_attribute("value") == data_consulta)
                wait.until(EC.presence_of_element_located((By.ID, "tblJanelasMRIO")))
                time.sleep(1)
            
            # Extração da coluna índice (horários)
            index_header = driver.find_element(By.XPATH, '//*[@id="tblJanelasMRIO"]/thead/tr[1]/th[1]').text.strip()
            index_elements = driver.find_elements(By.XPATH, "//*[starts-with(@id, 'CPH_Body_lvJanelasMultiRio_lblJanelaMultiRio_')]")
            index_column = [el.text.strip() for el in index_elements]
            print(f"Extraindo índices para a data {data_consulta}: {index_column}")

            # Extração do restante da tabela
            table = driver.find_element(By.ID, "tblJanelasMRIO")
            thead = table.find_element(By.TAG_NAME, "thead")
            header_rows = thead.find_elements(By.TAG_NAME, "tr")
            combined_headers = combine_headers(header_rows)
            final_headers = combined_headers[1:]
            print(f"Cabeçalhos extraídos (excluindo índice) para a data {data_consulta}: {final_headers}")

            tbody = table.find_element(By.TAG_NAME, "tbody")
            data_rows = tbody.find_elements(By.TAG_NAME, "tr")
            data = []
            for row in data_rows:
                cells = row.find_elements(By.TAG_NAME, "td")
                if len(cells) > len(final_headers):
                    cells = cells[1:]
                row_data = [cell.text.strip() for cell in cells]
                if len(row_data) < len(final_headers):
                    row_data += [""] * (len(final_headers) - len(row_data))
                elif len(row_data) > len(final_headers):
                    row_data = row_data[:len(final_headers)]
                data.append(row_data)

            df_dia = pd.DataFrame(data, columns=final_headers)
            if len(index_column) == len(df_dia):
                df_dia.insert(0, index_header, index_column)
            else:
                print("Atenção: Número de elementos da coluna índice ({}) difere do número de linhas dos dados ({})."
                      .format(len(index_column), len(df_dia)))
            df_dia["Data"] = data_consulta
            dfs.append(df_dia)

        df_final = pd.concat(dfs, ignore_index=True)
        output_file = os.path.join(download_dir, "janelas_multirio_corrigido.xlsx")
        df_final.to_excel(output_file, index=False)
        print(f"Dados extraídos e salvos em {output_file}")

    finally:
        driver.quit()

if __name__ == "__main__":
    main()
