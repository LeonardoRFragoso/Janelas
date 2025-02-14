from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import datetime
import time

def combine_headers(header_rows):
    """
    Recebe uma lista de elementos <tr> (do <thead>) e retorna uma lista com os títulos
    finais de cada coluna, combinando os textos de cada nível (respeitando os atributos colspan).
    """
    first_row_cells = header_rows[0].find_elements(By.XPATH, "./th | ./td")
    total_cols = sum(int(cell.get_attribute("colspan") or 1) for cell in first_row_cells)

    headers = [[] for _ in range(total_cols)]

    for row in header_rows:
        cells = row.find_elements(By.XPATH, "./th | ./td")
        col_index = 0
        for cell in cells:
            text = cell.text.strip().replace("\n", " ")
            colspan = int(cell.get_attribute("colspan") or 1)
            for i in range(colspan):
                if text:
                    headers[col_index + i].append(text)
            col_index += colspan

    return [" ".join(parts) for parts in headers]

# Configurações do ChromeDriver
chrome_driver_path = r"/home/dev/Documentos/Janelas/chromedriver"
chrome_options = Options()
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("--log-level=3")
chrome_options.add_argument("--silent")

service = Service(chrome_driver_path)
driver = webdriver.Chrome(service=service, options=chrome_options)

# Acessa o site
url = "https://www.multiterminais.com.br/janelas-disponiveis"
driver.get(url)

wait = WebDriverWait(driver, 20)
wait.until(EC.presence_of_element_located((By.ID, "tblJanelasMRIO")))

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
        time.sleep(1)

    # Verifica se há resultados
    try:
        no_results_message = driver.find_element(By.XPATH, '//*[@id="CPH_Body_pnlSemResultadosMRIO"]/p')
        print(f"Sem resultados para {data_consulta}. Pulando para o próximo dia.")
        continue  # Pula para o próximo dia
    except:
        pass  # Continua normalmente se não encontrou a mensagem de "Sem Resultados"

    # Extração dos índices
    index_header = driver.find_element(By.XPATH, '//*[@id="tblJanelasMRIO"]/thead/tr[1]/th[1]').text.strip()
    index_elements = driver.find_elements(By.XPATH, "//*[starts-with(@id, 'CPH_Body_lvJanelasMultiRio_lblJanelaMultiRio_')]")
    index_column = [el.text.strip() for el in index_elements]

    # Extração do restante da tabela
    table = driver.find_element(By.ID, "tblJanelasMRIO")
    thead = table.find_element(By.TAG_NAME, "thead")
    header_rows = thead.find_elements(By.TAG_NAME, "tr")
    combined_headers = combine_headers(header_rows)
    final_headers = combined_headers[1:]

    tbody = table.find_element(By.TAG_NAME, "tbody")
    data_rows = tbody.find_elements(By.TAG_NAME, "tr")

    data = []
    for row in data_rows:
        cells = row.find_elements(By.TAG_NAME, "td")
        if len(cells) > len(final_headers):
            cells = cells[1:]
        row_data = [cell.text.strip() for cell in cells]
        row_data += [""] * (len(final_headers) - len(row_data))
        data.append(row_data)

    df_dia = pd.DataFrame(data, columns=final_headers)
    if len(index_column) == len(df_dia):
        df_dia.insert(0, index_header, index_column)

    df_dia["Data"] = data_consulta
    dfs.append(df_dia)

# Criação do DataFrame final
if dfs:
    df_final = pd.concat(dfs, ignore_index=True)
else:
    print("Nenhuma data retornou resultados. Criando arquivo Excel vazio.")
    df_final = pd.DataFrame(columns=["Data"])  # Criando um DataFrame vazio com a coluna "Data"

output_file = "janelas_multirio_corrigido.xlsx"
df_final.to_excel(output_file, index=False)
print(f"Dados extraídos e salvos em {output_file}")

driver.quit()
