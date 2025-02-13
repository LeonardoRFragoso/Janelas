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
    # Calcula o total de colunas a partir da primeira linha do cabeçalho
    first_row_cells = header_rows[0].find_elements(By.XPATH, "./th | ./td")
    total_cols = 0
    for cell in first_row_cells:
        colspan = cell.get_attribute("colspan")
        try:
            colspan = int(colspan) if colspan and colspan.isdigit() else 1
        except:
            colspan = 1
        total_cols += colspan

    # Cria uma lista de listas para acumular os textos para cada coluna
    headers = [[] for _ in range(total_cols)]
    
    # Para cada linha de cabeçalho, distribui o texto nas posições correspondentes
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

    # Junta os pedaços para cada coluna, separando por espaço
    final_headers = [" ".join(parts) for parts in headers]
    return final_headers

# Configurações do ChromeDriver
chrome_driver_path = r"C:\Users\leonardo.fragoso\Desktop\Projetos\Depot-Project\chromedriver.exe"
chrome_options = Options()
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("--log-level=3")
chrome_options.add_argument("--silent")

service = Service(chrome_driver_path)
driver = webdriver.Chrome(service=service, options=chrome_options)

# Acessa o site
url = "https://www.multiterminais.com.br/janelas-disponiveis"
driver.get(url)

wait = WebDriverWait(driver, 20)  # aumento do timeout para 20 segundos
wait.until(EC.presence_of_element_located((By.ID, "tblJanelasMRIO")))

# Lista de dias a serem consultados: 0 (data atual), 1 (dia seguinte) e 2 (dois dias à frente)
dias_offset = [0, 1, 2]

# Lista para armazenar os DataFrames de cada dia
dfs = []

for offset in dias_offset:
    # Calcula a data com base no offset
    data_consulta = (datetime.datetime.now() + datetime.timedelta(days=offset)).strftime("%d/%m/%Y")
    
    if offset != 0:
        # Atualiza o campo de data
        date_field = driver.find_element(By.XPATH, '//*[@id="CPH_Body_txtData"]')
        date_field.click()
        date_field.send_keys(Keys.CONTROL, "a")
        date_field.send_keys(Keys.DELETE)
        date_field.send_keys(data_consulta)
        date_field.send_keys(Keys.RETURN)
        
        # Clica no botão Filtrar
        filter_button = driver.find_element(By.XPATH, '//*[@id="CPH_Body_btnFiltrar"]')
        filter_button.click()
        
        # Aguarda até que o campo de data exiba o valor desejado
        wait.until(lambda d: d.find_element(By.XPATH, '//*[@id="CPH_Body_txtData"]').get_attribute("value") == data_consulta)
        # Aguarda a presença da tabela atualizada
        wait.until(EC.presence_of_element_located((By.ID, "tblJanelasMRIO")))
        time.sleep(1)  # pausa extra para garantir que os dados sejam atualizados
    
    # ---------------------------
    # Extração da coluna índice (horários)
    # ---------------------------
    index_header = driver.find_element(By.XPATH, '//*[@id="tblJanelasMRIO"]/thead/tr[1]/th[1]').text.strip()
    index_elements = driver.find_elements(By.XPATH, "//*[starts-with(@id, 'CPH_Body_lvJanelasMultiRio_lblJanelaMultiRio_')]")
    index_column = [el.text.strip() for el in index_elements]
    print(f"Extraindo índices para a data {data_consulta}: {index_column}")

    # ---------------------------
    # Extração do restante da tabela
    # ---------------------------
    table = driver.find_element(By.ID, "tblJanelasMRIO")
    thead = table.find_element(By.TAG_NAME, "thead")
    header_rows = thead.find_elements(By.TAG_NAME, "tr")
    combined_headers = combine_headers(header_rows)
    # Remove o primeiro título (índice), pois já está extraído
    final_headers = combined_headers[1:]
    print(f"Cabeçalhos extraídos (excluindo índice) para a data {data_consulta}: {final_headers}")

    tbody = table.find_element(By.TAG_NAME, "tbody")
    data_rows = tbody.find_elements(By.TAG_NAME, "tr")

    data = []
    for row in data_rows:
        cells = row.find_elements(By.TAG_NAME, "td")
        # Se houver célula extra (coluna de índice), descarta a primeira
        if len(cells) > len(final_headers):
            cells = cells[1:]
        row_data = [cell.text.strip() for cell in cells]
        # Ajusta o número de células se necessário
        if len(row_data) < len(final_headers):
            row_data += [""] * (len(final_headers) - len(row_data))
        elif len(row_data) > len(final_headers):
            row_data = row_data[:len(final_headers)]
        data.append(row_data)

    # Cria o DataFrame e insere a coluna de índices
    df_dia = pd.DataFrame(data, columns=final_headers)
    if len(index_column) == len(df_dia):
        df_dia.insert(0, index_header, index_column)
    else:
        print("Atenção: Número de elementos da coluna índice ({}) difere do número de linhas dos dados ({})."
              .format(len(index_column), len(df_dia)))
    
    # Adiciona a coluna com a data consultada
    df_dia["Data"] = data_consulta

    dfs.append(df_dia)

# Concatena todos os DataFrames (dados dos três dias) em um único DataFrame final
df_final = pd.concat(dfs, ignore_index=True)

# Exporta para Excel
output_file = "janelas_multirio_corrigido.xlsx"
df_final.to_excel(output_file, index=False)
print(f"Dados extraídos e salvos em {output_file}")

driver.quit()
