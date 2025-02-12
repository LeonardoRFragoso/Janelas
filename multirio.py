from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

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
chrome_driver_path = r"C:\Users\leonardo.fragoso\Desktop\Projetos\multirio-janelas\chromedriver.exe"
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

wait = WebDriverWait(driver, 10)
wait.until(EC.presence_of_element_located((By.ID, "tblJanelasMRIO")))

# ---------------------------
# 1. Extração da coluna índice (horários)
# ---------------------------
# Extrai o título da coluna índice (primeira célula do cabeçalho)
index_header = driver.find_element(By.XPATH, '//*[@id="tblJanelasMRIO"]/thead/tr[1]/th[1]').text.strip()

# Extrai os elementos dos horários usando o XPath que seleciona todos os IDs iniciados por "CPH_Body_lvJanelasMultiRio_lblJanelaMultiRio_"
index_elements = driver.find_elements(By.XPATH, "//*[starts-with(@id, 'CPH_Body_lvJanelasMultiRio_lblJanelaMultiRio_')]")
index_column = [el.text.strip() for el in index_elements]
# Verifica se foram encontrados valores; caso contrário, pode ser necessário ajustar o tempo de espera ou o XPath.
print("Valores do índice extraídos:", index_column)

# ---------------------------
# 2. Extração do restante da tabela
# ---------------------------
table = driver.find_element(By.ID, "tblJanelasMRIO")

# Usa o <thead> para os cabeçalhos e <tbody> para os dados
thead = table.find_element(By.TAG_NAME, "thead")
header_rows = thead.find_elements(By.TAG_NAME, "tr")
# Combina os cabeçalhos (que, neste caso, inclui o cabeçalho do índice na posição 0)
combined_headers = combine_headers(header_rows)
# Removemos o primeiro título, pois a coluna de índice será inserida separadamente
final_headers = combined_headers[1:]
print("Cabeçalhos extraídos (excluindo índice):", final_headers)

tbody = table.find_element(By.TAG_NAME, "tbody")
data_rows = tbody.find_elements(By.TAG_NAME, "tr")

# Extração dos dados (assumindo que as células do <tbody> correspondem às colunas 1 em diante)
data = []
for row in data_rows:
    cells = row.find_elements(By.TAG_NAME, "td")
    # Se houver uma célula extra (da coluna de índice) no <tbody>, descartamos a primeira
    if len(cells) > len(final_headers):
        cells = cells[1:]
    row_data = [cell.text.strip() for cell in cells]
    # Ajusta o número de células se necessário
    if len(row_data) < len(final_headers):
        row_data += [""] * (len(final_headers) - len(row_data))
    elif len(row_data) > len(final_headers):
        row_data = row_data[:len(final_headers)]
    data.append(row_data)

# Cria o DataFrame com as demais colunas
df = pd.DataFrame(data, columns=final_headers)

# Insere a coluna de índice (horários) como a primeira coluna
# Verifica se o número de linhas extraídas corresponde ao DataFrame
if len(index_column) == len(df):
    df.insert(0, index_header, index_column)
else:
    print("Atenção: Número de elementos da coluna índice ({}) difere do número de linhas dos dados ({})."
          .format(len(index_column), len(df)))

# Exporta para Excel
output_file = "janelas_multirio_corrigido.xlsx"
df.to_excel(output_file, index=False)
print(f"Dados extraídos e salvos em {output_file}")

driver.quit()
