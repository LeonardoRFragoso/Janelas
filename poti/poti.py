import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

def iniciar_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 15)
    return driver, wait

def realizar_login(driver, wait, usuario, senha):
    driver.get("https://portaldeservicos.riobrasilterminal.com/tosp/Workspace/load#/CentralCeR")
    print("🔍 Aguardando carregamento da página de login...")
    try:
        campo_usuario = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="username"]')))
        campo_usuario.send_keys(usuario)
        campo_senha = driver.find_element(By.XPATH, '//*[@id="pass"]')
        campo_senha.send_keys(senha)
        botao_login = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/form/table/tbody/tr[2]/td[1]/button')
        botao_login.click()
        # Aguarda que o menu apareça após o login, por exemplo.
        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="bs-example-navbar-collapse-1"]/ul/li[2]/a/span')))
        print("✅ Login realizado com sucesso!")
    except TimeoutException:
        print("❌ Timeout durante o login. Verifique se a página ou os xpaths mudaram.")
        driver.quit()

def navegar_menu(wait):
    print("🔍 Navegando no menu...")
    menu_xpath = '//*[@id="bs-example-navbar-collapse-1"]/ul/li[2]/a/span'
    submenu_xpath = '//*[@id="bs-example-navbar-collapse-1"]/ul/li[2]/ul/li[1]/a/span'
    try:
        wait.until(EC.element_to_be_clickable((By.XPATH, menu_xpath))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, submenu_xpath))).click()
        print("✅ Menu navegado com sucesso!")
    except TimeoutException:
        print("❌ Erro ao navegar no menu. Verifique os xpaths.")

def processar_reserva(driver, wait, reserva):
    # Localizadores
    campo_busca_xpath = '//*[@id="opcoesBusca"]/div[4]/div/input'
    botao_busca_xpath = '//*[@id="opcoesBusca"]/div[11]/div[2]/button'
    botao_liberacao_xpath = '//*[@id="liberacao"]/div[7]/a'
    novo_botao_xpath = '//*[@id="#conteudo"]/div/tabset[2]/div/ul/li[4]/a'
    
    # Inserir o valor da reserva no campo de busca
    try:
        campo_busca = wait.until(EC.element_to_be_clickable((By.XPATH, campo_busca_xpath)))
        campo_busca.click()
        campo_busca.clear()
        campo_busca.send_keys(reserva)
        time.sleep(1)
        wait.until(EC.element_to_be_clickable((By.XPATH, botao_busca_xpath))).click()
        time.sleep(2)
        print("✅ Pesquisa concluída!")
    except TimeoutException:
        print(f"❌ Erro ao pesquisar a reserva: {reserva}")
        return None
    
    # Navegar para a aba da tabela
    try:
        wait.until(EC.element_to_be_clickable((By.XPATH, botao_liberacao_xpath))).click()
        time.sleep(2)
        wait.until(EC.element_to_be_clickable((By.XPATH, novo_botao_xpath))).click()
        time.sleep(3)
        print("✅ Navegação para a aba da tabela concluída!")
    except TimeoutException:
        print(f"❌ Erro ao navegar para a aba da tabela para a reserva: {reserva}")
        return None
    
    # Identificar a tabela com maior número de linhas visíveis
    try:
        tabelas = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'ui-grid-canvas')))
        print(f"✅ {len(tabelas)} tabelas detectadas.")
    except TimeoutException:
        print(f"❌ Nenhuma tabela encontrada para a reserva: {reserva}")
        return None
    
    tabela_selecionada = None
    max_linhas = 0
    for i, tabela in enumerate(tabelas):
        linhas = tabela.find_elements(By.XPATH, './/div[contains(@class, "ui-grid-row")]')
        num_linhas = len(linhas)
        print(f"📋 Tabela {i+1}: {num_linhas} linhas detectadas.")
        if num_linhas > max_linhas:
            max_linhas = num_linhas
            tabela_selecionada = tabela
    
    if not tabela_selecionada:
        print(f"❌ Nenhuma tabela válida encontrada para a reserva: {reserva}")
        return None
    print(f"✅ Tabela correta identificada com {max_linhas} linhas visíveis inicialmente!")
    
    # Rolagem incremental para capturar os registros
    try:
        tabela_container = tabela_selecionada.find_element(By.XPATH, '../../div[2]')
    except NoSuchElementException:
        print(f"❌ Container da tabela não encontrado para a reserva: {reserva}")
        return None
    
    driver.execute_script("arguments[0].scrollTop = 0", tabela_container)
    time.sleep(1)
    
    dados_tabela = []
    chaves_linha = set()  # Para evitar duplicatas
    
    while True:
        linhas_visiveis = tabela_selecionada.find_elements(By.XPATH, './/div[contains(@class, "ui-grid-row")]')
        for linha in linhas_visiveis:
            colunas = linha.find_elements(By.XPATH, ".//div[contains(@class, 'ui-grid-cell-contents')]")
            dados_linha = [coluna.text.strip() for coluna in colunas]
            chave = "|".join(dados_linha)
            if chave not in chaves_linha and len(dados_linha) >= 5:
                dados_tabela.append(dados_linha)
                chaves_linha.add(chave)
                
        scroll_top = driver.execute_script("return arguments[0].scrollTop", tabela_container)
        client_height = driver.execute_script("return arguments[0].clientHeight", tabela_container)
        scroll_height = driver.execute_script("return arguments[0].scrollHeight", tabela_container)
        
        if scroll_top + client_height >= scroll_height:
            print("📜 Chegou ao final do container de rolagem.")
            break
        else:
            driver.execute_script("arguments[0].scrollTop = arguments[0].scrollTop + 300", tabela_container)
            time.sleep(1)
    
    # Última verificação para capturar possíveis linhas restantes
    driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", tabela_container)
    time.sleep(2)
    linhas_visiveis = tabela_selecionada.find_elements(By.XPATH, './/div[contains(@class, "ui-grid-row")]')
    for linha in linhas_visiveis:
        colunas = linha.find_elements(By.XPATH, ".//div[contains(@class, 'ui-grid-cell-contents')]")
        dados_linha = [coluna.text.strip() for coluna in colunas]
        chave = "|".join(dados_linha)
        if chave not in chaves_linha and len(dados_linha) >= 5:
            dados_tabela.append(dados_linha)
            chaves_linha.add(chave)
            
    print(f"✅ Extração concluída para reserva {reserva}! Total de linhas extraídas: {len(dados_tabela)}")
    
    # Inserir a reserva como primeira coluna de cada linha extraída
    for row in dados_tabela:
        row.insert(0, reserva)
    
    return dados_tabela

def main():
    driver, wait = iniciar_driver()
    usuario = "redex.gate03@itracker.com.br"  # Considere obter via variável de ambiente
    senha = "123"  # Considere obter via variável de ambiente
    
    try:
        realizar_login(driver, wait, usuario, senha)
        navegar_menu(wait)
        
        # Ler a planilha de reservas
        excel_reservas = r"C:\Users\leonardo.fragoso\Desktop\Projetos\Depot-Project\Booking_List.xlsx"
        df_reservas = pd.read_excel(excel_reservas)
        lista_reservas = df_reservas["Booking"].tolist()
        
        todos_dados = []
        for reserva in lista_reservas:
            if pd.isna(reserva):
                print("⚠️ Valor de reserva inválido (NaN) encontrado, ignorando...")
                continue
            reserva_str = str(reserva).strip()
            print(f"\n🔍 Iniciando consulta para a reserva: {reserva_str}")
            dados = processar_reserva(driver, wait, reserva_str)
            if dados:
                todos_dados.extend(dados)
            # Limpar o campo de busca para a próxima reserva
            try:
                campo_busca = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="opcoesBusca"]/div[4]/div/input')))
                campo_busca.click()
                campo_busca.clear()
            except TimeoutException:
                print("❌ Erro ao limpar o campo de busca.")
            time.sleep(1)
        
        # Consolidar e salvar os dados extraídos em uma planilha
        colunas_tabela = ["Reserva", "Id", "Placa", "Condutor", "Reboque 1", "Reboque 2", "Nro. Eixos", "Mercadoria", "Encerrado"]
        df_saida = pd.DataFrame(todos_dados, columns=colunas_tabela)
        output_excel_path = r"C:\Users\leonardo.fragoso\Desktop\Projetos\Depot-Project\Dados_Extraidos_Corrigido.xlsx"
        df_saida.to_excel(output_excel_path, index=False)
        print(f"\n✅ Arquivo Excel consolidado salvo em: {output_excel_path}")
    finally:
        driver.quit()
        print("✅ Navegador fechado.")

if __name__ == '__main__':
    main()
