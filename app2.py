import time
import re
import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from datetime import datetime, timedelta


def iniciar_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 20)
    return driver, wait

def realizar_login(driver, wait, usuario, senha):
    driver.get("https://portaldeservicos.riobrasilterminal.com/tosp/Workspace/load#/CentralCeR")
    print("🔍 Aguardando carregamento da página de login...")
    try:
        campo_usuario = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="username"]')))
        campo_usuario.send_keys(usuario)
        driver.find_element(By.XPATH, '//*[@id="pass"]').send_keys(senha)
        driver.find_element(By.XPATH, '/html/body/div[1]/div/div/form/table/tbody/tr[2]/td[1]/button').click()
        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="bs-example-navbar-collapse-1"]/ul/li[2]/a/span')))
        print("✅ Login realizado com sucesso!")
    except TimeoutException:
        print("❌ Timeout durante o login. Verifique se a página ou os XPaths mudaram.")
        driver.quit()

def navegar_menu(wait):
    print("🔍 Navegando no menu...")
    menu_xpath = '//*[@id="bs-example-navbar-collapse-1"]/ul/li[2]/a/span'
    submenu_xpath = '//*[@id="bs-example-navbar-collapse-1"]/ul/li[2]/ul/li[1]/a/span'
    try:
        wait.until(EC.element_to_be_clickable((By.XPATH, menu_xpath))).click()
        time.sleep(1)
        wait.until(EC.element_to_be_clickable((By.XPATH, submenu_xpath))).click()
        time.sleep(2)
        print("✅ Menu navegado com sucesso!")
    except TimeoutException:
        print("❌ Erro ao navegar no menu. Verifique os XPaths.")

def abrir_segunda_aba(driver, url):
    driver.execute_script(f"window.open('{url}', '_blank');")
    print("✅ Segunda aba aberta com:", url)

# Variável global para indicar se o botão já foi clicado
botao_dialogo_clicado = False

def obter_dado_da_segunda_aba(driver, wait, tipo, primeira_vez=False):
    global botao_dialogo_clicado
    driver.switch_to.window(driver.window_handles[1])
    print("🔍 Você está na segunda aba (Looker Studio).")

    # 🔹 Se for a primeira vez acessando a URL, procurar pelo botão e clicar
    if primeira_vez and not botao_dialogo_clicado:
        try:
            botao_dialog_xpath = '/html/body/div[4]/md-dialog/md-dialog-actions/button[2]'
            botao_dialog = wait.until(EC.element_to_be_clickable((By.XPATH, botao_dialog_xpath)))
            botao_dialog.click()
            botao_dialogo_clicado = True  # Marcar que o botão já foi tratado
            print("✅ Botão 'Fechar' clicado com sucesso na primeira execução.")
        except TimeoutException:
            print("⚠️ Botão 'Fechar' não encontrado. Talvez já tenha sido fechado anteriormente.")

    time.sleep(3)

    # 🔹 XPath do dado a ser extraído
    if tipo == "exportacao":
        xpath_dado = '//*[@id="body"]/div[2]/div/ng2-reporting-plate/plate/div/div/div/div[1]/div[1]/div[2]/div/div/div/canvas-pancake-adapter/canvas-layout/div/div/div/div/div/div/ng2-report/ng2-canvas-container/div/div[2]/ng2-canvas-component/div/div/div/div/table-wrapper/div/ng2-table/div/div[3]/div[2]/div[2]/div[4]/span'
    else:
        xpath_dado = '//*[@id="body"]/div[2]/div/ng2-reporting-plate/plate/div/div/div/div[1]/div[1]/div[2]/div/div/div/canvas-pancake-adapter/canvas-layout/div/div/div/div/div/div/ng2-report/ng2-canvas-container/div/div[1]/ng2-canvas-component/div/div/div/div/table-wrapper/div/ng2-table/div/div[3]/div[2]/div[2]/div[4]/span'
    
    try:
        elemento_dado = wait.until(EC.visibility_of_element_located((By.XPATH, xpath_dado)))
        texto_dado = elemento_dado.text.strip()
        dado_filtrado = re.sub(r'[^A-Za-z0-9]', '', texto_dado)
        print("✅ Dado extraído e filtrado:", dado_filtrado)
        return dado_filtrado
    except TimeoutException:
        print("❌ Timeout ao tentar extrair o dado da segunda aba.")
        return None


def realizar_consulta_primeira_aba(driver, wait, dado, tipo):
    driver.switch_to.window(driver.window_handles[0])
    print("🔍 Retornando à primeira aba para realizar a consulta com o dado:", dado)
    
    if tipo == "exportacao":
        campo_busca_xpath = '//*[@id="opcoesBusca"]/div[4]/div/input'
    else:
        campo_busca_xpath = '//*[@id="opcoesBusca"]/div[8]/div/input'
    
    try:
        campo_busca = wait.until(EC.element_to_be_clickable((By.XPATH, campo_busca_xpath)))
        campo_busca.clear()
        campo_busca.send_keys(dado)
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="opcoesBusca"]/div[11]/div[2]/button'))).click()
        print("✅ Consulta realizada com o dado:", dado)
        time.sleep(2)

        # Verifica se a mensagem "Nenhum registro encontrado." está presente
        try:
            mensagem_xpath = '//*[@id="divHint"]/table/tbody/tr[2]/td/div'
            wait.until(EC.presence_of_element_located((By.XPATH, mensagem_xpath)))
            print("⚠️ Nenhum registro encontrado para o dado:", dado)
            return False  # Retorna False para indicar que não há dados a serem extraídos
        except TimeoutException:
            pass  # Se a mensagem não for encontrada, continua normalmente
        
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="liberacao"]/div[8]'))).click()
        print("✅ Botão 'Reservar Janela' clicado com sucesso!")
        time.sleep(2)
        tabela_xpath = "//*[contains(@id, '-grid-container')]/div[2]/div/div/div/div[3]/div"
        wait.until(EC.element_to_be_clickable((By.XPATH, tabela_xpath))).click()
        print("✅ Linha da tabela selecionada!")
        time.sleep(2)
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="#conteudo"]/div/div[4]/button[3]'))).click()
        print("✅ Botão (+) clicado com sucesso!")
        return True  # Indica que os dados foram encontrados e processados corretamente

    except TimeoutException:
        print("❌ Erro ao realizar a consulta na primeira aba.")
        return False  # Retorna False se houver erro


def selecionar_tipo_container(driver, wait):
    tipo_xpath = '//*[@id="manutenirCadastroReserva"]/div/div/div[2]/div[4]/div[1]/select'
    
    try:
        campo_tipo = wait.until(EC.element_to_be_clickable((By.XPATH, tipo_xpath)))
        campo_tipo.click()
        campo_tipo.send_keys(Keys.ARROW_DOWN)
        time.sleep(1)
        campo_tipo.send_keys(Keys.ENTER)
        print("✅ Tipo 'Container Cheio' selecionado com sucesso!")
    except TimeoutException:
        print("❌ Erro ao selecionar o tipo de container.")

def selecionar_area(driver, wait):
    area_xpath = '//*[@id="area"]'
    
    try:
        campo_area = wait.until(EC.element_to_be_clickable((By.XPATH, area_xpath)))
        campo_area.click()
        campo_area.send_keys(Keys.ARROW_DOWN)
        time.sleep(1)
        campo_area.send_keys(Keys.ARROW_DOWN)
        time.sleep(1)
        campo_area.send_keys(Keys.ENTER)
        print("✅ Área 'Primária' selecionada com sucesso!")
    except TimeoutException:
        print("❌ Erro ao selecionar a área.")

def inserir_data_hoje(driver, wait):
    data_xpath = '//*[@id="entidadedia"]'
    data_hoje = datetime.today().strftime('%d/%m/%Y')
    
    try:
        campo_data = wait.until(EC.element_to_be_clickable((By.XPATH, data_xpath)))
        campo_data.click()
        campo_data.clear()
        campo_data.send_keys(data_hoje)
        print(f"✅ Data de hoje '{data_hoje}' inserida com sucesso!")
    except TimeoutException:
        print("❌ Erro ao inserir a data de hoje.")

def clicar_botao_laranja(driver, wait):
    botao_laranja_xpath = '//*[@id="manutenirCadastroReserva"]/div/div/div[2]/div[5]/div[2]/button'
    
    try:
        botao_laranja = wait.until(EC.element_to_be_clickable((By.XPATH, botao_laranja_xpath)))
        botao_laranja.click()
        print("✅ Botão laranja clicado com sucesso!")
    except TimeoutException:
        print("❌ Erro ao clicar no botão laranja.")

def clicar_janela_dia(driver, wait):
    janela_dia_xpath = '//*[@id="janelaDia"]'
    
    try:
        campo_janela_dia = wait.until(EC.element_to_be_clickable((By.XPATH, janela_dia_xpath)))
        campo_janela_dia.click()
        print("✅ Janela Dia clicada com sucesso!")
    except TimeoutException:
        print("❌ Erro ao clicar em Janela Dia.")

def extrair_informacoes_janela(driver, wait):
    janela_xpath = '//*[@id="janelaDia"]'  # XPath do <select> que contém as opções
    informacoes_janelas = []

    try:
        # Aguarda o elemento do dropdown estar presente no DOM
        campo_janela = wait.until(EC.presence_of_element_located((By.XPATH, janela_xpath)))
        
        # Clica no dropdown para garantir que está ativo
        campo_janela.click()
        time.sleep(2)  # Aguarda para garantir que as opções aparecem

        # Encontra todas as opções dentro do <select>
        opcoes = campo_janela.find_elements(By.TAG_NAME, "option")

        for opcao in opcoes:
            texto_opcao = opcao.text.strip()
            if texto_opcao:  # Garante que não está pegando opções vazias
                informacoes_janelas.append(texto_opcao)
                print(f"✅ Opção extraída: {texto_opcao}")

        if not informacoes_janelas:
            print("⚠️ Nenhuma opção encontrada no dropdown!")

        return informacoes_janelas

    except TimeoutException:
        print("❌ Erro ao tentar extrair as informações da janela. Verifique se o XPath está correto.")
        return []

def formatar_dados_janelas(dados_janelas):
    dados_formatados = []
    pattern = r"(.+?)\s*\((\d{2}:\d{2})\s*-\s*(\d{2}:\d{2})\)\s*\[QTD:\s*(\d+)\]"
    
    for dado in dados_janelas:
        match = re.search(pattern, dado)
        if match:
            nome_janela = match.group(1).strip()
            hora_inicial = match.group(2)
            hora_final = match.group(3)
            qtd_veiculos = match.group(4)
            dados_formatados.append([nome_janela, hora_inicial, hora_final, qtd_veiculos])

    return dados_formatados

def inserir_data(driver, wait, dias_a_frente=0):
    data_xpath = '//*[@id="entidadedia"]'
    data_desejada = (datetime.today() + timedelta(days=dias_a_frente)).strftime('%d/%m/%Y')
    
    try:
        campo_data = wait.until(EC.element_to_be_clickable((By.XPATH, data_xpath)))
        campo_data.click()
        campo_data.clear()
        campo_data.send_keys(data_desejada)
        print(f"✅ Data '{data_desejada}' inserida com sucesso!")
    except TimeoutException:
        print(f"❌ Erro ao inserir a data '{data_desejada}'.")

def salvar_dados_janelas(dados_janelas, data_consulta, arquivo_output):
    df = pd.DataFrame(dados_janelas, columns=["Janela", "Hora Inicial", "Hora Final", "Qtd Veículos Reservados"])
    df.insert(0, "Dia", data_consulta)  # Adiciona a coluna "Dia" no índice 0

    # Se o arquivo já existir, carrega os dados e adiciona novas informações
    if os.path.exists(arquivo_output):
        df_existente = pd.read_excel(arquivo_output)
        df = pd.concat([df_existente, df], ignore_index=True)

    df.to_excel(arquivo_output, index=False)
    print(f"✅ Dados salvos na planilha {arquivo_output}.")

def loop_de_extracao(driver, wait, tipo):
    driver.switch_to.window(driver.window_handles[1])  # Garantir que estamos na segunda aba
    print("🔍 Começando o loop de extração na segunda aba.")

    primeira_execucao = True  # Variável de controle para clicar no botão "Fechar" apenas uma vez

    while True:
        # Obter o dado de cada registro da segunda aba
        dado = obter_dado_da_segunda_aba(driver, wait, tipo, primeira_vez=primeira_execucao)
        primeira_execucao = False  # Depois da primeira execução, não procuramos mais o botão "Fechar"

        if not dado:
            break  # Se não encontrar dados, termina o loop

        # Realizar consulta na primeira aba com o dado extraído
        dados_encontrados = realizar_consulta_primeira_aba(driver, wait, dado, tipo)
        if not dados_encontrados:
            print(f"➡️ Pulando {dado}, pois não há registros disponíveis.")
            continue  # Pula para o próximo dado sem tentar reservar janelas

        # 🔹 Processando informações do dia atual
        selecionar_tipo_container(driver, wait)
        selecionar_area(driver, wait)
        inserir_data(driver, wait, dias_a_frente=0)
        clicar_botao_laranja(driver, wait)
        clicar_janela_dia(driver, wait)
        dados_janelas_hoje = extrair_informacoes_janela(driver, wait)

        if dados_janelas_hoje:
            dados_formatados_hoje = formatar_dados_janelas(dados_janelas_hoje)
            salvar_dados_janelas(dados_formatados_hoje, datetime.today().strftime('%d/%m/%Y'), "informacoes_janelas.xlsx")

        # 🔹 Processando informações do dia seguinte
        inserir_data(driver, wait, dias_a_frente=1)
        clicar_botao_laranja(driver, wait)
        clicar_janela_dia(driver, wait)
        dados_janelas_amanha = extrair_informacoes_janela(driver, wait)

        if dados_janelas_amanha:
            dados_formatados_amanha = formatar_dados_janelas(dados_janelas_amanha)
            salvar_dados_janelas(dados_formatados_amanha, (datetime.today() + timedelta(days=1)).strftime('%d/%m/%Y'), "informacoes_janelas.xlsx")

            # 🔹 Após extrair e salvar os dados do dia seguinte, clicar no botão "Cancelar"
            try:
                botao_cancelar_xpath = '//*[@id="manutenirCadastroReserva"]/div/div/div[2]/div[7]/button[2]'
                botao_cancelar = wait.until(EC.element_to_be_clickable((By.XPATH, botao_cancelar_xpath)))
                botao_cancelar.click()
                print("✅ Botão 'Cancelar' clicado com sucesso após a extração do dia seguinte.")
            except TimeoutException:
                print("❌ Erro ao clicar no botão 'Cancelar'. Verifique o XPath.")

        # 🔹 Voltar para a segunda aba para extrair o próximo registro
        driver.switch_to.window(driver.window_handles[1])
        print("🔄 Voltando para a segunda aba para extrair o próximo registro.")

        try:
            seta_xpath = '//*[@id="body"]/div[2]/div/ng2-reporting-plate/plate/div/div/div/div[1]/div[1]/div[2]/div/div/div/canvas-pancake-adapter/canvas-layout/div/div/div/div/div/div/ng2-report/ng2-canvas-container/div/div[2]/ng2-canvas-component/div/div/div/div/table-wrapper/div/ng2-table/div/div[6]/div[3]'
            seta = wait.until(EC.element_to_be_clickable((By.XPATH, seta_xpath)))
            seta.click()
            print("✅ Navegando para o próximo registro.")
            time.sleep(2)  # Atraso para garantir que o próximo registro seja carregado
        except TimeoutException:
            print("⚠️ Não foi possível encontrar a seta para navegar para o próximo registro.")
            break  # Se não encontrar a seta, encerra o loop



def main():
    driver, wait = iniciar_driver()
    usuario = "redex.gate03@itracker.com.br"
    senha = "123"
    tipo = "exportacao"
    arquivo_saida = "informacoes_janelas.xlsx"

    try:
        realizar_login(driver, wait, usuario, senha)
        navegar_menu(wait)
        abrir_segunda_aba(driver, "https://lookerstudio.google.com/u/0/reporting/55ec93e4-3114-46d5-9125-79d7191b1c0a/page/p_5jlrz7xapd")
        
        # Chama o loop para extrair dados de todos os registros
        loop_de_extracao(driver, wait, tipo)
        
        print("✅ Processo concluído e dados salvos com datas exatas.")

    except Exception as e:
        print(f"❌ Erro inesperado: {e}")
    finally:
        input("Pressione Enter para fechar o navegador...")
        driver.quit()
        print("✅ Navegador fechado.")

if __name__ == '__main__':
    main()