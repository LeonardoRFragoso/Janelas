#!/usr/bin/env python3
# export.py
import time
import re
import os
import pandas as pd
from datetime import datetime, timedelta
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

def verificar_dialogo(driver, wait):
    xpath = '/html/body/div[4]/md-dialog/md-dialog-actions/button[2]'
    try:
        botao = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        botao.click()
        print("✅ [Exportação] Botão de diálogo fechado com sucesso.")
    except TimeoutException:
        print("⚠️ [Exportação] Nenhum diálogo para fechar ou já foi fechado anteriormente.")

def obter_dado_da_segunda_aba(driver, wait, tipo, max_retries=3):
    """
    Extrai o valor de "DI / BOOKING / CTE" na segunda aba com retries.
    Caso não seja possível extrair o valor após max_retries tentativas,
    retorna "N/D" para indicar que o dado não está disponível.
    """
    # Garante que estamos na aba de extração
    driver.switch_to.window(driver.window_handles[1])
    print(f" [Exportação] Extraindo DI / BOOKING / CTE da segunda aba para {tipo.upper()}...")
    time.sleep(3)  # Aguarda que a tabela seja renderizada
    xpath = (
        '//*[@id="body"]/div[2]/div/ng2-reporting-plate/plate/div/div/div/div[1]/div[1]/div[2]/div/div/div/'
        'canvas-pancake-adapter/canvas-layout/div/div/div/div/div/div/ng2-report/ng2-canvas-container/div/div[2]/'
        'ng2-canvas-component/div/div/div/div/table-wrapper/div/ng2-table/div/div[3]/div[2]/div[2]/div[3]/span'
    )
    attempt = 0
    while attempt < max_retries:
        try:
            elemento = wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
            texto = elemento.text.strip()
            if texto:
                # Remove caracteres indesejados (mantém apenas letras e números)
                dado_filtrado = re.sub(r'[^A-Za-z0-9]', '', texto)
                print(f"✅ [Exportação] Dado extraído ({tipo}): {dado_filtrado}")
                return dado_filtrado
            else:
                print(f"⚠️ [Exportação] Tentativa {attempt+1}: Dado extraído está vazio. Tentando novamente...")
                attempt += 1
                time.sleep(2)
        except TimeoutException as e:
            print(f"❌ [Exportação] Tentativa {attempt+1}: Erro na extração - {e}. Tentando novamente...")
            attempt += 1
            time.sleep(2)
    print(f"❌ [Exportação] Não foi possível extrair o valor de DI / BOOKING / CTE para {tipo} após {max_retries} tentativas. Usando valor 'N/D'.")
    return "N/D"

def realizar_consulta_primeira_aba(driver, wait, dado, tipo):
    driver.switch_to.window(driver.window_handles[0])
    print(f" [Exportação] Retornando à primeira aba para consulta com o dado: {dado}")
    xpath_input = '//*[@id="opcoesBusca"]/div[4]/div/input'
    try:
        campo = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_input)))
        campo.clear()
        campo.send_keys(dado)
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="opcoesBusca"]/div[11]/div[2]/button'))).click()
        print("✅ [Exportação] Consulta realizada com o dado:", dado)
        time.sleep(2)
        try:
            xpath_msg = '//*[@id="divHint"]/table/tbody/tr[2]/td/div'
            wait.until(EC.presence_of_element_located((By.XPATH, xpath_msg)))
            print("⚠️ [Exportação] Nenhum registro encontrado para o dado:", dado)
            return False
        except TimeoutException:
            pass
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="liberacao"]/div[8]'))).click()
        print("✅ [Exportação] Botão 'Reservar Janela' clicado com sucesso!")
        time.sleep(2)
        xpath_tabela = "//*[contains(@id, '-grid-container')]/div[2]/div/div/div/div[3]/div"
        wait.until(EC.element_to_be_clickable((By.XPATH, xpath_tabela))).click()
        print("✅ [Exportação] Linha da tabela selecionada!")
        time.sleep(2)
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id=\"#conteudo\"]/div/div[4]/button[3]'))).click()
        print("✅ [Exportação] Botão (+) clicado com sucesso!")
        return True
    except TimeoutException:
        print("❌ [Exportação] Erro na consulta na primeira aba.")
        return False

def selecionar_tipo_container(driver, wait):
    xpath = '//*[@id="manutenirCadastroReserva"]/div/div/div[2]/div[4]/div[1]/select'
    try:
        campo = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        campo.click()
        time.sleep(1)
        campo.send_keys(Keys.ARROW_DOWN)
        time.sleep(1)
        campo.send_keys(Keys.ENTER)
        print("✅ [Exportação] Tipo 'Container Cheio' selecionado com sucesso!")
    except TimeoutException:
        print("❌ [Exportação] Erro ao selecionar o tipo de container.")

def selecionar_area(driver, wait):
    xpath = '//*[@id="area"]'
    try:
        campo = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        campo.click()
        time.sleep(1)
        campo.send_keys(Keys.ARROW_DOWN)
        time.sleep(1)
        campo.send_keys(Keys.ARROW_DOWN)
        time.sleep(1)
        campo.send_keys(Keys.ENTER)
        print("✅ [Exportação] Área 'Primária' selecionada com sucesso!")
    except TimeoutException:
        print("❌ [Exportação] Erro ao selecionar a área.")

def inserir_data_hoje(driver, wait):
    xpath = '//*[@id="entidadedia"]'
    data = datetime.today().strftime('%d/%m/%Y')
    try:
        campo = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        campo.click()
        campo.clear()
        campo.send_keys(data)
        print(f"✅ [Exportação] Data de hoje '{data}' inserida com sucesso!")
    except TimeoutException:
        print("❌ [Exportação] Erro ao inserir a data de hoje.")

def clicar_botao_laranja(driver, wait):
    xpath = '//*[@id="manutenirCadastroReserva"]/div/div/div[2]/div[5]/div[2]/button'
    try:
        botao = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        botao.click()
        print("✅ [Exportação] Botão laranja clicado com sucesso!")
    except TimeoutException:
        print("❌ [Exportação] Erro ao clicar no botão laranja.")

def clicar_janela_dia(driver, wait):
    xpath = '//*[@id="janelaDia"]'
    try:
        campo = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        campo.click()
        print("✅ [Exportação] Janela Dia clicada com sucesso!")
    except TimeoutException:
        print("❌ [Exportação] Erro ao clicar em Janela Dia.")

def extrair_informacoes_janela(driver, wait):
    xpath = '//*[@id="janelaDia"]'
    try:
        campo = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
        campo.click()
        time.sleep(2)
        opcoes = campo.find_elements(By.TAG_NAME, "option")
        opcoes_text = [opcao.text.strip() for opcao in opcoes if opcao.text.strip()]
        for opcao in opcoes_text:
            print("✅ [Exportação] Opção extraída:", opcao)
        if not opcoes_text:
            print("⚠️ [Exportação] Nenhuma opção encontrada no dropdown!")
        return opcoes_text
    except TimeoutException:
        print("❌ [Exportação] Erro ao extrair informações da janela. Verifique o XPath.")
        return []

def formatar_dados_janelas(dados):
    pattern = r"(.+?)\s*\((\d{2}:\d{2})\s*-\s*(\d{2}:\d{2})\)\s*\[QTD:\s*(\d+)\]"
    dados_formatados = []
    for dado in dados:
        match = re.search(pattern, dado)
        if match:
            nome = match.group(1).strip()
            h_ini = match.group(2)
            h_fin = match.group(3)
            qtd = match.group(4)
            dados_formatados.append([nome, h_ini, h_fin, qtd])
    return dados_formatados

def inserir_data(driver, wait, dias=0):
    xpath = '//*[@id="entidadedia"]'
    data = (datetime.today() + timedelta(days=dias)).strftime('%d/%m/%Y')
    try:
        campo = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        campo.click()
        campo.clear()
        campo.send_keys(data)
        print(f"✅ [Exportação] Data '{data}' inserida com sucesso!")
    except TimeoutException:
        print(f"❌ [Exportação] Erro ao inserir a data '{data}'.")

def salvar_dados_janelas(dados, data_consulta, arquivo, tipo, di_booking_cte):
    """
    Salva os dados extraídos na planilha, inserindo as colunas:
      0. DI / BOOKING / CTE
      1. Dia
      2. Tipo
      3. Janela
      4. Hora Inicial
      5. Hora Final
      6. Qtd Veículos Reservados
    """
    df = pd.DataFrame(dados, columns=["Janela", "Hora Inicial", "Hora Final", "Qtd Veículos Reservados"])
    df.insert(0, "Tipo", tipo)
    df.insert(0, "Dia", data_consulta)
    df.insert(0, "DI / BOOKING / CTE", di_booking_cte)
    if os.path.exists(arquivo):
        df_existente = pd.read_excel(arquivo)
        df = pd.concat([df_existente, df], ignore_index=True)
    df.to_excel(arquivo, index=False)
    print(f"✅ [Exportação] Dados salvos na planilha {arquivo} (Tipo: {tipo}).")

def avancar_para_proximo_registro(driver, wait, tipo):
    xpath = (
        '//*[@id="body"]/div[2]/div/ng2-reporting-plate/plate/div/div/div/div[1]/div[1]/'
        'div[2]/div/div/div/canvas-pancake-adapter/canvas-layout/div/div/div/div/div/'
        'div/ng2-report/ng2-canvas-container/div/div[2]/ng2-canvas-component/div/div/'
        'div/div/table-wrapper/div/ng2-table/div/div[6]/div[3]'
    )
    try:
        botao = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        botao.click()
        print("✅ [Exportação] Botão de avançar clicável. Avançando para o próximo registro.")
        time.sleep(5)
        return True
    except TimeoutException:
        print("⚠️ [Exportação] A seta para avançar não está disponível. Final da lista de exportação.")
        return False

def loop_de_extracao(driver, wait, tipo):
    print(f" Iniciando loop de extração para {tipo}...")
    # Certifica que está na aba de extração e fecha eventuais diálogos
    driver.switch_to.window(driver.window_handles[1])
    verificar_dialogo(driver, wait)
    
    while True:
        # Extrai o DI/BOOKING/CTE da segunda aba
        current_record = obter_dado_da_segunda_aba(driver, wait, tipo, max_retries=3)
        if not current_record or current_record == "N/D":
            print(f"⚠️ Nenhum dado extraído para {tipo}. Encerrando loop.")
            break

        # Volta para a primeira aba para a consulta
        driver.switch_to.window(driver.window_handles[0])
        sucesso = realizar_consulta_primeira_aba(driver, wait, current_record, tipo)
        if not sucesso:
            print(f"⚠️ Consulta na primeira aba falhou para {current_record}. Tentando próximo registro...")
            driver.switch_to.window(driver.window_handles[1])
            if not avancar_para_proximo_registro(driver, wait, tipo):
                break
            continue

        try:
            selecionar_tipo_container(driver, wait)
            selecionar_area(driver, wait)
            inserir_data_hoje(driver, wait)
            clicar_botao_laranja(driver, wait)
            clicar_janela_dia(driver, wait)

            # Tenta extrair as janelas disponíveis para hoje
            dados_janelas_hoje = extrair_informacoes_janela(driver, wait)
            if dados_janelas_hoje:
                dados_formatados = formatar_dados_janelas(dados_janelas_hoje)
                salvar_dados_janelas(
                    dados_formatados,
                    datetime.today().strftime('%d/%m/%Y'),
                    "informacoes_janelas.xlsx",
                    tipo,
                    current_record
                )
                # Extração para o dia seguinte (D+1)
                inserir_data(driver, wait, dias=1)
                clicar_botao_laranja(driver, wait)
                clicar_janela_dia(driver, wait)
                dados_janelas_amanha = extrair_informacoes_janela(driver, wait)
                if dados_janelas_amanha:
                    dados_formatados = formatar_dados_janelas(dados_janelas_amanha)
                    salvar_dados_janelas(
                        dados_formatados,
                        (datetime.today() + timedelta(days=1)).strftime('%d/%m/%Y'),
                        "informacoes_janelas.xlsx",
                        tipo,
                        current_record
                    )
                # Extração para D+2 (dois dias após hoje)
                inserir_data(driver, wait, dias=2)
                clicar_botao_laranja(driver, wait)
                clicar_janela_dia(driver, wait)
                dados_janelas_d2 = extrair_informacoes_janela(driver, wait)
                if dados_janelas_d2:
                    dados_formatados = formatar_dados_janelas(dados_janelas_d2)
                    salvar_dados_janelas(
                        dados_formatados,
                        (datetime.today() + timedelta(days=2)).strftime('%d/%m/%Y'),
                        "informacoes_janelas.xlsx",
                        tipo,
                        current_record
                    )
                try:
                    xpath_cancelar = '//*[@id="manutenirCadastroReserva"]/div/div/div[2]/div[7]/button[2]'
                    botao_cancelar = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_cancelar)))
                    botao_cancelar.click()
                    print("✅ [Exportação] Botão 'Cancelar' clicado com sucesso após a extração dos dias seguintes.")
                    time.sleep(2)
                except TimeoutException:
                    print("❌ [Exportação] Erro ao clicar no botão 'Cancelar'. Verifique o XPath.")
                
                print(f"✅ Extração realizada com sucesso para o registro {current_record}. Encerrando loop.")
                break  # Interrompe o loop assim que encontrar um registro com janelas disponíveis
            else:
                print(f"⚠️ [Exportação] Nenhuma janela disponível para {current_record}. Tentando próximo registro...")
        except Exception as e:
            print(f"⚠️ [Exportação] Erro ao processar o registro {current_record}: {e}")

        driver.switch_to.window(driver.window_handles[1])
        if not avancar_para_proximo_registro(driver, wait, tipo):
            print("⚠️ [Exportação] Não foi possível avançar para o próximo registro. Encerrando loop.")
            break

    print("✅ Loop de extração finalizado.")

def run_export(driver, wait):
    summary = loop_de_extracao(driver, wait, "exportacao")
    return summary

if __name__ == '__main__':
    # Este módulo deve ser chamado via run_export() a partir do app.py
    pass
