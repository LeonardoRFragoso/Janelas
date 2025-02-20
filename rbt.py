import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager

def main():
    # 1) CONFIGURAÇÕES DE DOWNLOAD
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

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    driver.maximize_window()

    try:
        # 2) LOGIN
        driver.get("https://tosp-azr.ictsirio.com/tosp/")
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="username"]'))
        )
        driver.find_element(By.XPATH, '//*[@id="username"]').send_keys("alexandre.moura")
        time.sleep(1)
        driver.find_element(By.XPATH, '//*[@id="pass"]').send_keys("486cf3", Keys.RETURN)

        # 3) ACESSAR A URL DO DASHBOARD
        driver.get("https://tosp-azr.ictsirio.com/tosp/Workspace/load#/ManutenirMonitorOperacional")
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="maincontent"]'))
        )

        # 4) CLICAR NO BOTÃO DA LUPA
        lupa_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="maincontent"]/div[1]/div/div[1]/div/form/div/fieldset/table/tbody/tr/td[2]/button'))
        )
        lupa_button.click()

        # 5) LOCALIZAR E CLICAR EM "JANELAS DE AGENDAMENTO"
        wait = WebDriverWait(driver, 20)
        container = wait.until(
            EC.presence_of_element_located((By.XPATH, '//*[contains(@id, "-grid-container")]/div[2]'))
        )
        found = False
        for _ in range(20):
            try:
                janelas_agendamento = container.find_element(By.XPATH, './/div[contains(text(), "JANELAS DE AGENDAMENTO")]')
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", janelas_agendamento)
                janelas_agendamento.click()
                found = True
                break
            except Exception:
                driver.execute_script("arguments[0].scrollTop += 300;", container)

        if not found:
            print("Elemento 'JANELAS DE AGENDAMENTO' não encontrado após rolagem.")
            return

        # 6) CLICAR NO BOTÃO "SETAS"
        setas_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="maincontent"]/div[1]/div/div[1]/div/div/fieldset/div[1]/button[1]'))
        )
        setas_button.click()

        # 7) ESPERAR O MODAL ABRIR
        modal_xpath = '//*[@id="visualizarMonitorOperacional"]/div/div'
        WebDriverWait(driver, 30).until(
            EC.visibility_of_element_located((By.XPATH, modal_xpath))
        )
        print("Modal do Power BI aberto!")

        # 8) DENTRO DO MODAL, LOCALIZAR O IFRAME
        report_iframe = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#report-container iframe"))
        )
        # 9) TROCAR PARA O CONTEXTO DO IFRAME
        driver.switch_to.frame(report_iframe)
        print("Trocamos para o contexto do iframe do Power BI.")

        # 10) ESPERAR O POWER BI RENDERIZAR
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, "pvExplorationHost"))
        )
        print("Conteúdo do Power BI carregado dentro do iframe!")

        # 11) LOCALIZAR TODAS AS CÉLULAS DA TABELA
        cells_xpath = (
            '//*[@id="pvExplorationHost"]//visual-container-repeat/visual-container[7]'
            '//visual-modern//div/div/div[2]/div[1]/div[2]/div/div[*]/div/div/div[*]'
        )
        cells = WebDriverWait(driver, 20).until(
            EC.presence_of_all_elements_located((By.XPATH, cells_xpath))
        )
        print(f"Foram encontradas {len(cells)} células na tabela do Power BI.")

        # 12) ESCOLHER A PRIMEIRA CÉLULA (OU OUTRA)
        if not cells:
            print("Nenhuma célula encontrada na tabela.")
            return

        first_cell = cells[0]
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", first_cell)
        time.sleep(0.5)
        ActionChains(driver).move_to_element(first_cell).perform()
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", first_cell)
        print(f"Primeira célula clicada: {first_cell.text.strip()}")
        time.sleep(1)

        # 13) LOCALIZAR O BOTÃO "..." PELO XPATH FIXO
        ellipsis_button_xpath = (
            '//*[@id="pvExplorationHost"]/div/div/exploration/div/explore-canvas/div/div[2]/div/div[2]/'
            'div[2]/visual-container-repeat/visual-container[7]/transform/div/visual-container-header/'
            'div/div/div/visual-container-options-menu/visual-header-item-container/div/button'
        )
        try:
            ellipsis_button = driver.find_element(By.XPATH, ellipsis_button_xpath)
            if ellipsis_button.is_displayed():
                print("Botão '...' localizado pelo XPath fixo.")
            else:
                print("Botão '...' não está visível, mesmo com o XPath fixo.")
                return
        except Exception:
            print("Não foi possível localizar o botão '...' usando o XPath fixo.")
            return

        # 14) CLICAR NO BOTÃO "..."
        driver.execute_script("arguments[0].scrollIntoView(true);", ellipsis_button)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", ellipsis_button)
        print("Botão '...' clicado com sucesso!")

        # 14.1) CLICAR EM "EXPORTAR DADOS" (XPath: //*[@id="0"])
        exportar_dados_xpath = '//*[@id="0"]'
        try:
            exportar_dados_button = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, exportar_dados_xpath))
            )
            driver.execute_script("arguments[0].click();", exportar_dados_button)
            print("Opção 'Exportar Dados' clicada com sucesso!")
        except Exception as e:
            print("Erro ao clicar em 'Exportar Dados':", e)
            return

        # 15) CLICAR EM "DADOS RESUMIDOS"
        dados_resumidos_xpath = '//*[@id="pbi-radio-button-1"]/label/section/span'
        try:
            dados_resumidos_button = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, dados_resumidos_xpath))
            )
            driver.execute_script("arguments[0].click();", dados_resumidos_button)
            print("Opção 'Dados resumidos' clicada com sucesso!")
        except Exception as e:
            print("Erro ao clicar em 'Dados resumidos':", e)

        # 16) CLICAR NO BOTÃO "EXPORTAR"
        exportar_button_xpath = '//*[@id="mat-mdc-dialog-0"]/div/div/export-data-dialog/mat-dialog-actions/button[1]'
        try:
            exportar_button = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, exportar_button_xpath))
            )
            driver.execute_script("arguments[0].click();", exportar_button)
            print("Botão 'Exportar' clicado com sucesso!")
        except Exception as e:
            print("Erro ao clicar no botão 'Exportar':", e)

        # 17) AGUARDAR O DOWNLOAD DO ARQUIVO .XLSX
        downloaded_file = os.path.join(download_dir, "data.xlsx")
        timeout = 60  # segundos
        elapsed = 0
        while not os.path.exists(downloaded_file) and elapsed < timeout:
            time.sleep(1)
            elapsed += 1

        if os.path.exists(downloaded_file):
            print("Download concluído:", downloaded_file)
        else:
            print(f"Timeout: arquivo data.xlsx não foi baixado em {timeout} segundos.")

    finally:
        driver.quit()

if __name__ == "__main__":
    main()
