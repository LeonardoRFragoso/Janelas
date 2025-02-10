import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time
import logging
from dataclasses import dataclass
from typing import List
from selenium.webdriver.common.keys import Keys

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

@dataclass
class ScraperConfig:
    url: str = "https://portaldeservicos.riobrasilterminal.com/tosp/Workspace/load#/CentralCeR"
    username: str = "redex.gate03@itracker.com.br"
    password: str = "123"
    output_path: str = r"C:\\Users\\leona\\OneDrive\\Documentos\\DEPOT-PROJECT\\Dados_Extraidos_Corrigido.xlsx"
    wait_time: int = 15
    scroll_pause_time: int = 2  

class PortalScraper:
    def __init__(self, config: ScraperConfig):
        self.config = config
        self.setup_driver()
        self.wait = WebDriverWait(self.driver, config.wait_time)

    def setup_driver(self):
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_experimental_option("excludeSwitches", ["enable-logging"])
        self.driver = webdriver.Chrome(options=options)

    def wait_and_click(self, xpath: str) -> None:
        try:
            element = self.wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            element.click()
        except TimeoutException:
            logger.error(f"Timeout ao esperar elemento: {xpath}")
            raise

    def wait_and_send_keys(self, xpath: str, keys: str) -> None:
        try:
            element = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
            element.send_keys(keys)
        except TimeoutException:
            logger.error(f"Timeout ao enviar teclas para: {xpath}")
            raise

    def login(self) -> None:
        logger.info("Iniciando processo de login...")
        self.driver.get(self.config.url)
        self.wait_and_send_keys('//*[@id="username"]', self.config.username)
        self.wait_and_send_keys('//*[@id="pass"]', self.config.password)
        self.wait_and_click('/html/body/div[1]/div/div/form/table/tbody/tr[2]/td[1]/button')
        time.sleep(3)
        logger.info("Login realizado com sucesso")

    def navigate_to_table(self, valor_busca: str) -> None:
        logger.info("Navegando até a tabela...")
        self.wait_and_click('//*[@id="bs-example-navbar-collapse-1"]/ul/li[2]/a/span')
        time.sleep(2)
        self.wait_and_click('//*[@id="bs-example-navbar-collapse-1"]/ul/li[2]/ul/li[1]/a/span')
        time.sleep(3)
        self.wait_and_send_keys('//*[@id="opcoesBusca"]/div[4]/div/input', valor_busca)
        time.sleep(2)
        self.wait_and_click('//*[@id="opcoesBusca"]/div[11]/div[2]/button')
        time.sleep(3)
        self.wait_and_click('//*[@id="liberacao"]/div[7]/a')
        time.sleep(3)
        self.wait_and_click('//*[@id="#conteudo"]/div/tabset[2]/div/ul/li[4]/a')
        time.sleep(5)
        logger.info("Navegação concluída")

    def extract_table_data(self) -> List[List[str]]:
        """Extrai todos os registros da tabela garantindo que a rolagem funcione corretamente."""
        logger.info("Iniciando extração de dados com scroll dinâmico...")

        registros_unicos = set()
        dados_tabela = []

        rows_xpath = "//*[@id='1738427647414-grid-container']/div[2]/div/div"
        
        max_attempts = 20
        attempts = 0
        previous_count = 0

        try:
            scroll_container = self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//*[@id='1738427647414-grid-container']"))
            )
            logger.info("Contêiner de rolagem encontrado!")
        except TimeoutException:
            logger.error("Não foi possível encontrar o contêiner de rolagem da tabela!")
            return []

        while attempts < max_attempts:
            rows = self.driver.find_elements(By.XPATH, rows_xpath)
            logger.info(f"Linhas visíveis encontradas: {len(rows)}")

            for row in rows:
                try:
                    cells = row.find_elements(By.XPATH, ".//div[contains(@class, 'ui-grid-cell-contents')]")
                    if len(cells) >= 8:
                        row_data = [cell.text.strip() for cell in cells[:8]]
                        if row_data[0] and row_data[0] not in registros_unicos:
                            registros_unicos.add(row_data[0])
                            dados_tabela.append(row_data)
                            logger.info(f"Linha extraída: {row_data}")
                except Exception as e:
                    logger.error(f"Erro ao processar linha: {e}")

            current_count = len(dados_tabela)
            logger.info(f"Registros acumulados: {current_count}")

            if current_count == previous_count:
                attempts += 1
                logger.info(f"Nenhum novo registro encontrado, tentativa {attempts} de {max_attempts}.")
                self.driver.execute_script("arguments[0].scrollTop += arguments[0].offsetHeight;", scroll_container)
                time.sleep(2)
                if attempts % 3 == 0:
                    self.driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.PAGE_DOWN)
                    time.sleep(2)
                    logger.info("Pressionando 'Page Down' para tentar carregar mais registros.")
            else:
                attempts = 0
            previous_count = current_count

        logger.info(f"Extração finalizada. Total de registros: {len(dados_tabela)}")
        return dados_tabela

    def save_to_excel(self, dados: List[List[str]]) -> None:
        if not dados:
            logger.error("Nenhum dado para salvar")
            return
        df = pd.DataFrame(dados, columns=["Id", "Placa", "Condutor", "Reboque 1", "Reboque 2", "Nro. Eixos", "Mercadoria", "Encerrado"])
        df.to_excel(self.config.output_path, index=False)
        logger.info(f"Dados salvos em: {self.config.output_path}")

    def run(self, valor_busca: str) -> None:
        try:
            self.login()
            self.navigate_to_table(valor_busca)
            dados = self.extract_table_data()
            self.save_to_excel(dados)
        except Exception as e:
            logger.error(f"Erro durante a execução: {str(e)}")
            raise
        finally:
            self.driver.quit()
            logger.info("Navegador fechado")

def main():
    config = ScraperConfig()
    scraper = PortalScraper(config)
    scraper.run("33616194")

if __name__ == "__main__":
    main()
