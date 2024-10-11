import os
import time
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException

# Carregando as variáveis do arquivo .env
load_dotenv()

class Bot:
    """Classe responsável pelas principais funções do robô de automação de tarefas.

    Atributos:
        __email (str): Email utilizado para autenticação, obtido das variáveis de ambiente.
        __senha (str): Senha utilizada para autenticação, obtida das variáveis de ambiente.
        __email_teste (str): Email de teste para autenticação, obtido das variáveis de ambiente.
        __senha_teste (str): Senha de teste para autenticação, obtida das variáveis de ambiente.
        __url_sales (str): URL da página de vendas, obtida das variáveis de ambiente.
        __url_tracking (str): URL da página de rastreamento, obtida das variáveis de ambiente.
        driver (WebDriver): Instância do Selenium WebDriver para controlar o navegador (inicialmente None).
    """
    def __init__(self):
        self.__email = os.getenv('EMAIL')
        self.__senha = os.getenv('SENHA')
        self.__email_teste = os.getenv('EMAIL_TESTE')
        self.__senha_teste = os.getenv('SENHA_TESTE')
        self.__url_sales = os.getenv('URL_SALES')
        self.__url_tracking = os.getenv('URL_TRACKING')
        self.driver = None
    
    def set_driver(self):
        # Inicializando uma instância do Google Chrome
        service = Service()

        # Definindo a preferência para o brower do Chrome
        options = webdriver.ChromeOptions()

        # Iniciando a instância do Chrome com o service e options já definidos
        self.driver = webdriver.Chrome(service=service, options=options)
        return self.driver
    
    def set_wait(self, driver):
        driver.implicitly_wait(10)
        wait = WebDriverWait(driver, timeout=60)
        return wait
              
    def open_window(self, driver):
        """ Função responsável por abrir as duas abas principais: Sales Order e Tracking
        """
        driver.get(self.__url_sales)

        driver.execute_script("window.open('');") 
        driver.switch_to.window(driver.window_handles[1]) 
        driver.get(self.__url_tracking)

        driver.switch_to.window(driver.window_handles[0])         
    
    def close_cookie(self, driver: webdriver):
        """ Função responsável por clicar no botão de aceite dos Cookies necessários para prosseguirmos com a automação.

        Args:
            driver (WebDriver): Instância do Selenium WebDriver utilizada para interagir com a página.
        """
        try:
            # Localizando o botão de aceitação de cookies e clicando nele, caso exista
            button_cookie = driver.find_element(By.ID, 'onetrust-accept-btn-handler')
            button_cookie.click()
        except NoSuchElementException:
            # Caso o botão ja tenha sido clicado, ou não exista, seguiremos com a automação.
            print("Botão de aceitação de cookies não encontrado.")
        except ElementClickInterceptedException:
            print("A sobreposição de cookies estava bloqueando o clique.")
    
    def community_login(self, driver: webdriver, wait: WebDriverWait):
        """ Função responsável pelo login inicial, este login deve ser feito antes de iniciarmos a automação, deve ser utilizado o e-mail pessoal, caso não tenha uma conta será necessário a criação.

        Args:
            driver (WebDriver): Instância do Selenium WebDriver utilizada para interagir com a página.
            wait (WebDriverWait): Instância do Selenium WebDriverWait usada para aguardar elementos na página.
        """

        button_community = driver.find_element(By.ID, 'button_modal-login-btn__iPh6x')
        button_community.click()

        # Inserindo e-mail para login na plataforma
        e_mail_community = driver.find_element(By.ID, '43:2;a')
        e_mail_community.send_keys(str(self.__email))

        # Avançando para o processo de inserir a senha
        buttom_next = driver.find_element(By.CSS_SELECTOR, 'button.slds-button.slds-button_brand.button')
        buttom_next.click()

        # Inserindo senha
        senha_community = driver.find_element(By.XPATH, '//input[@placeholder="Password"]')
        senha_community.send_keys(str(self.__senha))

        # Realizando login
        button_login = driver.find_element(By.CSS_SELECTOR,'button.slds-button.slds-button_brand.button')
        button_login.click()
        wait.until(EC.visibility_of_element_located((By.ID, 'salesOrderInputEmail')))

        driver.switch_to.window(driver.window_handles[1])
        driver.refresh()

        wait.until(EC.visibility_of_element_located((By.ID, 'inputTrackingNo')))

        driver.switch_to.window(driver.window_handles[0]) 

    def sales_login(self, driver: webdriver):
        """ Função responsável por realizar o login na plataforme de sales order de acordo com as credenciais enviadas por e-mail

        Args:
            driver (WebDriver): Instância do Selenium WebDriver utilizada para interagir com a página.
        """
        email_teste_sales = driver.find_element(By.ID, 'salesOrderInputEmail')
        email_teste_sales.send_keys(str(self.__email_teste))

        senha_teste_sales = driver.find_element(By.ID, 'salesOrderInputPassword')
        senha_teste_sales.send_keys(str(self.__senha_teste))

        button_login_sales = driver.find_element(By.CSS_SELECTOR, 'a.btn.btn-primary.btn-user.btn-block')
        button_login_sales.click()
        
        button_sales_order = driver.find_element(By.XPATH,  '//a[@href="salesorder-applist.html"]')
        button_sales_order.click()
    
    def filter_product(self, driver: webdriver) -> list:
        """ Função reponsável pelo retorno das 50 linhas dos pedidos filtrados para que possamos coletar seus respectivos tracking numbers e filtrar os produtos em confirmação.

        Args:
            driver (WebDriver): Instância do Selenium WebDriver utilizada para interagir com a página.

        Returns:
            list: linhas dos 50 pedidos filtrados.
        """
        options_entries = Select(driver.find_element(By.NAME, "salesOrderDataTable_length"))
        options_entries.select_by_value('50') 

        tabela_order_status = driver.find_element(By.ID, 'salesOrderDataTable')
        body_tabela = tabela_order_status.find_element(By.TAG_NAME, "tbody")
        linhas_order_status = body_tabela.find_elements(By.XPATH, "tr")
        return linhas_order_status
    
    def expand_info_product(self, driver: webdriver, wait: WebDriverWait, cod_sales_order: list):
        """ Função responsável por clicar no botão de mais informações dos pedidos a serem verificados.

        Args:
            driver (WebDriver): Instância do Selenium WebDriver utilizada para interagir com a página.
            wait (WebDriverWait): Instância do Selenium WebDriverWait usada para aguardar elementos na página.
            cod_sales_order (list): lista contendo os ids dos pedidos em verificação.
        """
        button_expand = driver.find_element(By.XPATH, '//i[@class="fas fa-square-plus i-{}"]'.format(cod_sales_order))
        driver.execute_script("arguments[0].scrollIntoView(true);", button_expand)
        button_expand.click()
        
        wait.until(EC.visibility_of_element_located((By.XPATH, '//table[@class="sales-order-items t-{}"]'.format(cod_sales_order))))
        time.sleep(5)
        
    def shipment_overview(self, driver: webdriver, cod_sales_order: list) -> list:
        """ Função responsável pelo retorno das linhas da tabela shipment overview com as informações sobre a ultima atualização de envio do produto em verificação.

        Args:
            driver (WebDriver): Instância do Selenium WebDriver utilizada para interagir com a página.
            cod_sales_order (list): lista contendo os ids dos pedidos em verificação.

        Returns:
            list: lista contendo as linhas da tabela shipment overview.
        """
        tabela_tracking = driver.find_element(By.XPATH, '//table[@class="sales-order-items t-{}"]'.format(cod_sales_order))
        body_tabela_tracking = tabela_tracking.find_element(By.TAG_NAME, "tbody")
        linhas_tracking = body_tabela_tracking.find_elements(By.XPATH, "tr")
        return linhas_tracking
    
    def coleta_tracking_number(self, driver: webdriver, wait: WebDriverWait, tracking_number: str):
        """ Função responsável pela coleta do código de identificação de cada produto indexado ao pedido com os status filtrados anteriormente.

        Args:
            driver (WebDriver): Instância do Selenium WebDriver utilizada para interagir com a página.
            wait (WebDriverWait): Instância do Selenium WebDriverWait usada para aguardar elementos na página.
            tracking_number (str): Codigo do produto a ser verificado na página de tracking
        """
        driver.switch_to.window(driver.window_handles[1])
        
        send_trackin_number = driver.find_element(By.ID, 'inputTrackingNo')
        send_trackin_number.clear()
        send_trackin_number.send_keys(tracking_number)
        
        button_track = driver.find_element(By.ID, 'btnCheckStatus')
        button_track.click()
        
        wait.until(EC.visibility_of_element_located((By.ID, 'shippingTable')))
        time.sleep(5)
    
    def close_button(self, driver: webdriver, wait: WebDriverWait, cod_sales_order: list):
        """ Função responsável por clicar no botão Close quando existe pelo menos um produto com status diferente de confirmado.

        Args:
            driver (WebDriver): Instância do Selenium WebDriver utilizada para interagir com a página.
            wait (WebDriverWait): Instância do Selenium WebDriverWait usada para aguardar elementos na página.
            cod_sales_order (list): lista contendo os ids dos pedidos em verificação.
        """
        driver.switch_to.window(driver.window_handles[0])
        
        wait.until(EC.visibility_of_element_located((By.XPATH, '//button[@onclick="cancel(\'{}\')"]'.format(cod_sales_order))))
        time.sleep(5)
        
        button_close = driver.find_element(By.XPATH, '//button[@onclick="cancel(\'{}\')"]'.format(cod_sales_order))
        button_close.click()

    def generate_invoice_button(self, driver: webdriver, wait: WebDriverWait, cod_sales_order: list):
        """ Função responsável por clicar no botão Generate Invoice quando todos os produtos estão com status delivery confirmados

        Args:
            driver (WebDriver): Instância do Selenium WebDriver utilizada para interagir com a página.
            wait (WebDriverWait): Instância do Selenium WebDriverWait usada para aguardar elementos na página.
            cod_sales_order (list): lista contendo os ids dos pedidos em verificação.
        """
        driver.switch_to.window(driver.window_handles[0])
        
        wait.until(EC.visibility_of_element_located((By.XPATH, '//button[@onclick="markInvoiceSent(\'{}\')"]'.format(cod_sales_order))))
        time.sleep(5)
        
        button_close = driver.find_element(By.XPATH, '//button[@onclick="markInvoiceSent(\'{}\')"]'.format(cod_sales_order))
        button_close.click()
        time.sleep(5)   
    
    def coleta_status_delivery(self, driver: webdriver) -> str: 
        """ Função responsável por coletar o status delivery do pedido de acordo com o tracking number inserido.

        Args:
            driver (WebDriver): Instância do Selenium WebDriver utilizada para interagir com a página.

        Returns:
            str: O status de entrega mais recente do pedido.
        """
        tabela_shipment = driver.find_element(By.ID, 'shippingTable')
        body_tabela_shipment = tabela_shipment.find_element(By.TAG_NAME, "tbody")
        linhas_shipment = body_tabela_shipment.find_elements(By.TAG_NAME, "tr")
        
        status_delivery = linhas_shipment[-1].text
        return status_delivery