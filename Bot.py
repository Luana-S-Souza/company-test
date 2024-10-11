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
    """ Classe responsável pelas principais funções do robô.
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
        print(self.__url_sales)
    
    def close_cookie(self, driver):
        try:
            # Localizando o botão de aceitação de cookies e clicando nele, caso exista
            button_cookie = driver.find_element(By.ID, 'onetrust-accept-btn-handler')
            button_cookie.click()
            print("Botão de aceitação de cookies foi clicado.")
        except NoSuchElementException:
            # Caso o botão ja tenha sido clicado, ou não exista, seguiremos com a automação.
            print("Botão de aceitação de cookies não encontrado.")
        except ElementClickInterceptedException:
            print("A sobreposição de cookies estava bloqueando o clique.")
    
    def community_login(self, driver, wait):
    # Essa etapa precisa ser realizada em casos de primeiro login, é necessario criar um usuario e senha ou, caso já tenha, logar antes de iniciarmos a automação
        # Localizando o botão de community login
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

    def sales_login(self, driver):
        # Iniciando login no Sales order com as credenciais enviadas na apresentação do teste via email
        email_teste_sales = driver.find_element(By.ID, 'salesOrderInputEmail')
        email_teste_sales.send_keys(str(self.__email_teste))

        senha_teste_sales = driver.find_element(By.ID, 'salesOrderInputPassword')
        senha_teste_sales.send_keys(str(self.__senha_teste))

        button_login_sales = driver.find_element(By.CSS_SELECTOR, 'a.btn.btn-primary.btn-user.btn-block')
        button_login_sales.click()
        
        button_sales_order = driver.find_element(By.XPATH,  '//a[@href="salesorder-applist.html"]')
        button_sales_order.click()
    
    def filter_product(self, driver):
        options_entries = Select(driver.find_element(By.NAME, "salesOrderDataTable_length"))
        options_entries.select_by_value('50') 

        tabela_order_status = driver.find_element(By.ID, 'salesOrderDataTable')
        body_tabela = tabela_order_status.find_element(By.TAG_NAME, "tbody")
        linhas_order_status = body_tabela.find_elements(By.XPATH, "tr")
        return linhas_order_status