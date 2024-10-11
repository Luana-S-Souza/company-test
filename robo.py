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

# Acessando as variáveis de ambiente
email = os.getenv('EMAIL')
senha = os.getenv('SENHA')
email_teste = os.getenv('EMAIL_TESTE')
senha_teste = os.getenv('SENHA_TESTE')

# Inicializando uma instância do Google Chrome
service = Service()

# Definindo a preferência para o brower do Chrome
options = webdriver.ChromeOptions()

# Iniciando a instância do Chrome com o service e options já definidos
driver = webdriver.Chrome(service=service, options=options)
driver.implicitly_wait(10)
wait = WebDriverWait(driver, timeout=60)

# Iniciando automação
try:
    # Abrindo a pagina de login
    driver.get('https://pathfinder.automationanywhere.com/challenges/salesorder-applogin.html#')

    driver.execute_script("window.open('');") 
    driver.switch_to.window(driver.window_handles[1]) 
    driver.get('https://pathfinder.automationanywhere.com/challenges/salesorder-tracking.html')

    driver.switch_to.window(driver.window_handles[0]) 

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

    # Essa etapa precisa ser realizada em casos de primeiro login, é necessario criar um usuario e senha ou, caso já tenha, logar antes de iniciarmos a automação
    # Localizando o botão de community login
    button_community = driver.find_element(By.ID, 'button_modal-login-btn__iPh6x')
    button_community.click()

    # Inserindo e-mail para login na plataforma
    e_mail_community = driver.find_element(By.ID, '43:2;a')
    e_mail_community.send_keys(str(email))

    # Avançando para o processo de inserir a senha
    buttom_next = driver.find_element(By.CSS_SELECTOR, 'button.slds-button.slds-button_brand.button')
    buttom_next.click()

    # Inserindo senha
    senha_community = driver.find_element(By.XPATH, '//input[@placeholder="Password"]')
    senha_community.send_keys(str(senha))

    # Realizando login
    button_login = driver.find_element(By.CSS_SELECTOR,'button.slds-button.slds-button_brand.button')
    button_login.click()

    wait.until(EC.visibility_of_element_located((By.ID, 'salesOrderInputEmail')))

    driver.switch_to.window(driver.window_handles[1])
    driver.refresh()

    wait.until(EC.visibility_of_element_located((By.ID, 'inputTrackingNo')))

    driver.switch_to.window(driver.window_handles[0]) 

    # Iniciando login no Sales FOrce com as credenciais enviadas na apresentação do teste via email
    email_teste_sales = driver.find_element(By.ID, 'salesOrderInputEmail')
    email_teste_sales.send_keys(str(email_teste))

    senha_teste_sales = driver.find_element(By.ID, 'salesOrderInputPassword')
    senha_teste_sales.send_keys(str(senha_teste))

    button_login_sales = driver.find_element(By.CSS_SELECTOR, 'a.btn.btn-primary.btn-user.btn-block')
    button_login_sales.click()

    button_sales_order = driver.find_element(By.XPATH,  '//a[@href="salesorder-applist.html"]')
    button_sales_order.click()

    options_entries = Select(driver.find_element(By.NAME, "salesOrderDataTable_length"))
    options_entries.select_by_value('50') 

    tabela_order_status = driver.find_element(By.ID, 'salesOrderDataTable')
    body_tabela = tabela_order_status.find_element(By.TAG_NAME, "tbody")
    linhas_order_status = body_tabela.find_elements(By.XPATH, "tr")

    status_filter = ["Delivery Outstanding", "Confirmed"]
    cod_sales_order = []

    for linha in linhas_order_status:
        order_status = linha.find_elements(By.XPATH, 'td')
        if order_status[4].text in status_filter:
            cod_sales_order.append(order_status[1].text)

    for i in range(len(cod_sales_order)):
        print(cod_sales_order[i])

        button_expand = driver.find_element(By.XPATH, '//i[@class="fas fa-square-plus i-{}"]'.format(cod_sales_order[i]))
        driver.execute_script("arguments[0].scrollIntoView(true);", button_expand)
        button_expand.click()
        
        wait.until(EC.visibility_of_element_located((By.XPATH, '//table[@class="sales-order-items t-{}"]'.format(cod_sales_order[i]))))
        time.sleep(5)
        
        tabela_tracking = driver.find_element(By.XPATH, '//table[@class="sales-order-items t-{}"]'.format(cod_sales_order[i]))
        body_tabela_tracking = tabela_tracking.find_element(By.TAG_NAME, "tbody")
        linhas_tracking = body_tabela_tracking.find_elements(By.XPATH, "tr")
        
        count = 0
        for linha in linhas_tracking:         
            tracking_number = linha.find_elements(By.XPATH, 'td')[1].text
            driver.switch_to.window(driver.window_handles[1])
            
            send_trackin_number = driver.find_element(By.ID, 'inputTrackingNo')
            send_trackin_number.clear()
            send_trackin_number.send_keys(tracking_number)
            
            button_track = driver.find_element(By.ID, 'btnCheckStatus')
            button_track.click()
            
            wait.until(EC.visibility_of_element_located((By.ID, 'shippingTable')))
            time.sleep(5)
            
            tabela_shipment = driver.find_element(By.ID, 'shippingTable')
            body_tabela_shipment = tabela_shipment.find_element(By.TAG_NAME, "tbody")
            linhas_shipment = body_tabela_shipment.find_elements(By.TAG_NAME, "tr")
            
            status_delivery = linhas_shipment[-1].text
            
            if "Delivered" not in status_delivery.split():
                driver.switch_to.window(driver.window_handles[0])
                
                wait.until(EC.visibility_of_element_located((By.XPATH, '//button[@onclick="cancel(\'{}\')"]'.format(cod_sales_order[i]))))
                time.sleep(5)
                
                button_close = driver.find_element(By.XPATH, '//button[@onclick="cancel(\'{}\')"]'.format(cod_sales_order[i]))
                button_close.click()                
                break
            else:
                count+=1
                if count == len(linhas_tracking):
                    driver.switch_to.window(driver.window_handles[0])
                    
                    wait.until(EC.visibility_of_element_located((By.XPATH, '//button[@onclick="markInvoiceSent(\'{}\')"]'.format(cod_sales_order[i]))))
                    time.sleep(5)
                    
                    button_close = driver.find_element(By.XPATH, '//button[@onclick="markInvoiceSent(\'{}\')"]'.format(cod_sales_order[i]))
                    button_close.click()
                    time.sleep(5)                    
                    
            driver.switch_to.window(driver.window_handles[0])

except Exception as e:
    print(f'Ocorreu um erro: {e}')
finally:
    driver.quit()

