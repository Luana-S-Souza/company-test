import os
import time
from dotenv import load_dotenv
from Bot import Bot
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

# Instanciando a classe
bot = Bot()

# Inicializando uma instância Google Chrome
driver = bot.set_driver()
# Definindo as esperas do navegador
wait = bot.set_wait(driver)

# Iniciando automação
try:
    # Abrindo as duas janelas principais
    bot.open_window(driver)
    
    # Aceitando os cookies necessários
    bot.close_cookie(driver)

    # Realizando community login inicial
    bot.community_login(driver, wait)
    
    # Realizando sales order login
    bot.sales_login(driver)
    
    # Filtrando os 50 produtos em sales order
    linhas_order_status = bot.filter_product(driver)

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

