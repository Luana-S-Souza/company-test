
import logging
from Bot import Bot
from selenium.webdriver.common.by import By

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s', 
    datefmt='%Y-%m-%d %H:%M:%S', 
    handlers=[
        logging.FileHandler("app.log", encoding="utf-8"),
        logging.StreamHandler() 
    ]
)

logging.info("Iniciando as iinstâncias do Google Chrome")
# Instanciando a classe
bot = Bot()
# Inicializando uma instância Google Chrome
driver = bot.set_driver()
# Definindo as esperas do navegador
wait = bot.set_wait(driver)

logging.info("Iniciando o funcionamento do robô.")
try:
    try:
        # Abrindo as duas janelas principais
        bot.open_window(driver)
        
        # Aceitando os cookies necessários
        bot.close_cookie(driver)

        # Realizando community login inicial
        bot.community_login(driver, wait)
        
        # Realizando sales order login
        bot.sales_login(driver)
        
        logging.info("Janelas abertas e logins efetuados com sucesso")
        
    except Exception as e:
        logging.error(f'Ocorreu um erro: {e}')    
    
    try: 
        logging.info("Filtrando apenas 50 produtos e buscando os status Delivery Outstanding e Confirmed")
        # Filtrando os 50 produtos em sales order
        linhas_order_status = bot.filter_product(driver)
        
        # Filtrando apenas os pedidos com os status: "Delivery Outstanding", "Confirmed"
        status_filter = ["Delivery Outstanding", "Confirmed"]
        cod_sales_order = []

        # Iterando as linhas da tabela e guardando em uma lista somente o código do produto que esteja nos status filtrados anteriormente
        for linha in linhas_order_status:
            order_status = linha.find_elements(By.XPATH, 'td')
            if order_status[4].text in status_filter:
                cod_sales_order.append(order_status[1].text)
        
        logging.info("Pedidos filtrados com sucesso.")
    except Exception as e:
        logging.error(f'Ocorreu um erro: {e}')
    
    try:
        logging.info("Iniciando coleta de informações de acordo com o tracking number de cada pedido")
        # Para cada código de pedido, vamos expandir suas informações e coletar seu respectivo tracking number
        for i in range(len(cod_sales_order)):
            bot.expand_info_product(driver, wait, cod_sales_order[i])
            linhas_tracking = bot.shipment_overview(driver, cod_sales_order[i])
            
            count = 0
            for linha in linhas_tracking:
                tracking_number = linha.find_elements(By.XPATH, 'td')[1].text
                bot.coleta_tracking_number(driver, wait, tracking_number) 
                status_delivery = bot.coleta_status_delivery(driver)       
                
                logging.info("Tracking Number: {}  Status Delivery: {}\n".format(tracking_number,status_delivery[18:]))
                
                if "Delivered" not in status_delivery.split():
                    bot.close_button(driver, wait, cod_sales_order[i])
                    logging.info("Botão Close clicado para: {}".format(tracking_number))
                    break
                else:
                    count+=1
                    if count == len(linhas_tracking):
                        bot.generate_invoice_button(driver, wait, cod_sales_order[i]) 
                        logging.info("Botão Generate Invoice clicado para: {}".format(tracking_number))              
                driver.switch_to.window(driver.window_handles[0])
        logging.info("Verificação realizada com sucesso.")
    except Exception as e:
        logging.error(f'Ocorreu um erro: {e}')
except Exception as e:
    print(f'Ocorreu um erro: {e}')
finally:
    logging.info("Autmoção finalizada, fechando todas as janelas.")
    driver.quit()

