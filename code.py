# Import libraries
from selenium.webdriver.support.ui import Select
from selenium import webdriver
import pandas as pd
import holidays
import time
import os

# Initializing Chrome Webdriver
options = webdriver.ChromeOptions()
options.binary_location = "C:/Program Files/Google/Chrome/Application/chrome.exe" # Google Chrome path in local directory. <CHANGE>
chrome_driver_binary = "C:/Program Files/Google/Chrome/Application/chromedriver.exe" # WebDriver path in Local directory. <CHANGE>

# ----------- Code.
sdate = str(input('Digite a data de início (m/d/yyyy): ')) # Start date defined by user input.
edate= str(input('Digite a data final (m/dd/yyyy): ')) # End date defined by user input.
feriados = holidays.Brazil() # List containing national holidays.

# Date format:
days = pd.bdate_range(sdate,edate,freq='b', holidays=feriados) # Range of business days between start date and end date.
days_format1 = days.strftime('%d/%m/%Y') # Date format as 'dd/mm/yyyy'
days_format2 = days.strftime('%Y%m%d') # Date format as 'yyyymmdd'

# Opening login url:
browser = webdriver.Chrome(chrome_driver_binary, options=options)
browser.get("https://portal.provedoradeprecos.com.br/efetua_login/")
time.sleep(2)

# Sending login data:
username = browser.find_element_by_id("id_text_login")
password = browser.find_element_by_id("id_text_senha")
usuario = str(input('Digite seu nome de usuário: '))
senha = str(input('Digite sua senha: '))
username.send_keys(usuario)
password.send_keys(senha)

login_attempt = browser.find_element_by_xpath("//*[@type='submit']")
login_attempt.submit()
time.sleep(5)

# Once the access is allowed, the following For Loop starts to download files:
for i,j in zip(days_format1, days_format2):
    print("Iniciando "+i)
    tipo_produto = browser.find_element_by_id("id_tipo_produto")
    select = Select(tipo_produto)

    try:
        select.select_by_visible_text('Títulos Privados')
    except:
        print('The item does not exists')

    selecionar_dia = browser.find_element_by_id('id_data')
    select = Select(selecionar_dia)

    try:
        select.select_by_visible_text(i)
    except:
        print('The date does not exists')

    produto = browser.find_element_by_id("id_produto")
    select = Select(produto)

    try:
        select.select_by_visible_text('---')
    except:
        print('The item does not exists')

    filter_attempt = browser.find_element_by_xpath("//*[@type='submit']")
    filter_attempt.submit()
    time.sleep(10)
    importar = browser.find_element_by_link_text("baixar dados data (todos)").click()
    time.sleep(5)

    # Rename default file to 'POP_Prices_' + reference date.
    path = r'C:/Users/iamju/Downloads/' # Path to 'Download' folder. <CHANGE>
    nome_arq = 'pop_meus_precos'

    old_name = path + nome_arq + '.xlsx'
    new_name = path + 'POP_Prices_' + j + '.xlsx'
    os.rename(old_name, new_name)
    print("Finalizado: " + i)
