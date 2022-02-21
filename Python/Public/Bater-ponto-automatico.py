#
# Script criado para realizar a batida de ponto no sistema Ahgora(SONDA) de forma automatizada
# Utilizando webbrowser e selenium
# Criado em 18/02/2022 por Lucas Moreira
#

# Importando bibliotecas

from webbrowser import Chrome
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time

# Definindo variaveis iniciais
site = "https://www.ahgora.com.br/novabatidaonline/?defaultDevice=a774091"
email = "e-mail"
password = "password"

# Iniciando o navegador (Chrome)
driver = webdriver.Chrome()
driver.get(site)
time.sleep(1)

# Comando para clicar no botão da primeira tela
driver.find_element_by_xpath('//*[@id="root"]/div/div[2]/div[1]/div[2]/button').click()
time.sleep(1)

# Comando para clicar no botão da segunda tela
driver.find_element_by_xpath('//*[@id="root"]/div/div[1]/div[3]/div[2]/button/span[1]/p').click()
time.sleep(1)

# Declarando elementos do site como variaveis
txtbox_email = driver.find_element_by_name('UserName')
txtbox_password = driver.find_element_by_name('Password')

# Preenchendo textbox com valores de usuário e senha e realizando Login
txtbox_email.send_keys(email)
txtbox_password.send_keys(password)
txtbox_password.send_keys(Keys.ENTER)
time.sleep(1)

# Comando para clicar no botão e realizar a batida de ponto.
driver.find_element_by_xpath('//*[@id="root"]/div/div[1]/div[3]/div[2]/button[2]').click()
time.sleep(1)

# Fechando navegador
driver.quit()


# Log de atualizações (upgrades)
# 
# 21/02/2022 - Removido dados pessoais de login e convertido variáveis de PT/BR para EN/US
