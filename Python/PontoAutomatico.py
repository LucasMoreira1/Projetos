#!/usr/bin/python
#
# Script criado para realizar a batida de ponto no sistema Ahgora(SONDA) de forma automatizada
# Utilizando webbrowser e selenium
# Criado em 18/02/2022 por Lucas Moreira
#

# Importando bibliotecas

from webbrowser import Chrome
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC 
from selenium.webdriver.support.ui import WebDriverWait
import time

# Definindo variaveis iniciais
site = "https://www.ahgora.com.br/novabatidaonline/?defaultDevice=a774091"
email = "lucas.ferreiram@sonda.com"
senha = "!Parrudo292702"

# Iniciando o navegador (Chrome)
navegador = webdriver.Chrome()
navegador.get(site)
time.sleep(2)

# Comando para clicar no botão da primeira tela
navegador.find_element(By.XPATH, '//*[@id="root"]/div/div[2]/div[1]/div[2]/button').click()
time.sleep(1)

# Comando para clicar no botão da segunda tela
navegador.find_element(By.XPATH, '//*[@id="root"]/div/div[1]/div[3]/div[2]/button/span[1]/p').click()
time.sleep(1)

# Declarando elementos do site como variaveis
campo_email = navegador.find_element(By.NAME, 'UserName')
campo_senha = navegador.find_element(By.NAME, 'Password')

# Preenchendo textbox com valores de usuário e senha e realizando Login
campo_email.send_keys(email)
campo_senha.send_keys(senha)
campo_senha.send_keys(Keys.ENTER)
time.sleep(2)

# Comando para clicar no botão e realizar a batida de ponto.
navegador.find_element(By.XPATH,'//*[@id="root"]/div/div[1]/div[3]/div[2]/button[2]').click()
time.sleep(2)

# Fechando navegador
navegador.quit()