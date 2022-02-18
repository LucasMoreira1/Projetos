from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import urllib
import pyautogui
import pyperclip
import pandas as pd


contatos_df = pd.read_excel("leads.xlsx",engine='openpyxl')
contatos_df

navegador = webdriver.Chrome()
navegador.get('http://web.whatsapp.com')

while len(navegador.find_element_by_id('side')) < 1:
    time.sleep(1)

for i, mensagem in enumerate(contatos_df['Mensagem']):
    pessoa = contatos_df.loc[i, 'Pessoa']
    numero = contatos_df.loc[i, 'Número']
    texto = urllib.parse.quote(f"Oi {pessoa}! {mensagem}")
    link = f"http://web.whatsapp.com/send?phone={numero}&text={texto}"
    navegador.get(link)
    while len(navegador.find_element_by_id('side')) < 1:
        time.sleep(1)
    navegador.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]').send_keys(Keys.ENTER)
    time.sleep(10)

## Versao de testes: 
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import urllib
import pyautogui
import pyperclip
import pandas as pd

contatos_df = pd.read_excel("automacao.xlsx",engine='openpyxl')
contatos_df

navegador = webdriver.Chrome()
navegador.get('http://web.whatsapp.com')

while len(navegador.find_element_by_xpath('//*[@id="pane-side"]')) < 1:
    time.sleep(1)

for i, mensagem in enumerate('Tudo bem?'):
    nome = contatos_df.loc[i, 'Nome do Cliente']
    numero = contatos_df.loc[i, 'Telefone']
    texto = urllib.parse.quote(f"Oi {nome}! {mensagem}")
    link = f"http://web.whatsapp.com/send?phone=55{numero}&text={texto}"
    navegador.get(link)
    while len(navegador.find_element_by_id('pane-side')) < 1:
        time.sleep(1)
    #navegador.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]').send_keys(Keys.ENTER)
    #time.sleep(10)

