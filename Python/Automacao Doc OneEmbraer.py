from webbrowser import Chrome
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

navegador = webdriver.Chrome()
navegador.get('http://embshare-com.embraer.com.br/procedimento_ti/Lists/Procedimento%20No%20SAP/AllItems.aspx')
