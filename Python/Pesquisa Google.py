from webbrowser import Chrome
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
import numpy as np
import csv

arquivo=open("pesquisa_google.txt",'a')

driver=webdriver.Chrome()
pesquisa="condominio"
time.sleep(2)
driver.get("http://www.google.com")
time.sleep(2)

search = driver.find_element(By.NAME, 'q')
time.sleep(1)
search.send_keys(pesquisa)
time.sleep(1)
search.send_keys(Keys.RETURN)

time.sleep(1)
link=driver.find_elements(By.CLASS_NAME,'cXedhc')
link.click()

try:
    main = WebDriverWait(driver,10).until(
        EC.presence_of_element_located((By.ID, "main"))
    )
    nomes = driver.find_elements(By.CLASS_NAME,'cXedhc')
    
    for nome in nomes:
        header = nome.find_elements(By.CLASS_NAME, 'rllt__details')
        texto = header.text
        print(texto)
        with open("pesquisa_google.txt","a") as arquivo:
            arquivo.write("\n")
            arquivo.write(texto)

        # data = pd.DataFrame({'Dados': 'dados'})
        # data.append({'Dados': [header.text]},ignore_index=True)
        # datatoexcel = pd.ExcelWriter("doPython.xlsx",engine='xlsxwriter')
        # data.to_excel(datatoexcel, sheet_name='Sheet1')

        # datatoexcel.save()
finally:
        arquivo.close()
        driver.quit()