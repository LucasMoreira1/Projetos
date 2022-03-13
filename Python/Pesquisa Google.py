from webbrowser import Chrome
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from re import search
import requests
from bs4 import BeautifulSoup
import time
import csv
import codecs
import tkinter as tk
from tkinter import ttk

win = tk.Tk() # Nome da aplicação
win.title("Digite o tema da pesquisa: ") # Label
lbl = ttk.Label(win, text = "Pesquisar: ").grid(column = 0, row = 0) # Click event
def click():
    print("Estamos pesquisando por: " + assunto.get())
assunto = tk.StringVar()
nameEntered = ttk.Entry(win, width = 12, textvariable = assunto).grid(column = 0, row = 1)
button = ttk.Button(win, text = "Pesquisar", command = click).grid(column = 1, row = 1)
win.mainloop()

    

driver=webdriver.Chrome()
#pesquisa=input("Digite o que deseja pesquisar: ")
pesquisa=assunto.get()
time.sleep(2)
driver.get("https://www.google.com/search?q="+pesquisa)
time.sleep(2)

botao = driver.find_element(By.XPATH, '//*[@id="Odp5De"]/div/div/div/div[2]/div[1]/div[2]/g-more-link/a/div')
botao.click()
time.sleep(2)

codigofonte=driver.page_source

paginahtml = codecs.open("resultado.html", mode="w",encoding="utf-8")
paginahtml.write(codigofonte)
paginahtml.close()

with open('resultado.html') as html_file:
    soup = BeautifulSoup(html_file, 'lxml')
    

    arquivo = open("resultado.csv","a")

    arquivo.write("Resultados: ")
        
    arquivo.write("\n\n")
    arquivo.close() 
    time.sleep(2)

for div in soup.find_all('div', class_='cXedhc'):
    nome = div.span.text
    print(nome)
    
    print(div.text)

    print()
    
    arquivo = open("resultado.csv","a")

    arquivo.write(div.text)
    arquivo.write("\n")

    arquivo.close() 