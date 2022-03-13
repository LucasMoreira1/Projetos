
from re import search
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
import time
import csv

## Fazer a pesquisa Google
# url = "https://www.google.com/search?q="
# user_input = input("Digite o que deseja pesquisar: ")
# print("Pesquisando...")

## Criar arquivo código-fonte HTML com o resultado

# google_search = requests.get(url+user_input)

#google_search = requests.get("https://www.google.com/search?sa=X&tbs=lf:1,lf_ui:2&tbm=lcl&sxsrf=APq-WBvx0odW9hxVXYd0KBHjEhCtaAjCzg:1646224554418&q=condominio&rflfq=1&num=10&ved=2ahUKEwjplKWuuKf2AhWWGbkGHXs6AtAQjGp6BAgQEAE&biw=1366&bih=657&dpr=1#cobssid=s&rlfi=hd:;si:;mv:[[-23.2164085,-45.904273499999995],[-23.2248216,-45.9131438]];tbs:lrf:!1m4!1u3!2m2!3m1!1e1!1m4!1u2!2m2!2m1!1e1!2m1!1e2!2m1!1e3!3sIAE,lf:1,lf_ui:2")

# codificar resultado para aparecer caracteres especiais

with open('pesquisa.html') as html_file:
    soup = BeautifulSoup(html_file, 'lxml')
    
    arquivo = open("arquivo.csv","a")

    arquivo.write("Resultados: ")
    
    arquivo.write("\n\n")
    arquivo.close() 

for div in soup.find_all('div', class_='cXedhc'):
    nome = div.span.text
    print(nome)
    
    print(div.text)

    print()
    
    arquivo = open("arquivo.csv","a")

    arquivo.write(div.text)
    arquivo.write("\n")

    arquivo.close() 