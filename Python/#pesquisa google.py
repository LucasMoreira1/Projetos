from re import search
import requests, webbrowser
from bs4 import BeautifulSoup

user_input = input("digite oq deseja pesquisar: ")
print("pesquisando...")

google_search = requests.get("https://www.google.com/search?q="+user_input)

soup=BeautifulSoup(google_search.text, 'html.parser')

search_results = soup.select('Telefone')
print(search_results)