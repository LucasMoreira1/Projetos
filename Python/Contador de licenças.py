import pyperclip
import pyautogui
import pandas
import time

pyautogui.PAUSE = 2

pyautogui.click(421,756) 
pyautogui.hotkey('Ctrl','w')
pyautogui.hotkey('Ctrl','w')
pyautogui.hotkey('Ctrl','w')
time.sleep(5)
pyautogui.click(350,242)
pyautogui.write('All Jobs')
pyautogui.click(185,342)
time.sleep(5)
pyautogui.hotkey('Ctrl','l')
time.sleep(5)
pyautogui.click(487,312)
pyautogui.click(487,312)
pyautogui.hotkey('Ctrl','a')
pyautogui.hotkey('Ctrl','c')

df = pandas.read_clipboard()
colunas = len(df)
pyperclip.copy(colunas)
licencas = colunas

pyautogui.alert(licencas,' licenças')