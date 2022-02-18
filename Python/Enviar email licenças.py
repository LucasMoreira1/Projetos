import pyperclip
import pyautogui
import time
import datetime

pyautogui.PAUSE = 1

pyautogui.click(171,747)
pyautogui.click(28,103)
time.sleep(5)
destino = pyperclip.copy('<claudionor.ouvera@embraer.net.br>; <awsantana@stefanini.com>; <asjesus3@stefanini.com>; <scheduller@stefanini.com>; <rsalmeida3@stefanini.com>; <wlcruz@stefanini.com>; <rbsantos3@stefanini.com>; <SONDA-EmbraerScheduler@sondachile.onmicrosoft.com>; <mvferreira@stefanini.com>; <rfferreira@stefanini.com>; <luiz.panace@embraer.com.br>')
pyautogui.hotkey('ctrl','v')
pyautogui.press('tab')
pyautogui.press('tab')

from datetime import datetime
hoje = datetime.today().strftime('%d-%m')
assunto = pyperclip.copy('Comunicado de consumo de licenças do Control-M - ' + hoje)
pyautogui.hotkey('ctrl','v')
pyautogui.press('tab') 
msg = pyperclip.copy('Bom dia! ')
pyautogui.hotkey('ctrl','v')
pyautogui.press('enter')
pyautogui.press('enter')
msg = pyperclip.copy('Segue a quantidade de licenças ordenadas no Control-M da Embraer Comercial, Executiva, Melbourne e Jacksonville na data de hoje.')
pyautogui.hotkey('ctrl','v')
pyautogui.press('enter')
pyautogui.press('enter')

pyautogui.click(233,744)
time.sleep(3)
pyautogui.moveTo(327,278)
pyautogui.dragTo(327,649,duration=1)
pyautogui.hotkey('ctrl','c')
pyautogui.hotkey('win','r')
pyautogui.write('mspaint')
pyautogui.press('enter')
time.sleep(5)
pyautogui.hotkey('ctrl','v')
pyautogui.hotkey('ctrl','c')

pyautogui.click(171,747)
pyautogui.click(271,662)
pyautogui.hotkey('ctrl','v')
time.sleep(5)
pyautogui.click(283,522)
pyautogui.click(580,50)
pyautogui.click(762,75)

pyautogui.alert("Fim")