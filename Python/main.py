import pyperclip
import pyautogui
import time

pyautogui.PAUSE = 1

pyautogui.press('win')
pyautogui.write("chrome")
pyautogui.press('Enter')

link = "http://web.whatsapp.com"

pyperclip.copy(link)
pyautogui.hotkey('ctrl','v')
pyautogui.press('Enter')

