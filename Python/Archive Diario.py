#Passo 1 - Abrir e logar no SAP
import pyautogui
import time

usuario = 'schedule2'
senha = 'sCHeduLe2##'

pyautogui.PAUSE = 2
pyautogui.alert('Precione "OK" para iniciar o processamento')
#Abrindo SAP
pyautogui.hotkey('win','r')
pyautogui.write('sapgui sapaebp1 41')
pyautogui.press('Enter')
time.sleep(30) #Pausa para processamento

#Logando
pyautogui.write(usuario)
pyautogui.press('tab')
pyautogui.write(senha)
pyautogui.press('Enter')
time.sleep(5) #Pausa para processamento
pyautogui.press('Enter')
pyautogui.press('tab')
pyautogui.press('space')
time.sleep(5) #Pausa para processamento

#Acessando SARA, realizando seleção e deleção dos arquivos
time.sleep(5)

pyautogui.write('SARA')
pyautogui.press('Enter')
pyautogui.write('IDOC')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('space')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('space')
pyautogui.press('F8')
pyautogui.press('Enter')

#Data inicio
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('space')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('space')
pyautogui.hotkey('Ctrl','s')

#Parametros spool
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('space')
pyautogui.write('LOCAL')
pyautogui.hotkey('Shift','F1')
pyautogui.press('tab')
pyautogui.press('space')
time.sleep(3)
pyautogui.press('F8')

#Voltando a pagina inicial
pyautogui.press('F12')
pyautogui.press('F12')

#Aguardando execução dos Jobs

time.sleep(300) #10 minutos de espera para execução
#usar função while para captar a informação de quando o Job concluir
#pyautogui.alert('Verifique se o Job Finalizou com sucesso')

#Passo 3 - Após conclusão liberar Job no CTM
pyautogui.click(421,756) 
pyautogui.hotkey('Ctrl','w')
pyautogui.hotkey('Ctrl','w')
pyautogui.hotkey('Ctrl','w')
pyautogui.hotkey('Ctrl','w')
pyautogui.hotkey('Ctrl','w')
pyautogui.hotkey('Ctrl','w')
pyautogui.hotkey('Ctrl','w')
time.sleep(5)
pyautogui.click(x=339, y=249)
pyautogui.write('all jobs')
pyautogui.click(x=165, y=343)
time.sleep(5)
pyautogui.hotkey('Ctrl','f')
pyautogui.write('ZARC_ZRSRLDREL0_S01')
pyautogui.press('Enter')
pyautogui.click(190,619)
pyautogui.click(751,67)
pyautogui.press('Enter')


pyautogui.alert("Execução completa")

