import pyautogui
import pyperclip
import pandas as pd
import numpy
import openpyx1
from time import sleep

#pega posição
'''sleep(5)
var = pyautogui.position()
print(var)'''

#PAUSA ENTRE OS PYAUTOGUI
pyautogui.pause = 1
#Entrar no sistema da empresa (no caso vai ser o link do drive)
pyautogui.click(x=185, y=557, clicks=2)
sleep(1)
pyperclip.copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
sleep(5) #DEMORA ALGUNS SEGUNGOS

#NAVEGAR NO SISTEMA E ENCONTRA A BASE DE DADOS (ENTRA NA PORTA EXPORTAR)
pyautogui.click(x=349, y=290, clicks=2)
sleep(2)

#XPORTAR/FAZER DOWNLOAD DA BASE DE DADOS
pyautogui.click(x=355, y=290) #CLICAR NO ARQUIVO
sleep(1)
pyautogui.click(x=1156, y=189) #CLICAR NOS 3 PONTINHOS
sleep(1)
pyautogui.click(x=924, y=592) #CLICAR NO FAZER DOWNLOAD
sleep(5)

#IMPORTAR A BASE DE DADOS PARA O PYTHON
tabela = pd.read_excel(r'C:\Users\FLEXLAN\Downloads\Vendas - Dez.xlsx')
print(tabela)

#CALCULAR OS INDICADORES
#faturamento é soma da coluna valor final.
faturamento = tabela['Valor Final'].sum()
print(' {:,.2f}'.format(faturamento))

#quantidade de produtos é a soma da coluna quantidade.
quantidade = tabela['Quantidade'].sum()
print('R${}'.format(quantidade))

#ENVAIR UM EMAIL PARA DIRETORIA COM O RELATORIO
#Abrir o email ( link:https://mail.google.com/mail/u/0/#inbox)
pyautogui.hotkey('ctrl', 't')
pyperclip.copy('https://mail.google.com/mail/u/0/#inbox')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
sleep(3)
#clicar no escrever
pyautogui.click(x=79, y=199)
sleep(1)
#digita o destinatario
pyautogui.write('joaoljl134@hotmail.com')
pyautogui.press('tab') # SELECIONANDO O DESTINATARIO
pyautogui.press('tab')# PASSA PARA O CAMPO DE ASSUNTO

#ESCREVA O ASSUNTO
pyperclip.copy('Relatório de vendas')
pyautogui.hotkey('ctrl','v')
pyautogui.press('tab')

#ESCREVA O CORPO DO EMAIL
texto = '''
Prezados, bom dia

O faturamento de ontem foi de {:,.2f}
A quantidade de produtos foi de {:,}  

abs
João Lucas 
'''.format(faturamento,quantidade)
pyperclip.copy(texto)
pyautogui.hotkey('ctrl', 'v')

#ENVIAR O EMAIL
pyautogui.click(x=842, y=640)
