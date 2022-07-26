import pyautogui
import pyperclip
import time
import pandas as pd
import numpy
import openpyxl
#pyautogui.click #clicar com ouse
#pyautogui.write #escrever um texto
#pandas importa base de dados pro python



pyautogui.press("win")
pyautogui.write("Opera GX")
pyautogui.press("enter")
time.sleep(1)
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(4)  #tempo de delay
pyautogui.click(x=425, y=326, clicks=2)
time.sleep(2)
pyautogui.click(x=425, y=326, button="right")
time.sleep(2)
pyautogui.click(x=627, y=929)


tabela = pd.read_excel(r"C:\Users\curve\Downloads\Vendas - Dez.xlsx")
print(tabela)

quantidade = tabela["Quantidade"].sum() # soma da coluna quantidade
faturamento = tabela["Valor Final"].sum() # soma da coluna valor final
print(quantidade)
print(faturamento)

pyautogui.press("win")
pyautogui.write("Opera GX")
pyautogui.press("enter")
time.sleep(2)
pyautogui.click(x=329, y=50)
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox") #seu email precisa estar logado
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(4)
pyautogui.click(x=98, y=227)
pyperclip.copy("destinatario@gmail.com") #insira aqui algum destinatário
time.sleep(1)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")
assunto = "Relatório de vendas"
pyperclip.copy(assunto)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")

texto = f"""Prezados, bom dia

O faturamento de ontem foi de : R$ {faturamento:,.2f}
A Quantidade de produtos foi de: {quantidade:,}


Daniel Python
"""

pyautogui.write(texto)







# \n
# r diz pro python pra le o texto e nao le as contrabarras


#clicando com o mouse com o py
# pyautogui.sleep(4)
# print(pyautogui.position()) 
