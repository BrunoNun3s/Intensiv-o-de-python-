import pyautogui
import pyperclip
import time
import datetime

pyautogui.PAUSE = 4

pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("enter")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
pyautogui.click(x=457, y=323, clicks=3)
pyautogui.click(x=457, y=323)
pyautogui.click(x=918, y=199)
pyautogui.click(x=662, y=688)
time.sleep(10)

import pandas as pd
import numpy as np
import openpyxl

tabela = pd.read_excel(r"C:\Users\João Victor\Downloads\Vendas - Dez.xlsx")
#print(tabela)
quantidade = tabela["Quantidade"].sum()
faturamento = tabela["Valor Final"].sum()
#https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox

pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com/mail/u/1/?ogbl#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(7)
pyautogui.click(x=98, y=209)
pyperclip.copy("fabi.nunes.franca@gmail.com")
pyautogui.hotkey("ctrl", "v")
pyautogui.click(x=420, y=579)
pyperclip.copy("Relatório de vendas")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")

texto = f"""
Bom dia!

O faturamento de ontem foi de: R${faturamento:,.2f}
E a quantidade de produtos vendidos foi de:{quantidade:,}
"""

pyautogui.write(texto)
time.sleep(2)
pyautogui.hotkey("ctrl", "enter")
