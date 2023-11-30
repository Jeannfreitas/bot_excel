import openpyxl
import pyautogui
workbok = openpyxl.load_workbook('vendas_de_produtos.xlsx')
vendas = workbok ['vendas']

for linha in vendas.iter_rows(min_row=2):
    #nome
    pyautogui.click(671,359,duration=1.5)
    pyautogui.write(linha[0].value)
    #produto
    pyautogui.click(667,384,duration=1.5)
    pyautogui.write(linha[1].value)
    #quantidade
    pyautogui.click(661,412,duration=1.5)
    pyautogui.write(str(linha[2].value))
    #categoria
    pyautogui.click(738,441,duration=1.5)
    pyautogui.write(linha[3].value)
    pyautogui.click(593,468,duration=1.5)
    pyautogui.click(670,429,duration=1.5)