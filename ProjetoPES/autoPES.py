import pydirectinput
import pandas as pd
import time
import math

excel_file = "ConversorPES.xlsx"
tabela = pd.read_excel(excel_file, sheet_name="resultado")
linha = 0
coluna = 1
valor = tabela.iloc[linha, coluna]
linhagk = 0
colunagk = 4
valorgk = tabela.iloc[linhagk, colunagk]
def alttab():
    pydirectinput.keyDown("alt")
    pydirectinput.press("tab")
    pydirectinput.keyUp("alt")
def resetaratt():
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
def contas():
    resultado_valor = math.floor(valorconta / 10)
    decimais_valor = math.floor(((valorconta / 10 - resultado_valor) * 10))
    pydirectinput.keyDown("a")
    for n in range(resultado_valor):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimais_valor):
        pydirectinput.press("right")
def contasgk():
    resultadogk = math.floor(valorgk / 10)
    decimaisgk = math.floor(((valorgk / 10 - resultadogk) * 10))
    pydirectinput.keyDown("a")
    for n in range(resultadogk):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimaisgk):
        pydirectinput.press("right")
tipo = input("Escolha linha ou gol")
if tipo == "linha":
    alttab()
    for _ in range(20):
        resetaratt()
        contas()
        pydirectinput.press("down")
        linha += 1
elif tipo == "gol":
    alttab()
    for _ in range(7):
        resetaratt()
        contasgk()
        pydirectinput.press("down")
        linhagk += 1
    for _ in range(3):
        pydirectinput.press("down")
    for _ in range(5):
        resetaratt()
        contasgk()
        pydirectinput.press("down")
        linhagk += 1
else:
    print("erro")