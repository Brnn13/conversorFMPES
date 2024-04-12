import pydirectinput
import pandas as pd
import time
import math

excel_file = "ConversorPES.xlsx"
tabela = pd.read_excel(excel_file, sheet_name="resultado")

tof = tabela.iloc[0, 1]
cbola = tabela.iloc[1, 1]
drib = tabela.iloc[2, 1]
cfirme = tabela.iloc[3, 1]
passe = tabela.iloc[4, 1]
palto = tabela.iloc[5, 1]
fin = tabela.iloc[6, 1]
cabec = tabela.iloc[7, 1]
ccolo = tabela.iloc[8, 1]
curv = tabela.iloc[9, 1]
velo = tabela.iloc[10, 1]
acel = tabela.iloc[11, 1]
fchut = tabela.iloc[12, 1]
impul = tabela.iloc[13, 1]
cfisi = tabela.iloc[14, 1]
equi = tabela.iloc[15, 1]
resist = tabela.iloc[16, 1]
tdef = tabela.iloc[17, 1]
desar = tabela.iloc[18, 1]
agress = tabela.iloc[19, 1]

vgk = tabela.iloc[0, 4]
agk = tabela.iloc[1, 4]
fcgk = tabela.iloc[2, 4]
imgk = tabela.iloc[3, 4]
cfgk = tabela.iloc[4, 4]
eqgk = tabela.iloc[5, 4]
resgk = tabela.iloc[6, 4]
talgk = tabela.iloc[7, 4]
firmgk = tabela.iloc[8, 4]
afastgk = tabela.iloc[9, 4]
relexgk = tabela.iloc[10, 4]
alcagk = tabela.iloc[11, 4]

tipo = input("Escolha linha ou gol")
if tipo == "linha":
    pydirectinput.keyDown("alt")
    pydirectinput.press("tab")
    pydirectinput.keyUp("alt")
    # selecionando talento ofensivo
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta Talento Ofensivo
    resultado_tof = math.floor(tof / 10)
    decimais_tof = math.ceil(((tof / 10 - resultado_tof) * 10))
    pydirectinput.keyDown("a")
    for n in range(resultado_tof):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimais_tof):
        pydirectinput.press("right")
    # selecionando controle de bola
    pydirectinput.press("down")
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta controle de bola
    resultado_cbola = math.floor(cbola / 10)
    decimaiscbola = math.ceil(((cbola / 10 - resultado_cbola) * 10))
    pydirectinput.keyDown("a")
    for n in range(resultado_cbola):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimaiscbola):
        pydirectinput.press("right")
    # selecionando drible
    pydirectinput.press("down")
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta drible
    resultado_drib = math.floor((drib) / 10)
    decimaisdrib = math.ceil(((drib) / 10 - resultado_drib) * 10)
    pydirectinput.keyDown("a")
    for n in range(resultado_drib):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimaisdrib):
        pydirectinput.press("right")
    # selecionando condução
    pydirectinput.press("down")
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta condução
    resultado_cfirme = math.floor((cfirme) / 10)
    decimaiscfirme = math.ceil(((cfirme) / 10 - resultado_cfirme) * 10)
    pydirectinput.keyDown("a")
    for n in range(resultado_cfirme):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimaiscfirme):
        pydirectinput.press("right")
    # selecionando passe
    pydirectinput.press("down")
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta passe
    resultado_passe = math.floor(passe / 10)
    decimais_passe = math.ceil((passe / 10 - resultado_passe) * 10)
    pydirectinput.keyDown("a")
    for n in range(resultado_passe):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimais_passe):
        pydirectinput.press("right")
    # selecionando cruz
    pydirectinput.press("down")
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta cruz
    resultado_palto = math.floor((palto) / 10)
    decimaispalto = math.ceil(((palto) / 10 - resultado_palto) * 10)
    pydirectinput.keyDown("a")
    for n in range(resultado_palto):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimaispalto):
        pydirectinput.press("right")
    # selecionando fin
    pydirectinput.press("down")
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta fin
    resultado_fin = math.floor(fin / 10)
    decimais_fin = math.ceil((fin / 10 - resultado_fin) * 10)
    pydirectinput.keyDown("a")
    for n in range(resultado_fin):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimais_fin):
        pydirectinput.press("right")
    # selecionando cabec
    pydirectinput.press("down")
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta cabec
    resultado_cabec = math.floor((cabec) / 10)
    decimais_cabec = math.ceil((cabec / 10 - resultado_cabec) * 10)
    pydirectinput.keyDown("a")
    for n in range(resultado_cabec):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimais_cabec):
        pydirectinput.press("right")
    # selecionando ccolocado
    pydirectinput.press("down")
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta ccolocado
    resultado_ccolo = math.floor((ccolo) / 10)
    decimaisccolo = math.ceil(((ccolo) / 10 - resultado_ccolo) * 10)
    pydirectinput.keyDown("a")
    for n in range(resultado_ccolo):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimaisccolo):
        pydirectinput.press("right")
    # selecionando curva
    pydirectinput.press("down")
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta curva
    resultado_curv = math.floor((curv) / 10)
    decimaiscurv = math.ceil(((curv) / 10 - resultado_curv) * 10)
    pydirectinput.keyDown("a")
    for n in range(resultado_curv):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimaiscurv):
        pydirectinput.press("right")
    # selecionando velocidade
    pydirectinput.press("down")
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta velocidade
    resultado_velo = math.floor((velo) / 10)
    decimaisvelo = math.ceil(((velo) / 10 - resultado_velo) * 10)
    pydirectinput.keyDown("a")
    for n in range(resultado_velo):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimaisvelo):
        pydirectinput.press("right")
    # selecionando aceler
    pydirectinput.press("down")
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta velocidade
    resultado_acel = math.floor((acel) / 10)
    decimaisacel = math.ceil(((acel) / 10 - resultado_acel) * 10)
    pydirectinput.keyDown("a")
    for n in range(resultado_acel):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimaisacel):
        pydirectinput.press("right")
    # selecionando fchute
    pydirectinput.press("down")
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta velocidade
    resultado_fchut = math.floor((fchut) / 10)
    decimaisfchut = math.ceil(((fchut) / 10 - resultado_fchut) * 10)
    pydirectinput.keyDown("a")
    for n in range(resultado_fchut):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimaisfchut):
        pydirectinput.press("right")
    # selecionando impulsão
    pydirectinput.press("down")
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta impulsão
    resultado_impul = math.floor((impul) / 10)
    decimaisimpul = math.ceil(((impul) / 10 - resultado_impul) * 10)
    pydirectinput.keyDown("a")
    for n in range(resultado_impul):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimaisimpul):
        pydirectinput.press("right")
    # selecionando cfisico
    pydirectinput.press("down")
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta cfisico
    resultado_cfisi = math.floor((cfisi) / 10)
    decimaiscfisi = math.ceil(((cfisi) / 10 - resultado_cfisi) * 10)
    pydirectinput.keyDown("a")
    for n in range(resultado_cfisi):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimaiscfisi):
        pydirectinput.press("right")
    # selecionando equilibrio
    pydirectinput.press("down")
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta cfisico
    resultado_equi = math.floor((equi) / 10)
    decimaisequi = math.ceil(((equi) / 10 - resultado_equi) * 10)
    pydirectinput.keyDown("a")
    for n in range(resultado_equi):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimaisequi):
        pydirectinput.press("right")
    # selecionando resist
    pydirectinput.press("down")
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta resist
    resultado_resist = math.floor((resist) / 10)
    decimaisresist = math.ceil(((resist) / 10 - resultado_resist) * 10)
    pydirectinput.keyDown("a")
    for n in range(resultado_resist):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimaisresist):
        pydirectinput.press("right")
    # selecionando tdef
    pydirectinput.press("down")
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta tdef
    resultado_tdef = math.floor((tdef) / 10)
    decimaistdef = math.ceil(((tdef) / 10 - resultado_tdef) * 10)
    pydirectinput.keyDown("a")
    for n in range(resultado_tdef):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimaistdef):
        pydirectinput.press("right")
    # selecionando desar
    pydirectinput.press("down")
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta desar
    resultado_desar = math.floor((desar) / 10)
    decimaisdesar = math.ceil(((desar) / 10 - resultado_desar) * 10)
    pydirectinput.keyDown("a")
    for n in range(resultado_desar):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimaisdesar):
        pydirectinput.press("right")
    # selecionando desar
    pydirectinput.press("down")
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta desar
    resultado_agress = math.floor((agress) / 10)
    decimais_agress = math.ceil((agress / 10 - resultado_agress) * 10)
    pydirectinput.keyDown("a")
    for n in range(resultado_agress):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimais_agress):
        pydirectinput.press("right")
elif tipo == "gol":
    pydirectinput.keyDown("alt")
    pydirectinput.press("tab")
    pydirectinput.keyUp("alt")
    # selecionando vgk
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta vgk
    resultado_vgk = math.floor(vgk / 10)
    decimais_vgk = math.ceil((vgk / 10 - resultado_vgk) * 10)
    pydirectinput.keyDown("a")
    for n in range(resultado_vgk):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimais_vgk):
        pydirectinput.press("right")
    # selecionando agk
    pydirectinput.press("down")
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta agk
    resultado_agk = math.floor(agk / 10)
    decimaisagk = math.ceil((agk / 10 - resultado_agk) * 10)
    pydirectinput.keyDown("a")
    for n in range(resultado_agk):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimaisagk):
        pydirectinput.press("right")
    # selecionando fcgk
    pydirectinput.press("down")
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta fcgk
    resultado_fcgk = math.floor((fcgk) / 10)
    decimaisfcgk = math.ceil(((fcgk) / 10 - resultado_fcgk) * 10)
    pydirectinput.keyDown("a")
    for n in range(resultado_fcgk):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimaisfcgk):
        pydirectinput.press("right")
    # selecionando imgk
    pydirectinput.press("down")
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta imgk
    resultado_imgk = math.floor((imgk) / 10)
    decimaisimgk = math.ceil(((imgk) / 10 - resultado_imgk) * 10)
    pydirectinput.keyDown("a")
    for n in range(resultado_imgk):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimaisimgk):
        pydirectinput.press("right")
    # selecionando cfgk
    pydirectinput.press("down")
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta cfgk
    resultado_cfgk = math.floor((cfgk) / 10)
    decimais_cfgk = math.ceil(((cfgk) / 10 - resultado_cfgk) * 10)
    pydirectinput.keyDown("a")
    for n in range(resultado_cfgk):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimais_cfgk):
        pydirectinput.press("right")
    # selecionando eqgk
    pydirectinput.press("down")
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta eqgk
    resultado_eqgk = math.floor((eqgk) / 10)
    decimais_eqgk = math.ceil(((eqgk) / 10 - resultado_eqgk) * 10)
    pydirectinput.keyDown("a")
    for n in range(resultado_eqgk):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimais_eqgk):
        pydirectinput.press("right")
    # selecionando resgk
    pydirectinput.press("down")
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta resgk
    resultado_resgk = math.floor((resgk) / 10)
    decimais_resgk = math.ceil(((resgk) / 10 - resultado_resgk) * 10)
    pydirectinput.keyDown("a")
    for n in range(resultado_resgk):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimais_resgk):
        pydirectinput.press("right")
    # selecionando talgk
    for n in range(4):
        pydirectinput.press("down")
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta talgk
    resultado_talgk = math.floor((talgk) / 10)
    decimais_talgk = math.ceil(((talgk) / 10 - resultado_talgk) * 10)
    pydirectinput.keyDown("a")
    for n in range(resultado_talgk):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimais_talgk):
        pydirectinput.press("right")
    # selecionando firmgk
    pydirectinput.press("down")
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta firmgk
    resultado_firmgk = math.floor((firmgk) / 10)
    decimais_firmgk = math.ceil((firmgk / 10 - resultado_firmgk) * 10)
    pydirectinput.keyDown("a")
    for n in range(resultado_firmgk):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimais_firmgk):
        pydirectinput.press("right")
    # selecionando afastgk
    pydirectinput.press("down")
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta afastgk
    resultado_afastgk = math.floor((afastgk) / 10)
    decimais_afastgk = math.ceil((afastgk / 10 - resultado_afastgk) * 10)
    pydirectinput.keyDown("a")
    for n in range(resultado_afastgk):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimais_afastgk):
        pydirectinput.press("right")
    # selecionando relexgk
    pydirectinput.press("down")
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta relexgk
    resultado_relexgk = math.floor((relexgk) / 10)
    decimais_relexgk = math.ceil((relexgk / 10 - resultado_relexgk) * 10)
    pydirectinput.keyDown("a")
    for n in range(resultado_relexgk):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimais_relexgk):
        pydirectinput.press("right")
    # selecionando alcagk
    pydirectinput.press("down")
    pydirectinput.keyDown("a")
    pydirectinput.keyDown("left")
    time.sleep(0.8)
    pydirectinput.keyUp("a")
    pydirectinput.keyUp("left")
    # Conta alcagk
    resultado_alcagk = math.floor((alcagk) / 10)
    decimais_alcagk = math.ceil((alcagk / 10 - resultado_alcagk) * 10)
    pydirectinput.keyDown("a")
    for n in range(resultado_alcagk):
        pydirectinput.press("right")
    pydirectinput.keyUp("a")
    for n in range(decimais_alcagk):
        pydirectinput.press("right")

else:
    print("erro")
