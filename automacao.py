from numbers import Integral
import pyautogui as py
import pyautogui
import pytesseract as ocr
from PIL import Image, ImageDraw, ImageFont
import PIL.Image
import time
import pygetwindow
import cx_Oracle
import tkinter as tk
from tkinter import messagebox
from tkinter import *
from tkinter import ttk
import locale
import cv2
import numpy as np
import openai
import win32gui
import os

openai.api_key = "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"

dsn = cx_Oracle.makedsn('10.1.0.137', 1521, 'dbfesocons')
user = 'rm'
password = 'f/M701iv_LoAE1@'

dados_conexao = cx_Oracle.connect(user, password, dsn)

conexao = cx_Oracle.connect(user, password, dsn)
cursor = conexao.cursor()

window = tk.Tk()
window.title('RPA by Iago Chiapetta')
window.geometry('400x300')
#Recurso que impede que o usuário redimensione a janela
window.resizable(width=False, height=False)
w = Label(window, text='Escolha entre as opções abaixo: ', font=35)
w.place(x=90, y=10)

Checkbutton1 = IntVar()
Checkbutton2 = IntVar()

#Recurso que impossibilita que dois CheckButtons sejam selecionados ao mesmo tempo.
def clear2():
    Checkbutton2.set(0)

def clear1():
    Checkbutton1.set(0)

def vazio():
    messagebox.showwarning("Alerta de Campos Vazios", "Você precisa preencher todos os campos para iniciar o processamento.")

#CheckBox da Interface do App 
btn1 = Checkbutton(window, text="Baixa", variable=Checkbutton1, onvalue=1, offvalue=0, height=2, width=10, command=clear2)
btn2 = Checkbutton(window, text="Perda", variable=Checkbutton2, onvalue=1, offvalue=0, height=2, width=20, command=clear1)


btn1.pack_forget()#x=105, y=55)
btn2.place(x=116, y=85)

textoPeriodo = Label(window, text="Período: ", bg="#d3d3d3")
textoPeriodo.place(x=10, y=200)
entradaPeriodoMes = Entry(window)
entradaPeriodoMes.place(x=60, y=200, width=15)
barra = Label(window, text="/", bg="#d3d3d3")
barra.place(x=75, y=200)
entradaPeriodoAno = Entry(window)
entradaPeriodoAno.place(x=85, y=200, width=30)

textoConvenio = Label(window, text="Convenio: ", bg="#d3d3d3")
textoConvenio.place(x=150, y=200)
caixaSelecao = ttk.Combobox(window, 
values=[
    "AMIL",
    "ASSIM",
    "BRADESCO",
    "CAC",
    "CEF",
    "CORPO DE BOMBEIROS DO ESTADO DO RIO DE JANEIRO",
    "FUSMA",
    "GOLDEN CROSS",
    "MEDISERVICE",
    "PETROBRAS",
    "POSTAL SAUDE",
    "SUL AMERICA",
    "UNIMED"
]
)
caixaSelecao.place(x=210, y=200, width=170)

labelFatura = Label(window, text="Fatura: ", bg="#d3d3d3")
labelFatura.place(x=10 ,y=240)
entradaFatura = Entry(window)
entradaFatura.place(x=55, y=240, width=40)

textoUnidade = Label(window, text="Unidade: ", bg="#d3d3d3")
textoUnidade.place(x=160, y=240)
caixaUnidade = ttk.Combobox(window, values=[
    "AMBULATÓRIO (FATURAMENTO)",
    "HOSPITAL DAS CLÍNICAS DE TERESÓPOLIS",
    "MATERNIDADE DO HCTCO (FATURAMENTO)"
])
caixaUnidade.place(x=211, y=240, width=170)


#Aviso de recurso não disponível
def baixa():
    messagebox.showwarning("Recurso Futuro", "O Recurso Selecionado Ainda está sendo desenvolvido e estará disponível em breve. Fique tranquilo, nós avisaremos.")

def automacao():
    py.alert("A automação está sendo iniciada")

    py.PAUSE = 1.2

    x, y = py.locateCenterOnScreen('itens/tesouraria.png')
    print(x, y)
    py.click(x, y)
    x, y = py.locateCenterOnScreen('itens/controlerecebimentos.png')
    py.click(x, y)
    time.sleep(5)
    title = "Controle de Recebimentos (RDESKAMB.UNIFESO.LAN)"
    try:
        window = pygetwindow.getWindowsWithTitle(title)[0]
        x,y = py.locateCenterOnScreen('itens/fatura.png')
        py.click(x, y)

        x, y = py.locateCenterOnScreen('itens/competencia.png')
        py.click(x, y)
        py.write(str(entradaPeriodoMes.get()) + str(entradaPeriodoAno.get()), interval=0.3)

        x, y = py.locateCenterOnScreen('itens/convenio.png')
        py.click(x, y)

        valuePeriodoMes = entradaPeriodoMes.get()
        valuePeriodoAno = entradaPeriodoAno.get()
        valuePeriodo = str(valuePeriodoMes) + '/' + str(valuePeriodoAno)
        conexao = cx_Oracle.connect(user, password, dsn)
        cursor = conexao.cursor()
        sql = "SELECT COUNT(SIGLA) FROM SZCADGERAL WHERE CODCOLIGADA = 1.000000 AND CLASSIFICACAO = 'C' AND SITUACAO = 'A' AND TIPOFATURAMENTO = 1 AND CODGERAL IN (SELECT CODCOMPRADOR CODGERAL FROM SZFATURACONVENIO WHERE CODCOLIGADA = 1.000000 AND (TO_CHAR(DATAGERACAO,'MM')|| '/' || TO_CHAR(DATAGERACAO,'YYYY')) = '"+str(valuePeriodo)+"') ORDER BY SIGLA"
        cursor.execute(sql)
        tentativas = cursor.fetchone()[0]
        print(tentativas)
        max_tries = tentativas
        tries = 0
        if caixaSelecao.get() == "AMIL":
            codPlano = 5
            while tries != max_tries:
                try:
                    x, y = py.locateCenterOnScreen('itens/amil.png')
                    py.click(x, y)
                    break
                except Exception as e:
                    descer = py.locateCenterOnScreen('itens/descer.png')
                    py.click(descer)
        elif caixaSelecao.get() == "ASSIM":
            codPlano = 7
            while tries != max_tries:
                try:
                    x, y = py.locateCenterOnScreen('itens/assim.png')
                    py.click(x, y)
                    break
                except Exception as e:
                    descer = py.locateCenterOnScreen('itens/descer.png')
                    py.click('descer')
        elif caixaSelecao.get() == "BR DISTRIBUIDORA":
            codPlano = 1678
            while tries != max_tries:
                try:
                    x, y = py.locateCenterOnScreen('itens/amil.png')
                    py.click(x, y)
                    break
                except Exception as e:
                        descer = py.locateCenterOnScreen('itens/descer.png')
                        py.click(descer)
        elif caixaSelecao.get() == "BRADESCO":
            codPlano = 27
            while tries != max_tries:
                try:
                    x, y = py.locateCenterOnScreen('itens/bradesco.png')
                    py.click(x, y)
                    break
                except Exception as e:
                    descer = py.locateCenterOnScreen('itens/descer.png')
                    py.click(descer)
        elif caixaSelecao.get() == "CAC":
            codPlano = 10
            while tries != max_tries:
                try:
                    x, y = py.locateCenterOnScreen('itens/cac.png')
                    py.click(x, y)
                    break
                except Exception as e:
                    descer = py.locateCenterOnScreen('itens/descer.png')
                    py.click(descer)
        elif caixaSelecao.get() == "CEF":
            codPlano = 1151
            while tries != max_tries:
                try:
                    x, y = py.locateCenterOnScreen('itens/cef.png')
                    py.click(x, y)
                    break
                except Exception as e:
                    descer = py.locateCenterOnScreen('itens/descer.png')
                    py.click(descer)
        elif caixaSelecao.get() == "CORPO DE BOMBEIROS DO ESTADO DO RIO DE JANEIRO":
            codPlano = 2879
            while tries != max_tries:
                try:
                    x, y = py.locateCenterOnScreen('itens/corpobomb.png')
                    py.click(x, y)
                    break
                except Exception as e:
                    descer = py.locateCenterOnScreen('itens/descer.png')
                    py.click(descer)
        elif caixaSelecao.get() == "FUSMA":
            codPlano = 5383
            while tries != max_tries:
                try:
                    x, y = py.locateCenterOnScreen('itens/fusma.png')
                    py.click(x, y)
                    break
                except Exception as e:
                    descer = py.locateCenterOnScreen('itens/descer.png')
                    py.click(descer)
        elif caixaSelecao.get() == "GOLDEN CROSS":
            codPlano = 20
            while tries != max_tries:
                try:
                    x, y = py.locateCenterOnScreen('itens/golden.png')
                    py.click(x, y)
                    break
                except Exception as e:
                    descer = py.locateCenterOnScreen('itens/descer.png')
                    py.click(descer)
        elif caixaSelecao.get() == "MEDISERVICE":
            while tries != max_tries:
                try:
                    x, y = py.locateCenterOnScreen('itens/mediservice.png')
                    py.click(x, y)
                    break
                except Exception as e:
                    descer = py.locateCenterOnScreen('itens/descer.png')
                    py.click(descer)
                    codPlano = 24            
        elif caixaSelecao.get() == "PETROBRAS":
            while tries != max_tries:
                try:
                    x, y = py.locateCenterOnScreen('itens/petrobras.png')
                    py.click(x, y)
                    break
                except Exception as e:
                        descer = py.locateCenterOnScreen('itens/descer.png')
                        py.click(descer)
                        codPlano = 26
        elif caixaSelecao.get() == "POSTAL SAUDE":
            codPlano = 4113
            while tries != max_tries:
                try:
                    x, y = py.locateCenterOnScreen('itens/postal.png')
                    py.click(x, y)
                    break
                except Exception as e:
                    descer = py.locateCenterOnScreen('itens/descer.png')
        elif caixaSelecao.get() == "SUL AMERICA":
            codPlano = 31
            while tries != max_tries:
                try:
                    x, y = py.locateCenterOnScreen('itens/sulamerica.png')
                    py.click(x, y)
                    break
                except Exception as e:
                    descer = py.locateCenterOnScreen('itens/descer.png')
                    py.click(descer)
        elif caixaSelecao.get() == "UNIMED":
            codPlano = '4131'
            while tries != max_tries:
                try:
                    x, y = py.locateCenterOnScreen('itens/unimed.png')
                    py.click(x, y)
                    break
                except Exception as e:
                    descer = py.locateCenterOnScreen('itens/descer.png')
                    py.click(descer)
        else:
            print("Plano de saúde não encontrado")
    except:
            print("A janela de Recebimentos não saiu.")
    pointx, pointy = py.locateCenterOnScreen('itens/faturamento.png')
    print(pointx, pointy)
    py.click(pointx, pointy)

    if caixaUnidade.get() == "HOSPITAL DAS CLÍNICAS DE TERESÓPOLIS":
        codHospital = 3
        x, y = py.locateCenterOnScreen('itens/hct.png')
        py.click(x, y)
    elif caixaUnidade.get() == "AMBULATÓRIO (FATURAMENTO)":
        codHospital = 13
        x, y = py.locateCenterOnScreen('itens/ambulatorio.png')
        py.click(x, y)
    elif caixaUnidade.get() == "MATERNIDADE DO HCTCO (FATURAMENTO)":
        codHospital = 14
        x, y = py.locateCenterOnScreen('itens/maternidade.png')
        py.click(x, y)
    else:
        print("Local de atendimento não encontrado")

    pesquisa = py.locateCenterOnScreen('itens/pesquisa.png')
    py.click(pesquisa)
                
    fatura = str(entradaFatura.get())
    time.sleep(5)
    screenshot = pyautogui.screenshot()
    custom_config = r'--oem 1 --psm 11 tessedit_char_whitelist=1234567890'
    text = ocr.image_to_data(screenshot, output_type='dict', lang='por', config=custom_config)
    for i in range(len(text['text'])):
        if text['text'][i] == fatura:
            x, y, w, h = text['left'][i], text['top'][i], text['width'][i], text['height'][i]
            screen_x, screen_y = x + w / 2, y + h / 2
            py.click(screen_x, screen_y)
            break
    else:
        print("Texto não encontrado")
        i < 0
        while i < 10:
            rolar = py.locateCenterOnScreen('itens/rolar.png')
            py.click(rolar)
            custom_config = r'--oem 1 --psm 11 tessedit_char_whitelist=1234567890'
            text = ocr.image_to_data(thresh, output_type='dict', lang='por', config=custom_config)
            i += 1
            continue


    x, y = py.locateCenterOnScreen('itens/contas.png')
    py.click(x, y)

    value = int(entradaFatura.get())
    valuePeriodoMes = entradaPeriodoMes.get()
    valuePeriodoAno = entradaPeriodoAno.get()
    valorCompetencia = str(valuePeriodoMes) + '/' + str(valuePeriodoAno)

    #Consulta
    contador = 0
    while True:
        sql = "SELECT 'CONTA INICIAL' TIPO_CONTA,F.SIGLA,A.DATAENTRADA,A.DATASAIDA,B.NUMEROREMESSA REMESSA, to_char(G.DATAGERACAO, 'MM/YYYY') GERACAO,E.PRONTUARIO,E.CODPACIENTE,A.CODATENDIMENTO,E.NOMEPACIENTE PACIENTE,B.IDUNIDFAT,A.CODCOMPRADOR, A.SEQUENCIALCONTA FROM SZATENDIMENTO A INNER JOIN SZPARCIALATEND B ON  A.CODCOLIGADA = B.CODCOLIGADA     AND A.NUMEROCONTA = B.NUMEROCONTA  AND A.SEQUENCIALCONTA = B.SEQUENCIALCONTA INNER JOIN SZPACIENTE E ON  A.CODCOLIGADA = E.CODCOLIGADA  AND A.CODPACIENTE = E.CODPACIENTE INNER JOIN SZCADGERAL F ON A.CODCOLIGADA = F.CODCOLIGADA  AND B.CODCONVENIO = F.CODGERAL INNER JOIN SZFATURACONVENIO G ON A.CODCOLIGADA = G.CODCOLIGADA  AND B.CODCONVENIO = G.CODCOMPRADOR AND B.CODCOLIGADA = G.CODCOLIGADA  AND B.NUMEROREMESSA = G.NUMEROREMESSA WHERE A.CODCOLIGADA = 1 and  to_char(G.DATAGERACAO, 'MM/YYYY')= '"+str(valorCompetencia)+"' and g.numeroremessa = "+str(value)+" and g.codcomprador = "+str(codPlano)+" and B.IDUNIDFAT = "+str(codHospital)+" AND A.CODCOMPRADOR NOT IN (494, 25, 3087) UNION ALL SELECT 'CONTA DE RECURSO' TIPO_CONTA, F.SIGLA, A.DATAENTRADA, A.DATASAIDA, B.NUMEROREMESSA  REMESSA, to_char(G.DATAGERACAO, 'MM/YYYY') GERACAO, E.PRONTUARIO, E.CODPACIENTE, A.CODATENDIMENTO, E.NOMEPACIENTE  PACIENTE, B.IDUNIDFAT, A.CODCOMPRADOR, A.SEQUENCIALCONTA FROM SZATENDIMENTOREC A INNER JOIN SZPARCIALATEND B ON  A.CODCOLIGADA = B.CODCOLIGADA AND A.NUMEROCONTA = B.NUMEROCONTA AND A.SEQUENCIALCONTA = B.SEQUENCIALCONTA INNER JOIN SZPACIENTE E ON  A.CODCOLIGADA = E.CODCOLIGADA  AND A.CODPACIENTE = E.CODPACIENTE INNER JOIN SZCADGERAL F ON A.CODCOLIGADA = F.CODCOLIGADA  AND B.CODCONVENIO = F.CODGERAL INNER JOIN SZFATURACONVENIO G ON A.CODCOLIGADA = G.CODCOLIGADA  AND B.CODCONVENIO = G.CODCOMPRADOR  AND B.CODCOLIGADA = G.CODCOLIGADA  AND B.NUMEROREMESSA = G.NUMEROREMESSA WHERE A.CODCOLIGADA = 1 AND B.STATUS = 'F' and  to_char(G.DATAGERACAO, 'MM/YYYY')= '"+str(valorCompetencia)+"' and g.numeroremessa = "+str(value)+" and g.codcomprador = "+str(codPlano)+" and B.IDUNIDFAT = "+str(codHospital)+" AND G.STATUSPROTOCOLO = 'F' AND A.CODCOMPRADOR NOT IN (494, 25, 3087) ORDER BY PACIENTE"
        cursor.execute(sql)
        contaRecurso = cursor.fetchall()[contador]
        numeroConta = str(contaRecurso[12])
        tipo_conta = contaRecurso[0]
        print(tipo_conta, numeroConta, contador)
        sql2 = "SELECT COUNT(PACIENTE) FROM (SELECT 'CONTA INICIAL' TIPO_CONTA,F.SIGLA,A.DATAENTRADA,A.DATASAIDA,B.NUMEROREMESSA REMESSA, to_char(G.DATAGERACAO, 'MM/YYYY') GERACAO,E.PRONTUARIO,E.CODPACIENTE,A.CODATENDIMENTO,E.NOMEPACIENTE PACIENTE,B.IDUNIDFAT,A.CODCOMPRADOR, A.SEQUENCIALCONTA FROM SZATENDIMENTO A INNER JOIN SZPARCIALATEND B ON  A.CODCOLIGADA = B.CODCOLIGADA     AND A.NUMEROCONTA = B.NUMEROCONTA  AND A.SEQUENCIALCONTA = B.SEQUENCIALCONTA INNER JOIN SZPACIENTE E ON  A.CODCOLIGADA = E.CODCOLIGADA  AND A.CODPACIENTE = E.CODPACIENTE INNER JOIN SZCADGERAL F ON A.CODCOLIGADA = F.CODCOLIGADA  AND B.CODCONVENIO = F.CODGERAL INNER JOIN SZFATURACONVENIO G ON A.CODCOLIGADA = G.CODCOLIGADA  AND B.CODCONVENIO = G.CODCOMPRADOR AND B.CODCOLIGADA = G.CODCOLIGADA  AND B.NUMEROREMESSA = G.NUMEROREMESSA WHERE A.CODCOLIGADA = 1 and  to_char(G.DATAGERACAO, 'MM/YYYY')= '"+str(valorCompetencia)+"' and g.numeroremessa = "+str(value)+" and g.codcomprador = "+str(codPlano)+" and B.IDUNIDFAT = "+str(codHospital)+" AND A.CODCOMPRADOR NOT IN (494, 25, 3087) UNION ALL SELECT 'CONTA DE RECURSO' TIPO_CONTA, F.SIGLA, A.DATAENTRADA, A.DATASAIDA, B.NUMEROREMESSA  REMESSA, to_char(G.DATAGERACAO, 'MM/YYYY') GERACAO, E.PRONTUARIO, E.CODPACIENTE, A.CODATENDIMENTO, E.NOMEPACIENTE  PACIENTE, B.IDUNIDFAT, A.CODCOMPRADOR, A.SEQUENCIALCONTA FROM SZATENDIMENTOREC A INNER JOIN SZPARCIALATEND B ON  A.CODCOLIGADA = B.CODCOLIGADA AND A.NUMEROCONTA = B.NUMEROCONTA AND A.SEQUENCIALCONTA = B.SEQUENCIALCONTA INNER JOIN SZPACIENTE E ON  A.CODCOLIGADA = E.CODCOLIGADA  AND A.CODPACIENTE = E.CODPACIENTE INNER JOIN SZCADGERAL F ON A.CODCOLIGADA = F.CODCOLIGADA  AND B.CODCONVENIO = F.CODGERAL INNER JOIN SZFATURACONVENIO G ON A.CODCOLIGADA = G.CODCOLIGADA  AND B.CODCONVENIO = G.CODCOMPRADOR  AND B.CODCOLIGADA = G.CODCOLIGADA  AND B.NUMEROREMESSA = G.NUMEROREMESSA WHERE A.CODCOLIGADA = 1 AND B.STATUS = 'F' and  to_char(G.DATAGERACAO, 'MM/YYYY')= '"+str(valorCompetencia)+"' and g.numeroremessa = "+str(value)+" and g.codcomprador = "+str(codPlano)+" and B.IDUNIDFAT = "+str(codHospital)+" AND G.STATUSPROTOCOLO = 'F' AND A.CODCOMPRADOR NOT IN (494, 25, 3087) ORDER BY PACIENTE)"
        cursor.execute(sql2)
        numeroTentativas = cursor.fetchone()[0]
        print(numeroTentativas)
        time.sleep(3)
        if tipo_conta == "CONTA INICIAL":
            def find_window_by_title(title):
                handle = win32gui.FindWindow(None, title)
                if handle == 0:
                    print("Janela não encontrada.")
                    return None
                rect = win32gui.GetWindowRect(handle)
                return rect
            rect = find_window_by_title("Controle de Recebimentos (RDESKAMB.UNIFESO.LAN)")
            if rect:
                left, top, right, bottom = rect
                print("Posição da janela: ({}, {}), ({}, {})".format(left, top, right, bottom))
                screenshot1 = pyautogui.screenshot(region=(left, top, right-left, bottom-top))
            for i in range(len(text['text'])):
                if text['text'][i] == numeroConta:
                    x, y, w, h = text['left'][i], text['top'][i], text['width'][i], text['height'][i]
                    screen_x, screen_y = x + w / 2, y + h / 2
                    py.doubleClick(screen_x, screen_y)
                    break
            else:
                print("Texto não encontrado")
                i = 0
                while i < numeroTentativas:
                    rolar = py.locateCenterOnScreen('itens/rolar.png')
                    py.click(rolar)
                    screenshot1 = pyautogui.screenshot()
                    custom_config = r'--oem 1 --psm 11 --dpi 150 tessedit_char_whitelist=1234567890'
                    text = ocr.image_to_data(screenshot1, output_type='dict',lang='por', config=custom_config)
                    i += 1 
                    continue

        elif tipo_conta == "CONTA DE RECURSO":
            contador += 1
            continue

        x, y = py.locateCenterOnScreen('itens/recebimento.png')
        py.click(x, y)

        data = py.locateCenterOnScreen('itens/data.png')
        py.click(data)

        valor = py.locateCenterOnScreen('itens/valor.png')
        py.doubleClick(valor)
        py.press('backspace')

        perda = py.locateCenterOnScreen('itens/perda.png')
        py.doubleClick(perda)
        py.press('backspace')

        #Algoritmo para achar a perda e passar para o ocr
        screenshot = py.screenshot()
        gray = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2GRAY)
        item = cv2.imread("itens/valor_perda.png", 0)
        result = cv2.matchTemplate(gray, item, cv2.TM_CCOEFF_NORMED)
        _, _, _, loc = cv2.minMaxLoc(result)
        x, y = loc[0], loc[1]
        valorPerda = py.screenshot('testes.png', region=(x,y,65, 17))
        image = cv2.imread('testes.png')
        width = 97
        height = 24
        image = cv2.resize(image, (width, height))
        image = cv2.fastNlMeansDenoisingColored(image, None, 10, 10, 7, 21)
        image = cv2.addWeighted(image, 1.0, np.zeros(image.shape, image.dtype), 0, 50)
        gray = cv2.Canny(image, 100, 200)
        gray = cv2.Laplacian(image, cv2.CV_64F)
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        #gray = cv2.convertScaleAbs(image, alpha=1.5, beta=0)
        gray = cv2.bilateralFilter(gray,9,75,75)
        gray = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)
        gray = cv2.equalizeHist(gray)
        gray = cv2.GaussianBlur(gray, (1,1), 0)
        thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
        custom_config = r'--oem 1 --psm 11 tessedit_char_whitelist=1234567890'
        phrase1 = ocr.image_to_string(thresh, lang='por', config=custom_config)
        phrase1 = phrase1.replace(" ", "")
        locale.setlocale(locale.LC_ALL, "Portuguese_Brazil.1252")
        valor = locale.atof(phrase1)
        valor_string = str(valor)
        print(valor_string)
        py.write(valor_string.replace(".", ","), interval=0.3)

        gravar = py.locateCenterOnScreen('itens/gravar.png')
        py.click(gravar)        
        py.write('1')

        py.press('tab')
        py.press('tab')
        py.press('tab')
        py.press('enter')
        title = "Erro (RDESKAMB.UNIFESO.LAN)"
        time.sleep(10)
        py.press('enter')
        #ok = py.locateCenterOnScreen('itens/ok.png')
        #py.click(ok)
        cancelar = py.locateCenterOnScreen('itens/cancelar.png')
        py.click(cancelar)
        contador += 1


def recuperarLista():
    if Checkbutton1.get():
        baixa()
        exit()
    elif Checkbutton2.get() == 1:
        automacao()
    else:
        vazio()
        
btn = Button(window, text='Executar automação', command=recuperarLista, width=30, bg="#d3d3d3")
btn.place(x=90, y=140)

cursor.close()
window.mainloop()