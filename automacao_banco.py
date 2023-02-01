from numbers import Integral
import pyautogui as py
import pyautogui
import pytesseract as ocr
from PIL import Image
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
import logging

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
entradaOrdemFatura = Entry(window)
entradaOrdemFatura.place(x=55, y=240, width=40)

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
            except Exception as e:
                descer = py.locateCenterOnScreen('itens/descer.png')
                py.click(descer)
    elif caixaSelecao.get() == "ASSIM":
        codPlano = 7
        while tries != max_tries:
            try:
                x, y = py.locateCenterOnScreen('itens/assim.png')
                py.click(x, y)
            except Exception as e:
                descer = py.locateCenterOnScreen('itens/descer.png')
                py.click('descer')
    elif caixaSelecao.get() == "BR DISTRIBUIDORA":
        codPlano = 1678
        while tries != max_tries:
            try:
                x, y = py.locateCenterOnScreen('itens/amil.png')
                py.click(x, y)
            except Exception as e:
                descer = py.locateCenterOnScreen('itens/descer.png')
                py.click(descer)
    elif caixaSelecao.get() == "BRADESCO":
        codPlano = 27
        while tries != max_tries:
            try:
                x, y = py.locateCenterOnScreen('itens/bradesco.png')
                py.click(x, y)
            except Exception as e:
                descer = py.locateCenterOnScreen('itens/descer.png')
                py.click(descer)
    elif caixaSelecao.get() == "CAC":
        codPlano = 10
        while tries != max_tries:
            try:
                x, y = py.locateCenterOnScreen('itens/cac.png')
                py.click(x, y)
            except Exception as e:
                descer = py.locateCenterOnScreen('itens/descer.png')
                py.click(descer)
    elif caixaSelecao.get() == "CEF":
        codPlano = 1151
        while tries != max_tries:
            try:
                x, y = py.locateCenterOnScreen('itens/cef.png')
                py.click(x, y)
            except Exception as e:
                descer = py.locateCenterOnScreen('itens/descer.png')
                py.click(descer)
    elif caixaSelecao.get() == "CORPO DE BOMBEIROS DO ESTADO DO RIO DE JANEIRO":
        codPlano = 2879
        while tries != max_tries:
            try:
                x, y = py.locateCenterOnScreen('itens/corpobomb.png')
                py.click(x, y)
            except Exception as e:
                descer = py.locateCenterOnScreen('itens/descer.png')
                py.click(descer)
    elif caixaSelecao.get() == "FUSMA":
        codPlano = 5383
        while tries != max_tries:
            try:
                x, y = py.locateCenterOnScreen('itens/fusma.png')
                py.click(x, y)
            except Exception as e:
                descer = py.locateCenterOnScreen('itens/descer.png')
                py.click(descer)
    elif caixaSelecao.get() == "GOLDEN CROSS":
        codPlano = 20
        while tries != max_tries:
            try:
                x, y = py.locateCenterOnScreen('itens/golden.png')
                py.click(x, y)
            except Exception as e:
                descer = py.locateCenterOnScreen('itens/descer.png')
                py.click(descer)
    elif caixaSelecao.get() == "MEDISERVICE":
        while tries != max_tries:
            try:
                x, y = py.locateCenterOnScreen('itens/mediservice.png')
                py.click(x, y)
            except Exception as e:
                descer = py.locateCenterOnScreen('itens/descer.png')
                py.click(descer)
        codPlano = 24            
    elif caixaSelecao.get() == "PETROBRAS":
        while tries != max_tries:
            try:
                x, y = py.locateCenterOnScreen('itens/petrobras.png')
                py.click(x, y)
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
            except Exception as e:
                descer = py.locateCenterOnScreen('itens/descer.png')
    elif caixaSelecao.get() == "SUL AMERICA":
        codPlano = 31
        while tries != max_tries:
            try:
                x, y = py.locateCenterOnScreen('itens/sulamerica.png')
                py.click(x, y)
            except Exception as e:
                descer = py.locateCenterOnScreen('itens/descer.png')
                py.click(descer)
    elif caixaSelecao.get() == "UNIMED":
        codPlano = '4131'
        while tries != max_tries:
            try:
                x, y = py.locateCenterOnScreen('itens/unimed.png')
                py.click(x, y)
            except Exception as e:
                descer = py.locateCenterOnScreen('itens/descer.png')
                py.click(descer)
    else:
        print("Nenhuma informação encontrada")

faturamento = py.locateCenterOnScreen('itens/faturamento.png')
print(faturamento)
py.click(faturamento)

if caixaUnidade.get() == "HOSPITAL DAS CLÍNICAS DE TERESÓPOLIS":
    codHospital = 3
elif caixaUnidade.get() == "AMBULATÓRIO (FATURAMENTO)":
    codHospital = 13
                    
elif caixaUnidade.get() == "MATERNIDADE DO HCTCO (FATURAMENTO)":
   codHospital = 14
                    

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