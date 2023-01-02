import pyautogui as py
import pyautogui
import pytesseract as ocr
from PIL import Image
import PIL.Image
import time
import pygetwindow
import pyodbc
import pyttsx3
import tkinter as tk
from tkinter import messagebox
from tkinter import *
from tkinter import ttk

#Instanciando Tkinter e criando interface de usuário
window = tk.Tk()
window.title('Automatizador de Tarefas HCTCO')
window.geometry('400x300')
window.resizable(width=False, height=False)
#Recurso que impede que o usuário redimensione a janela
w = Label(window, text='Escolha entre as opções abaixo: ', font=35)
w.place(x=90, y=10)

Checkbutton1 = IntVar()
Checkbutton2 = IntVar()

#Recurso que impossibilita que dois CheckButtons sejam selecionados ao mesmo tempo.
def clear2():
    Checkbutton2.set(0)

def clear1():
    Checkbutton1.set(0)


btn1 = Checkbutton(window, text="Baixa", variable=Checkbutton1, onvalue=1, offvalue=0, height=2, width=10, command=clear2)
btn2 = Checkbutton(window, text="Controle de Recebimentos", variable=Checkbutton2, onvalue=1, offvalue=0, height=2, width=20, command=clear1)


btn1.place(x=105, y=55)
btn2.place(x=126, y=85)

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
caixaSelecao = ttk.Combobox(window, values=[
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
])
caixaSelecao.place(x=210, y=200, width=170)

textoUnidade = Label(window, text="Unidade de Faturamento: ", bg="#d3d3d3")
textoUnidade.place(x=70, y=240)
caixaUnidade = ttk.Combobox(window, values=[
    "AMBULATÓRIO (FATURAMENTO)",
    "HOSPITAL DAS CLÍNICAS DE TERESÓPOLIS",
    "MATERNIDADE DO HCTCO (FATURAMENTO)"
])
caixaUnidade.place(x=211, y=240, width=170)

#Motor de sintetização de voz
engine = pyttsx3.init()
voices = engine.getProperty("voices")
engine.setProperty('voice', voices[0].id)
engine.setProperty("voice", "brazil")
engine.setProperty("rate", 225)
engine.runAndWait()

def baixa():
    messagebox.showwarning("Recurso Futuro", "O Recurso Selecionado Ainda está sendo desenvolvido e estará disponível em breve. Fique tranquilo, nós avisaremos.")


dados_conexao = ('DRIVER={Devart ODBC Driver for Oracle};Host=10.1.0.137;Port=1521;SID=DBFESOCONS;UID=GTIC_078463;Password=gtic_sup23_078463;Direct=True')
#Caminho até a automação
def automacao():
    inicio = time.time()
    py.alert("Atenção! A automação está sendo iniciada.")
    engine.say("Atenção! A automação está sendo iniciada.")
    engine.runAndWait()
    py.PAUSE = 1.2
    res = py.size()
    if res.width == 1366 and res.height == 768:
        try:
            #Acessar onde ficam os dados a serem obtidos
            py.moveTo(784, 26)
            py.click()
            py.moveTo(782, 142)
            py.click()
            title = "Controle de Recebimentos"
            window = pygetwindow.getWindowsWithTitle(title)[0]
            if window.title == "Controle de Recebimentos (RDESKAMB.UNIFESO.LAN)":
                py.moveTo(573, 697)
                py.click()
                #Selecionar mês dos resultados
                py.moveTo(325, 392)
                py.click()
                py.write(entradaPeriodoMes.get() + entradaPeriodoAno.get())
                py.moveTo(525, 396)
                py.click()
                #Convenio
                if caixaSelecao.get() == "AMIL":
                    py.moveTo(589, 416)
                    py.click()
                elif caixaSelecao.get() == "ASSIM":
                    py.press('down')
                    py.press('down')
                    py.press('enter')
                elif caixaSelecao.get() == "BRADESCO":
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('enter')
                elif caixaSelecao.get() == "CAC":
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('enter')
                elif caixaSelecao.get() == "CEF":
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('enter')
                elif caixaSelecao.get() == "CORPO DE BOMBEIROS DO ESTADO DO RIO DE JANEIRO":
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('enter')
                elif caixaSelecao.get() == "FUSMA":
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('enter')
                elif caixaSelecao.get() == "GOLDEN CROSS":
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('enter')
                elif caixaSelecao.get() == "MEDISERVICE":
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('enter')
                elif caixaSelecao.get() == "PETROBRAS":
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('enter')
                elif caixaSelecao.get() == "POSTAL SAUDE":
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('enter')
                elif caixaSelecao.get() == "SUL AMERICA":
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('enter')
                elif caixaSelecao.get() == "UNIMED":
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('down')
                    py.press('enter')

                py.moveTo(699, 388)
                py.click()
                if caixaUnidade.get() == "HOSPITAL DAS CLÍNICAS DE TERESÓPOLIS":
                    py.moveTo(836, 427)
                    py.click()
                py.moveTo(914, 393)
                py.click()
            else:
                py.alert("A janela de Recebimentos não startou. Encerrando...")
                exit()
        except KeyboardInterrupt as kex:
            print("Programa encerrado pelo usuário via teclado.")
            exit()
        except:
            print("Resolução de tela não cadastrada no sistema")
            exit()        
    else:
        py.alert("Resolução de tela não cadastrada. Encerrando...")
        exit()

    #Gerador de repetições
    conexao = pyodbc.connect(dados_conexao)
    cursor = conexao.cursor()

    im3 = pyautogui.screenshot(f'db/image.png', region=(333,435,28,15))
    time.sleep(1)
    phrase2 = ocr.image_to_string(PIL.Image.open('db/image.png'), lang='por')
    print(phrase2)
    cursor.execute(f"SELECT count(*) FROM RM.SZCONTROLRECEB WHERE DOCUMENTO = '{phrase2}' AND CODCOMPRADOR = 5 and DTRECEBIMENTO BETWEEN '01/{entradaPeriodoMes}/{entradaPeriodoAno.get()}' AND '31/{entradaPeriodoMes}/{entradaPeriodoAno.get()}';")
    tables = int(cursor.fetchval())
    print(f"Foram encontradas {tables} ocorrências.")
    engine.say(f"Foram encontradas {tables} ocorrências.")
    engine.runAndWait()
    contador = 0
    while contador != tables:
        #Variável de Geração de Screenshots
        im = pyautogui.screenshot(f'imagens/image{contador}.png', region=(400,235,54,20))
        time.sleep(1)
        #Variável de processamento de imagens e obtenção de resultado para o sistema 100% testada e funcionando
        phrase = ocr.image_to_string(PIL.Image.open(f'imagens/image{contador}.png'), lang='por')
        print(f"O valor obtido foi de {phrase} reais.")
        if res.width == 1366 and res.height == 768:
            #Clique na seção de Contas
            py.moveTo(610,356)
            py.click()
            #Marcar conta np sistema
            if contador == 0:
                py.moveTo(467,458)
                py.doubleClick()
            elif contador == 1:
                py.moveTo(537,478)
                py.doubleClick()
            elif contador == 2:
                py.moveTo(537,495)
                py.doubleClick()
            elif contador == 3:
                py.moveTo(467,510)
                py.doubleClick()
            elif contador == 4:
                py.moveTo(479,527)
                py.doubleClick()
            elif contador == 5:
                py.moveTo(477,543)
                py.doubleClick()
            elif contador == 6:
                py.moveTo(454, 560)
                py.doubleClick()
            elif contador == 7:
                py.moveTo(479,578) 
                py.doubleClick()
            elif contador == 8:
                py.moveTo(474, 596)
                py.doubleClick()
            elif contador == 9:
                py.moveTo(469,612)
                py.doubleClick()
            elif contador == 10:
                py.moveTo(452,627)
                py.doubleClick()
            elif contador == 11:
                py.moveTo(466,644)
                py.doubleClick()
            elif contador == 12:
                py.moveTo(515, 645)
                py.click()
                py.press('down')
                py.doubleClick()
            elif contador > 12 and contador <= tables:
                py.moveTo(517, 645)
                py.click()
                py.press('down')
                i = 13
                while i == contador:
                    py.press('down')
                    i += 2
                i = i + 1
                py.doubleClick()
            #recebimento
            py.moveTo(479,147)
            py.click()
            #Clicar na data
            py.moveTo(412, 252)
            py.click()
            im1 = pyautogui.screenshot(f'valor/image{contador}.png', region=(337,195,30,14))
            time.sleep(1)
            phrase1 = ocr.image_to_string(PIL.Image.open(f'valor/image{contador}.png'), lang='por')
            #Recortar valor
            py.moveTo(494, 245)
            py.click()
            py.doubleClick()
            py.press('backspace')
            #Colar em Perda
            py.moveTo(493,291)
            py.click()

            py.press('backspace')
            py.write(phrase1)
            #Gravando Processo no sistema
            py.press('tab')
            py.press('enter')
            #Último Processo 
            py.write('1')
            py.press('enter')
            #Saindo
            py.press('enter')

            py.press('enter')
            py.press('enter')
            time.sleep(3)
            title = "Erro"
            window = pygetwindow.getWindowsWithTitle(title)[0]
            if window.title == "Erro (RDESKAMB.UNIFESO.LAN)":
                try:
                    py.press('enter')
                except:
                    print('Tela de conclusão era requerida para continuar!')
                    engine.say('Tela de conclusão era requerida para continuar!')
                    engine.runAndWait()
                    exit()
            #Cancelando
            py.press('tab')
            py.press('tab')
            py.press('enter')
            fim = time.time()
            resultado = (fim - inicio)
            print(f"Duração de execução até aqui foi de {resultado: .2f} segundos. E a ordem desse paciente  é a {contador}.")
            engine.say(f"Duração de execução até aqui foi de {resultado: .2f} segundos. E a ordem desse paciente  é a {contador}.")
            engine.runAndWait()
            contador = contador + 1
        else:
            py.alert("Resoluçao de tela não compatível com o sistema. Finalizando...")
            exit()
    #Alerta de encerramento
    py.alert("A automação foi finalizada. A máquina está liberada pra uso.")
    engine.say("A automação foi finalizada. A máquina está liberada pra uso.")
    engine.runAndWait()
    engine.stop()
    window.mainloop()
    exit()

def recuperarLista():
    if Checkbutton1.get() == 1:
        baixa()
        exit()
    elif Checkbutton2.get() == 1:
        automacao()
        
btn = Button(window, text='Executar automação', command=recuperarLista, width=30, bg="#d3d3d3")
btn.place(x=90, y=140)

window.mainloop()


