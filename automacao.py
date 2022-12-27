import pyautogui as py
import pyautogui
import pytesseract as ocr
from PIL import Image
import time
import pygetwindow
import pyodbc

dados_conexao = ('DRIVER={Devart ODBC Driver for Oracle};Host=10.1.0.137;Port=1521;SID=DBFESOCONS;UID=GTIC_078463;Password=gtic_sup23_078463;Direct=True')
#Caminho até a automação
inicio = time.time()
py.alert("Atenção! A automação está sendo iniciada.")
py.PAUSE = 1.2
res = py.size()
if res.width == 1366 and res.height == 768:
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
        mes = py.write("012021")
        py.moveTo(525, 396)
        py.click()
        #Convenio
        py.moveTo(589, 416)
        py.click()
        py.moveTo(699, 388)
        py.click()
        py.moveTo(836, 427)
        py.click()
        py.moveTo(914, 393)
        py.click()
    else:
        py.alert("A janela de Recebimentos não startou. Encerrando...")
        exit()
else:
    py.alert("Resolução de tela não cadastrada. Encerrando...")
    exit()

#Gerador de repetições
conexao = pyodbc.connect(dados_conexao)
cursor = conexao.cursor()

cursor.execute("SELECT count(*) FROM SZFATURACONVENIO INNER JOIN SZCADGERAL ON SZFATURACONVENIO.CODCOMPRADOR = SZCADGERAL.CODGERAL WHERE DATAGERACAO BETWEEN '01-jan-2021' AND '30-jan-2021' AND RAZAOSOCIALPRESTADOR = 'HOSPITAL DAS CLINICAS DE TERESOPOLIS' AND CODGERAL BETWEEN 1 AND 204;" )
tables = int(cursor.fetchval())
contador = 0
while contador != tables:
    #Variável de Geração de Screenshots
    im = pyautogui.screenshot(f'imagens/image{contador}.png', region=(400,235,54,20))
    time.sleep(1)
    #Variável de processamento de imagens e obtenção de resultado para o sistema 100% testada e funcionando
    phrase = ocr.image_to_string(Image.open(f'imagens/image{contador}.png'), lang='por')
    print(phrase)
    py.alert(phrase)
    if res.width == 1366 and res.height == 768:
        #Clique na seção de Contas
        py.moveTo(610,356)
        py.click()
        #Marcar conta np sistema
        py.moveTo(467,458)
        py.doubleClick()
        #recebimento
        py.moveTo(479,147)
        py.click()
        #Clicar na data
        py.moveTo(412, 252)
        py.click()
        im1 = pyautogui.screenshot(f'valor/image{contador}.png', region=(337,195,30,14))
        time.sleep(1)
        phrase1 = ocr.image_to_string(Image.open(f'valor/image{contador}.png'), lang='por')
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
        time.sleep(1)
        title = "Erro"
        window = pygetwindow.getWindowsWithTitle(title)[0]
        if window.title == "Erro (RDESKAMB.UNIFESO.LAN)":
            py.press('enter')
        else:

            py.alert('Mensagem de erro não apareceu')
            exit()
        #Cancelando
        py.press('tab')
        py.press('tab')
        py.press('enter')
        fim = time.time()
        resultado = (fim - inicio)
        print(resultado)
        contador = contador + 1
    else:
        py.alert("Resoluçao de tela não compatível com o sistema. Finalizando...")
        exit()
#Alerta de encerramento     
py.alert("A automação foi finalizada. A máquina está liberada pra uso.")
exit()