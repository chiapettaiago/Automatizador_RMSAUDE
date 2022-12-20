import pyautogui as py
import pyautogui
import PIL
import pytesseract as ocr
from PIL import Image
import time

#Caminho até a automação
py.alert("Atenção! A automação está sendo iniciada.")
py.PAUSE = 1.5
'''
#Dar dois cliques na janela do aplicativo saúde
py.moveTo(107,334)
py.doubleClick()
#Inserir a senha
time.sleep(3)
#py.write(senha.senha_saude)
#Aguardar o app carregar

#Acessar o App
py.write(senha.senha_login_saude)
py.press('enter')
time.sleep(40)

'''
#Acessar onde ficam os dados a serem obtidos
py.moveTo(784, 26)
py.click()
py.moveTo(782, 142)
py.click()
time.sleep(3)
py.moveTo(573, 697)
py.click()
#Selecionar mês dos resultados
py.moveTo(325, 392)
py.click()
py.write("01")
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

#Variável de Geração de Screenshots
im = pyautogui.screenshot('image.jpg', region=(400,235,54,20))
time.sleep(10)
#Variável de processamento de imagens e obtenção de resultado para o sistema 100% testada e funcionando
phrase = ocr.image_to_string(Image.open('image.jpg'), lang='por')
print(phrase)
py.alert(phrase)

py.press('down')

#Alerta de encerramento 
py.alert("A automação foi finalizada. A máquina está liberada pra uso.")