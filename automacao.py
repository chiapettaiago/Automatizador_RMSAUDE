import pyautogui as py
import pyautogui
import PIL
import pytesseract as ocr
from PIL import Image
import time
import senha

#Caminho até a automação
py.alert("Atenção! A automação está sendo iniciada.")
py.PAUSE = 1.5
#Dar dois cliques na janela do aplicativo saúde
py.moveTo(107,334)
py.doubleClick()
#Inserir a senha
py.write(senha.senha_saude)
#Aguardar o app carregar

#Acessar o App
py.write(senha.senha_login_saude)
py.press('enter')



im = pyautogui.screenshot('image.jpg', region=(400,235,50,20))
time.sleep(10)
phrase = ocr.image_to_string(Image.open('image.jpg'), lang='por')
print(phrase)