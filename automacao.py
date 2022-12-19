import pyautogui
import PIL
import pytesseract as ocr
from PIL import Image
import time

#py.alert("Atenção! A automação está sendo iniciada.")
#py.PAUSE = 1.5
#py.moveTo(107,334)
#py.doubleClick()

#im = pyautogui.screenshot('image.jpg')
#time.sleep(3)
phrase = ocr.image_to_string(Image.open('image.jpg'), lang='por')
print(phrase)