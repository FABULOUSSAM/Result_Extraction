import pytesseract
import cv2
import numpy as np
pytesseract.pytesseract.tesseract_cmd=r'C:\Users\Shubham\AppData\Local\Tesseract-OCR\tesseract.exe'

img = cv2.imread('./Crop_Image/crop_9.png')
gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
blurred = cv2.GaussianBlur(gray, (5, 5), 0)
#edged = cv2.Canny(blurred, 50, 200, 255)
custom_config = r'--oem 3 --psm 6'
print(pytesseract.image_to_string(blurred, config=custom_config))