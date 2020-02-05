from PIL import Image
from pptx import Presentation
from pytesseract import image_to_string
import pytesseract
import docx
import sys
import os
import comtypes.client
import numpy as np
import cv2

#CONTRASTING THE IMAGE

# read
img = cv2.imread('image.jpg', cv2.IMREAD_GRAYSCALE)

# increase contrast
pxmin = np.min(img)
pxmax = np.max(img)
imgContrast = (img - pxmin) / (pxmax - pxmin) * 255

# increase line width
kernel = np.ones((3, 3), np.uint8)
imgMorph = cv2.erode(imgContrast, kernel, iterations = 1)

# write
#cv2.imwrite('out.png', imgMorph)



#CONVERTING TO STRING

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
p=image_to_string( imgMorph,lang='eng')
print(p)
