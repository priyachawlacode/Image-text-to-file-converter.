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
#print(p)
ch=int(input("enter your choice\n 1. conversion to word \n 2. conversion to pdf \n 3. conversion to ppt"))

"""
#TRAINING THE MODEL

class DataProvider():
	"this class creates machine-written text for a word list. TODO: change getNext() to return your samples."

	def __init__(self, wordList):
		self.wordList = wordList
		self.idx = 0

	def hasNext(self):
		"are there still samples to process?"
		return self.idx < len(self.wordList)

	def getNext(self):
		"TODO: return a sample from your data as a tuple containing the text and the image"
		img = np.ones((32, 128), np.uint8)*255
		word = self.wordList[self.idx]
		self.idx += 1
		cv2.putText(img, word, (2,20), cv2.FONT_HERSHEY_SIMPLEX, 0.4, (0), 1, cv2.LINE_AA)
		return (word, img)


def createIAMCompatibleDataset(dataProvider):
	"this function converts the passed dataset to an IAM compatible dataset"

	# create files and directories
	f = open('words.txt', 'w+')
	if not os.path.exists('sub'):
		os.makedirs('sub')
	if not os.path.exists('sub/sub-sub'):
		os.makedirs('sub/sub-sub')

	# go through data and convert it to IAM format
	ctr = 0
	while dataProvider.hasNext():
		sample = dataProvider.getNext()
		
		# write img
		cv2.imwrite('sub/sub-sub/sub-sub-%d.png'%ctr, sample[1])
		
		# write filename, dummy-values and text
		line = 'sub-sub-%d'%ctr + ' X X X X X X X ' + sample[0] + '\n'
		f.write(line)
		
		ctr += 1
		
		
if __name__ == '__main__':
	words = ['plugin', 'exclusively', 'captured', 'on', 'film', 'create', 'smartphones']
	dataProvider = DataProvider(words)
	createIAMCompatibleDataset(dataProvider)



"""





#CONVERSION TO WORD

if (ch==1):
     # create an instance of a word document 
     doc = docx.Document() 
  
     # add a heading of level 0 (largest heading) 
     doc.add_heading('Heading for the document', 0) 
  
     # add a paragraph and store  
     doc_para = doc.add_paragraph(p)
     # pictures can also be added to our word document 
     # width is optional 
     doc.add_picture('test.jpg') 

     # now save the document to a location 
     doc.save('hello')




#CONVERSION TO PDF

if (ch==2):
 
     """wdFormatPDF = 17

     in_file = os.path.abspath(sys.argv[1])
     out_file = os.path.abspath(sys.argv[2])

     word = comtypes.client.CreateObject('hello')
     doc = word.Documents.Open(in_file)
     doc.SaveAs(out_file, FileFormat=wdFormatPDF)
     doc.close()
     word.Quit()"""

# CONVERSION TO PPT

if (ch==3):
    
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    
    title.text = "Hello"
    subtitle.text = p

    prs.save('test.pptx')

    
    """
    f=open('temp.pptx')
    prs = Presentation(f)
    # Use the output from analyze_ppt to understand which layouts and placeholders
    # to use
    # Create a title slide first
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[3]
    title.text = "Quarterly Report"
    subtitle.text = "Generated on {:%m-%d-%Y}"
    subtitle.text = p
    prs.save('test.pptx')
    f.close()
    """


