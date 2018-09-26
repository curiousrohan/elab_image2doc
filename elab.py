from docx import Document
from docx.shared import Inches
import glob

document = Document()

sections = document.sections
for section in sections:
	section.top_margin = Inches(1.9)
	section.bottom_margin = Inches(0.7)
	section.left_margin = Inches(1)
	section.right_margin = Inches(1)

#/create array containing names of all .png files in current directory
pictures = glob.glob('./*.png')

for picture in pictures:
	#parse file name (remove ./)
	pic = picture[2:]
	try:
		document.add_picture(pic,width=Inches(7),height=Inches(7))
	except:
		print('Picture Name: '+pic+' could not be processed')

filename=input("Enter the final filename...\n")
filename=filename+'.docx'
document.save(filename)
