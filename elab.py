from docx import Document
from docx.shared import Inches

document = Document()

sections = document.sections
for section in sections:
	section.top_margin = Inches(1.9)
	section.bottom_margin = Inches(0.7)
	section.left_margin = Inches(1)
	section.right_margin = Inches(1)

for i in range(0,68):
	if (i==0):
		picture='report.png'
	else:
		picture='report ('+str(i)+').png'
	document.add_picture(picture,width=Inches(7),height=Inches(7))

filename=input("Enter the final filename...\n")
filename=filename+'.docx'
document.save(filename)