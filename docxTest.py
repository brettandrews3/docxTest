# This is the sample program from the python-docx documentation.
# I plan on copying the code, commenting throughout, and iterating to make it my own.

from docx import Document
from docx.shared import Inches

document = Document()

#This line sets the heading for the document. 0 = largest heading?
document.add_heading('Document Title', 0)

p = document.add_paragraph('A plain paragraph that\'s got some ')
p.add_run('bold').bold = True
p.add_run(' and some ')
p.add_run('italics.') = True