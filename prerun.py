from docx import Document
from styles import Styles

old_document = Document("c2.docx")


styles = Styles(old_document)


old_document.save('input/work_c.docx')
