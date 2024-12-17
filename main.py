import docx
from docx import Document
from styles import Styles


old_document = Document("c.docx")
new_document = Document("empty.docx")
new_document.save('new_c.docx')
new_document = Document("new_c.docx")

docx_iter = old_document.iter_inner_content()
for i in range(5):

    print(next(docx_iter).style.type)
    print(next(docx_iter).text, end="\n\n\n\n\n")













new_document.save('output/new_c.docx')
