
import docx
from docx import Document
# from styles import Styles

old_document = Document("c.docx")
old_document = Document("new.docx")
new_document = Document("empty.docx")
new_document.save('output/new_c.docx')
new_document = Document("output/new_c.docx")
new_document.save('output/new_c.docx')

# s = Styles(old_document)
print([x.name for x in old_document.styles])
docx_iter = old_document.iter_inner_content()
f = next(docx_iter)
# print(f.style.base_style)
