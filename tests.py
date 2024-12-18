
import docx
from docx import Document
from styles import Styles
from text import Decider
from text import BulletListText

old_document = Document("c.docx")
new_document = Document("empty.docx")
new_document.save('output/new_c.docx')
new_document = Document("output/new_c.docx")
new_document.save('output/new_c.docx')


docx_iter = old_document.iter_inner_content()
f = next(docx_iter)
print(f.style.base_style)
