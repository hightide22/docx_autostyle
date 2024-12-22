from docx.enum.style import WD_STYLE_TYPE
import docx
from docx import Document
from styles import Styles
from text import Decider
from styles import create_main, create_header1

old_document = Document("c.docx")
new_document = Document("empty.docx")
new_document.save('output/new_c.docx')
new_document = Document("output/new_c.docx")
new_document.save('output/new_c.docx')

styles = Styles(old_document)
docx_iter = old_document.iter_inner_content()
# s = set()
# for i in range(6):
#     print(next(docx_iter).text)
# exit()

bullet_buffer = []
r_style = create_header1(old_document)
for p in docx_iter:
   r_style = Decider.get_style(p, styles)

   if not r_style:
      continue
   p.style = r_style



old_document.save('output/new_old_c.docx')
new_document.save('output/new_c.docx')
