from docx import Document
from styles import Styles, Decider
from text import Control

old_document = Document("input/work_c.docx")
changes = Document("input/work_c.docx")

styles = Styles(old_document)

custom_style_names = True
if custom_style_names:
   Decider.custom_names(styles)

docx_iter = old_document.iter_inner_content()
changes_iter = changes.iter_inner_content()

for p in docx_iter:
   Control.handle_paragraph(p, styles)
   Control.get_difference(next(changes_iter), p)





old_document.save("output/work_c.docx")
changes.save("output/diffs.docx")
print("!")