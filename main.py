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
# s = set()
# for i in range(6):
#     print(next(docx_iter).text)
# exit()

bullet_buffer = []
for p in docx_iter:
    p_obj = Decider.get_style(p)
    if p_obj == 0:
        continue
    if isinstance(p_obj, BulletListText):
        bullet_buffer.append(p_obj)
    else:
        if bullet_buffer:
            BulletListText.compile_list(bullet_buffer, new_document)
            bullet_buffer = []
        p_obj.add_paragraph(new_document)




new_document.save('output/new_c.docx')
