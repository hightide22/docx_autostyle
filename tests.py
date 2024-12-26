
import docx
from docx import Document
import docx.document
# from styles import Styles
# old_document = Document("c1.docx")
eq = Document("eq.docx")
# def create_id_map(doc: Document) -> dict[int, str]:
#     numbering = doc.part.numbering_part
#     list_type_dict = {}
#     abstractNumId_numId = {}
#
#     if numbering is not None:
#         for n in numbering.element.findall('.//w:num', namespaces=doc.part.element.nsmap):
#             numId = n.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numId')
#             abstractNumId = n.find('.//w:abstractNumId', namespaces=doc.part.element.nsmap).get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
#             abstractNumId_numId[int(abstractNumId)] = int(numId)
#         for num in numbering.element.findall('.//w:abstractNum', namespaces=doc.part.element.nsmap):
#             for lvl in num.findall('.//w:lvl', namespaces=doc.part.element.nsmap):
#                 if int(lvl.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ilvl')) != 0:
#                     continue
#                 for fmt in lvl.findall('.//w:numFmt', namespaces=doc.part.element.nsmap):
#                     # print(num.items())
#                     # num.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}abstractNumId', "100")
#                     # print(num.items())
#                     abstractNumId = int(num.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}abstractNumId'))
#                     list_type_dict[abstractNumId_numId[abstractNumId]] = fmt.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
#         # print(abstractNumId_numId)
#         return list_type_dict
#
#
#
#
# def is_bulleted_paragraph(paragraph, d):
#     _el = paragraph._element
#     p_xml = _el.find('.//w:pPr', namespaces=paragraph.part.element.nsmap)
#     a = p_xml.find('.//w:numPr', namespaces=paragraph.part.element.nsmap)
#     if a is None:
#         if "маркир" in paragraph.style.name.lower() or "bullet" in paragraph.style.name.lower():
#             return "Bullet"
#         if paragraph.style.base_style:
#             if "маркир" in paragraph.style.base_style.name.lower() or "bullet" in paragraph.style.base_style.name.lower():
#                 return "Bullet"
#         return 0
#     b = a.find('.//w:numId', namespaces=paragraph.part.element.nsmap)
#     id = int(b.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val'))
#     b.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', "10")
#     print(d[id])
#     return a


    # return numPr is not None
# d = create_id_map(eq)
i = eq.iter_inner_content()
a = None
for p in i:
    print(p.style.paragraph_format.alignment, p.paragraph_format.alignment)

eq.save("output/eq2.docx")



