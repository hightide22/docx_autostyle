from docx import Document
import docx.document


"""
Перенос всех стилей из методички в файл
"""

def add_styles(target_document: docx.document.Document, styles_document: docx.document.Document):
    s = styles_document.part._styles_part.element
    for style in s.findall('.//w:style', namespaces=styles_document.part.element.nsmap):
        el = style.find(".//w:name", namespaces=styles_document.part.element.nsmap)
        if "ПМ" in el.val:
            target_document.part._styles_part.element.append(style)

def prerun():
    styles = Document("PM/PM.docx")
    old_document = Document("PM/nf.docx")
    add_styles(old_document, styles)
    old_document.save("PM/result.docx")

prerun()