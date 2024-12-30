from docx import Document
import docx.document
from styles import Styles, Decider


def add_styles(target_document: docx.document.Document, styles_document: docx.document.Document):
    s = styles_document.part._styles_part.element
    for style in s.findall('.//w:style', namespaces=styles_document.part.element.nsmap):
        el = style.find(".//w:name", namespaces=styles_document.part.element.nsmap)
        if el.val in ("main", "1list bullet", "1list num", "header1", "header2", "picture", "source_header", "eq"):
            target_document.part._styles_part.element.append(style)


def prerun(old_document: docx.document.Document):

    styles = Document("input/def_styles.docx")
    add_styles(old_document, styles)
    styles = Styles(old_document)
    custom_style_names = True
    if custom_style_names:
        Decider.custom_names(styles)


if __name__ == "__main__":
    prerun()
    print("!")
