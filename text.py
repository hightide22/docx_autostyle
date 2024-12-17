from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.document import Document
from styles import Styles
from docx.text.paragraph import Paragraph

class ParagraphText:
    def __init__(self, text=""):
        self.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        self.text = text
        self.first_line_indent =
        self.style = Styles().main

    def add_paragraph(self, file: Document):
        file.add_paragraph(text=self.text, style=self.style)

    def replace_bad_symbols(self):
        self.text = self.text.replace("«", '"')
        self.text = self.text.replace("»", '"')
        self.text = self.text.replace("“", '"')
        self.text = self.text.replace("”", '"')

        self.text = self.text.replace(" - ", ' — ')

        self.text = self.text.replace("  ", " ") # 2 spaces


class Decider:
    def get_style(self, p: Paragraph):
        """
        :param p: paragraph
        :return:
        """
        if p.style.name.lower().startswith("List Bullet"):
            return
        if p.style.name.lower().startswith("List Number"):
            return
        style = p.style
        size = style.font.size
        p.



class MainText(ParagraphText):
    def __init__(self, text=""):
        self.text = text
        self.style = Styles().main


    def is_P_too_long(self, p: str) -> bool:
        p.count()