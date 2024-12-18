from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.document import Document
from styles import Styles
from docx.text.paragraph import Paragraph
from string import punctuation
from docx.shared import Mm

class Decider:
    @staticmethod
    def get_style(p: Paragraph):
        """
        :param p: paragraph
        :return:
        """
        if p.style.type == WD_STYLE_TYPE.PARAGRAPH:
            if "список" in p.style.name.lower() or "list" in p.style.name.lower():
                if "bullet" in p.style.name.lower() or "марк" in p.style.name.lower():
                    return BulletListText(p.text)
                else:
                    return NumsListText(p.text)
            if p.style.base_style is None:
                print(p.text, p.style.name)
                return 0

            if p.style.base_style.name == "Heading 1":
                return HeaderBigText(p.text)
            elif p.style.base_style.name == "Heading 2":
                return HeaderSmallText(p.text)
            else:
                return ParagraphText(p.text)

        else:
            return 0



class ParagraphText:
    def __init__(self, text: str):
        self.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        self.text = text
        self.first_line_indent = Mm(12.5)
        self.style = Styles().main
        self.list = False

    def add_paragraph(self, file: Document) -> Paragraph:
        self.handle_text()
        par = file.add_paragraph(text=self.text, style=self.style)
        # par.paragraph_format.first_line_indent = Mm(12.5)
        return par


    def replace_bad_symbols(self):
        self.text = self.text.replace("«", '"')
        self.text = self.text.replace("»", '"')
        self.text = self.text.replace("“", '"')
        self.text = self.text.replace("”", '"')

        self.text = self.text.replace(" - ", ' — ')

        self.text = self.text.replace("  ", " ") # 2 spaces
        self.text = self.text.replace(".[", ". [")  #

    def handle_text(self):
        self.replace_bad_symbols()

    def get_text(self):
        print("!!!ParagraphText!!!")
        return self.text


class HeaderBigText(ParagraphText):
    def __init__(self, text: str):
        super().__init__(text)
        self.style = Styles().header1


class HeaderSmallText(ParagraphText):
    def __init__(self, text: str):
        super().__init__(text)
        self.style = Styles().header2





class ListText(ParagraphText):
    def __init__(self, text: str):
        super().__init__(text)
        self.list = True

    def add_paragraph(self, file: Document) -> Paragraph:
        self.handle_text()
        par = file.add_paragraph(text=self.text, style=self.style)
        par.paragraph_format.first_line_indent = Mm(-10)
        par.paragraph_format.left_indent = Mm(22.5)
        return par


class BulletListText(ListText):
    def __init__(self, text: str):
        super().__init__(text)
        self.style = Styles().list_bullet
        self.is_last = False
        self.is_lower = False

    def correct_end(self):
        if self.is_lower:
            self.text = self.text[0].lower() + self.text[1:]
        else:
            self.text = self.text[0].upper() + self.text[1:]

        if self.text[-1] in punctuation:
            self.text = self.text[:-1] + (";" if self.is_lower and not self.is_last else ".")
        else:
            self.text = self.text + (";" if self.is_lower and not self.is_last else ".")

    def handle_text(self):
        self.correct_end()
        self.replace_bad_symbols()

    @staticmethod
    def compile_list(l: list["BulletListText"], d: Document):
        is_lower = l[0].text[0].islower()
        l[-1].is_last = True
        for p in l:
            p.is_lower = is_lower
            p.add_paragraph(d)




class NumsListText(ListText):
    def __init__(self, text=""):
        super().__init__(text)
        self.style = Styles().list_num


class MainText(ParagraphText):
    def __init__(self, text=""):
        super().__init__(text)
        self.text = text
        self.style = Styles().main
        self.list = False


    def is_P_too_long(self, p: str) -> bool:
        p.count()