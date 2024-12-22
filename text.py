from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.document import Document
from styles import Styles
from docx.text.paragraph import Paragraph, ParagraphStyle
from string import punctuation
from docx.shared import Mm, RGBColor
import regex
print("не забыть про ссылки в конце списка")


class Decider:
    @staticmethod
    def get_style(p: Paragraph, style: Styles):
        if p.style.type == WD_STYLE_TYPE.PARAGRAPH:
            if "список" in p.style.name.lower() or "list" in p.style.name.lower():
                if "bullet" in p.style.name.lower() or "марк" in p.style.name.lower():
                    return style.list_bullet
                else:
                    return style.get_numlist_style(p)


            if "рисунок" in p.text.lower():
                return style.pictures

            if p.style.base_style:
                if p.style.base_style.name == "Heading 1" or p.style.name == "Heading 1":
                    return style.header1
                elif p.style.base_style.name == "Heading 2" or p.style.name == "Heading 2":
                    return style.header2

            if "часть" in p.text.lower() and len(p.text.split()) <= 3:
                return style.source_header

            return style.main

        else:
            return 0

    @staticmethod
    def normalizer(p: Paragraph):
        if not p.runs:
            return
        p.style.font.color.rgb = RGBColor(0, 0, 0)
        pf = p.paragraph_format
        spf = p.style.paragraph_format


        if pf.left_indent != spf.left_indent:
            pf.left_indent = spf.left_indent
        if pf.alignment != spf.alignment:
            pf.alignment = spf.alignment
        if pf.first_line_indent != spf.first_line_indent:
            if not p.text.lower().startswith("где"):
                pf.first_line_indent = spf.first_line_indent
        if pf.line_spacing_rule != spf.line_spacing_rule:
            pf.line_spacing_rule = spf.line_spacing_rule

        if p.text.lower().startswith("где") and "—" in p.text:
            pf.first_line_indent = Mm(0)





        # if pf.left_indent and pf.left_indent != spf.left_indent:
        #     pf.left_indent = spf.left_indent
        # if pf.left_indent and pf.left_indent != spf.left_indent:
        #     pf.left_indent = spf.left_indent
        # if pf.left_indent and pf.left_indent != spf.left_indent:
        #     pf.left_indent = spf.left_indent
        # if pf.left_indent and pf.left_indent != spf.left_indent:
        #     pf.left_indent = spf.left_indent

class ParagraphText:
    @staticmethod
    def handle_text(text: str) -> str:
        text = ParagraphText.replace_bad_symbols(text)
        text = ParagraphText.handle_quotes(text)
        text = ParagraphText.replace_bad_spaces(text)
        return text

    @staticmethod
    def replace_bad_symbols(text: str) -> str:
        # text = text.replace("«", '"').replace("»", '"')
        text = text.replace("“", '"').replace("”", '"')

        text = text.replace(" - ", ' — ')
        # «–» минус
        return text

    @staticmethod
    def handle_quotes(text: str) -> str:
        if "«" in text and "»" in text:
            mask = r'«.*?»'
            quotes = regex.findall(mask, text)
            for q in quotes:
                for ch in q[1:-1]:
                    if ord(ch) < 700:
                        text = text.replace(q, q.replace("«", '"').replace("»", '"'))
                        break
        if '"' in text:
            mask = r'".*?"'
            quotes = regex.findall(mask, text)
            for q in quotes:
                for ch in q[1:-1]:
                    if ord(ch) > 700:
                        text = text.replace(q, f"«{q[1:-1]}»")
                        break
        return text


    @staticmethod
    def replace_bad_spaces(text: str) -> str:
        text = text.replace("  ", " ") # 2 spaces
        text = text.replace(".[", ". [") #
        text = text.replace("( ", "(").replace(" )", ")")

        text = text.replace(" .", ".").replace(" ,", ",").replace(" :", ":")
        text = text.replace('« ', '«').replace(' »', "»")

        return text


class BulletListText(ParagraphText):
    @staticmethod
    def handle_text(text: str, last=False) -> str:
        text = super().handle_text(text)
        text = BulletListText.handle_list(text, last)
        return text

    @staticmethod
    def handle_list(text: str, last: bool) -> str:
        text = text[0].lower() + text[1:] # Текст в маркированном списке начинается с маленькой (строчной) буквы
        if text[-1] == "]":
            return text
        if last:
            if text[-1] in punctuation:
                text = text[:-1] + "."
            else:
                text = text + "."
        else:
            if text[-1] in punctuation:
                text = text[:-1] + ";"
            else:
                text = text + ";"
        return text

class NumListText(BulletListText):
    @staticmethod
    def handle_list(text: str, last) -> str:
        text = text[0].upper() + text[1:]  # Текст в нумерованном списке должен начинаться с прописной буквы
        if text[-1] == "]":
            return text
        if text[-1] in punctuation:
            text = text[:-1] + "."
        else:
            text = text + "."

class PictureText(ParagraphText):
    @staticmethod
    def handle_text(text: str) -> str:
        text = super().handle_text(text)
        return text

    @staticmethod
    def handle_picture(text: str) -> str:
        text = text.replace("рисунок", "Рисунок")
        if "-" in text and "—" not in text:
            text = text.replace("-",  "—", 1)
