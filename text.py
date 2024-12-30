from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.enum.style import WD_STYLE_TYPE
from collections import defaultdict
from styles import Styles, Decider
from docx.text.paragraph import Paragraph
from string import punctuation
from docx.shared import Mm, RGBColor
import regex


class ParagraphText:
    @staticmethod
    def handle_text(p: Paragraph):
        """
        p.text = ... Удалит все формулы из параграфа
        """
        for run in p.runs:
            if len(run.text) > 1:
                run.text = ParagraphText._handle_text(run.text)
        while p.runs and p.runs[-1].text.endswith(" "):
            p.runs[-1].text = p.runs[-1].text[:-1]
        ParagraphText._handle_eq(p)

    @staticmethod
    def _handle_eq(p: Paragraph):
        if "[eq]" in p.text or u"\u200b" in p.text:
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                if r.text in ("[", "eq", "]", "[eq]"):
                    r.text = u"\u200b"
                r.text = r.text.replace("[eq]", u"\u200b")

    @staticmethod
    def _handle_text(text: str) -> str:
        text = ParagraphText._replace_bad_symbols(text)
        # text = ParagraphText._handle_quotes(text) # Не работает с run
        text = ParagraphText._replace_bad_spaces(text)
        return text

    @staticmethod
    def _replace_bad_symbols(text: str) -> str:
        # text = text.replace("«", '"').replace("»", '"')
        text = text.replace("“", '"').replace("”", '"')

        text = text.replace(" - ", ' — ')

        text = text.replace(u"\u00A0", " ")  # non-breaking spaces
        return text

    @staticmethod
    def _handle_quotes(text: str) -> str:
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
    def _replace_bad_spaces(text: str) -> str:
        text = text.replace("  ", " ")
        text = text.replace(".[", ". [")
        text = text.replace("( ", "(").replace(" )", ")")

        text = text.replace(" .", ".").replace(" ,", ",").replace(" :", ":")
        text = text.replace('« ', '«').replace(' »', "»")
        return text


class BulletListText(ParagraphText):
    @staticmethod
    def handle_text(p: Paragraph, last=False):
        ParagraphText.handle_text(p)
        if p.runs:
            if p.runs[0].text:
                p.runs[0].text = BulletListText._capitalize(p.runs[0].text)
            if p.runs[-1].text:
                p.runs[-1].text = BulletListText._handle_list(p.runs[-1].text, last)

    @staticmethod
    def _handle_text(text: str, last=False) -> str:
        text = BulletListText._handle_list(text, last)
        return text

    @staticmethod
    def _capitalize(text: str) -> str:
        text = text[0].lower() + text[1:]  # Текст в маркированном списке начинается с маленькой (строчной) буквы
        return text

    @staticmethod
    def _handle_list(text: str, last: bool) -> str:
        if not text:
            return text
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
    def handle_text(p: Paragraph, last=False):
        BulletListText.handle_text(p, last)
        if p.runs:
            if p.runs[0].text:
                p.runs[0].text = NumListText._capitalize(p.runs[0].text)
            if p.runs[-1].text:
                p.runs[-1].text = NumListText._handle_list(p.runs[-1].text, False)

    @staticmethod
    def _handle_text(text: str, last=False) -> str:
        text = NumListText._handle_list(text, last)
        return text

    @staticmethod
    def _capitalize(text):
        return text[0].upper() + text[1:]  # Текст в нумерованном списке должен начинаться с прописной буквы

    @staticmethod
    def _handle_list(text: str, last) -> str:
        if text[-1] == "]":
            return text
        if text[-1] in punctuation:
            text = text[:-1] + "."
        else:
            text = text + "."
        return text


class PictureText(ParagraphText):
    @staticmethod
    def handle_text(p: Paragraph):
        ParagraphText.handle_text(p)
        for run in p.runs:
            if len(run.text) > 1:
                run.text = PictureText._handle_text(run.text)

    @staticmethod
    def _handle_text(text: str) -> str:
        text = PictureText._handle_picture(text)
        return text

    @staticmethod
    def _handle_picture(text: str) -> str:
        text = text.replace("рисунок", "Рисунок")
        if "-" in text and "—" not in text:
            text = text.replace("-",  "—", 1)
        return text


class Control:
    _bullet_list_buffer: Paragraph = None
    _style_diffs = defaultdict(list)

    @staticmethod
    def handle_paragraph(p: Paragraph, style_obj: Styles):
        style_dict = {
            "main": ParagraphText,
            "header1": ParagraphText,
            "header2": ParagraphText,
            "source_header": ParagraphText,
            "picture": PictureText,
            "1list bullet": BulletListText,
            "1list num": NumListText
        }
        style = Decider.get_style(p, style_obj)
        if Decider._is_eq(p) and len(p.text) < 8:
            if not p.runs:
                p.add_run(" ")
            p.style = style_obj.eq
            return
        if style == 0:
            return
        p.style = style
        Control.normalize(p)
        if style.name not in style_dict.keys():
            style_name = "1list num"
        else:
            style_name = style.name
        if style.name == "1list bullet":
            Control._bullet_list_buffer = p
            style_dict[style.name].handle_text(p)
        else:
            if Control._bullet_list_buffer:
                style_dict["1list bullet"].handle_text(Control._bullet_list_buffer, True)
                Control._bullet_list_buffer = None
            style_dict[style_name].handle_text(p)

    @staticmethod
    def normalize(p: Paragraph):
        if not p.runs:
            return
        p.style.font.color.rgb = RGBColor(0, 0, 0)
        pf = p.paragraph_format
        spf = p.style.paragraph_format

        if pf.space_before != spf.space_before:
            pf.space_before = spf.space_before
        if pf.space_after != spf.space_after:
            pf.space_after = spf.space_after
        if pf.left_indent != spf.left_indent:
            pf.left_indent = spf.left_indent
        if pf.alignment != spf.alignment:
            pf.alignment = spf.alignment
        if pf.first_line_indent != spf.first_line_indent:
            if not p.text.lower().startswith("где"):
                pf.first_line_indent = spf.first_line_indent
        if pf.line_spacing_rule != spf.line_spacing_rule:
            pf.line_spacing_rule = spf.line_spacing_rule

        if p.text.lower().startswith("где") and "—" in p.text and len(p.text) < 100:
            pf.first_line_indent = Mm(0)

    @staticmethod
    def get_difference(old_p: Paragraph, new_p: Paragraph):
        if old_p.style.type != WD_STYLE_TYPE.PARAGRAPH:
            return
        if Control.is_style_different(old_p, new_p):
            for r in old_p.runs:
                r.font.highlight_color = WD_COLOR_INDEX.YELLOW
        for old_r, new_r in zip(old_p.runs, new_p.runs):
            if old_r.text != new_r.text and old_r.text[:-1] != new_r.text:
                old_r.font.color.rgb = RGBColor(255, 0, 0)
                old_r.text = "{" + f"old:[{old_r.text}] new:[{new_r.text}]" + "}"

    @staticmethod
    def is_style_different(old_p: Paragraph, new_p: Paragraph) -> bool:
        if old_p.style.name == new_p.style.name:
            return False
        if not old_p.runs:
            return False
        x = old_p.paragraph_format
        old_p_params = {"left_indent": x.left_indent, "first_line_indent": x.first_line_indent, "line_spacing_rule": x.line_spacing_rule, "space_before": x.space_before, "space_after": x.space_after, "alignment": x.alignment}
        x = new_p.style.paragraph_format
        new_p_params = {"left_indent": x.left_indent, "first_line_indent": x.first_line_indent, "line_spacing_rule": x.line_spacing_rule, "space_before": x.space_before, "space_after": x.space_after, "alignment": x.alignment}
        x = old_p.style.paragraph_format
        old_p_style_params = {"left_indent": x.left_indent, "first_line_indent": x.first_line_indent, "line_spacing_rule": x.line_spacing_rule, "space_before": x.space_before, "space_after": x.space_after, "alignment": x.alignment}

        for key in old_p_style_params:
            old = old_p_params[key]
            new = new_p_params[key]
            old_style = old_p_style_params[key]
            if old is not None or old_style is not None:
                if old:
                    if old != new:
                        Control._style_diffs[old_p.text[:10]].append((key, "old"))
                        # if new_p.style.name.lower() == "main":
                        #     print((key, "old"), old, new, old_style, new_p.text[:10])
                        return True
                if old_style:
                    if old_style != new:
                        Control._style_diffs[old_p.text[:10]].append((key, "old_style"))
                        # if new_p.style.name.lower() == "main":
                        #     print((key, "old_style"), old, new, old_style, new_p.text[:10])
                        return True
            else:
                if old != new and (old or new):
                    # if new_p.style.name.lower() == "main":
                    #     print((key, "old_style2"), old, new, old_style, new_p.text[:10])
                    return True
        if old_p.style.font.color.rgb != RGBColor(0, 0, 0) and old_p.style.font.color.rgb:
            Control._style_diffs[old_p.text[:10]].append(("color", 1))
            return True
        return False
