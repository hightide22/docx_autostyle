from docx.document import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Pt, Mm, RGBColor
from docx.styles.style import ParagraphStyle
from docx.text.paragraph import Paragraph




















class Styles:
    def _create_numlist_style(self, style_name: str):
        if f"1list num {style_name}" in [x.name for x in self.nf.styles]:
            style = self.nf.styles[f"1list num {style_name}"]
        else:
            style = self.nf.styles.add_style(f"1list num {style_name}", WD_STYLE_TYPE.PARAGRAPH)
        style.quick_style = True
        style.base_style = self.nf.styles[style_name]
        style.name = f"1list num {style_name}"
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        style.paragraph_format.first_line_indent = Mm(-10)
        style.paragraph_format.left_indent = Mm(22.5)
        style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE

        style.font.name = "Times New Roman"
        style.font.size = Pt(14)
        style.font.bold = False
        style.font.color.rgb = RGBColor(0, 0, 0)
        return style

    def _create_id_map(self, doc: Document) -> dict[int, str]:
        numbering = doc.part.numbering_part
        list_type_dict = {}
        abstractNumId_numId = {}

        if numbering is not None:
            for n in numbering.element.findall('.//w:num', namespaces=doc.part.element.nsmap):
                numId = n.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numId')
                abstractNumId = n.find('.//w:abstractNumId', namespaces=doc.part.element.nsmap).get(
                    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                abstractNumId_numId[int(abstractNumId)] = int(numId)
            for num in numbering.element.findall('.//w:abstractNum', namespaces=doc.part.element.nsmap):
                for lvl in num.findall('.//w:lvl', namespaces=doc.part.element.nsmap):
                    if int(lvl.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ilvl')) != 0:
                        continue
                    for fmt in lvl.findall('.//w:numFmt', namespaces=doc.part.element.nsmap):
                        abstractNumId = int(
                            num.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}abstractNumId'))
                        list_type_dict[abstractNumId_numId[abstractNumId]] = fmt.get(
                            '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
            return list_type_dict

    def create_numlist(self):
        if "1list num" in [x.name for x in self.nf.styles]:
            style = self.nf.styles["1list num"]
        else:
            style = self.nf.styles.add_style("1list num", WD_STYLE_TYPE.PARAGRAPH)
        style.quick_style = True
        style.base_style = self.nf.styles["main"]
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        style.paragraph_format.first_line_indent = Mm(-10)
        style.paragraph_format.left_indent = Mm(22.5)
        style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE

        style.font.name = "Times New Roman"
        style.font.size = Pt(14)
        style.font.bold = False
        style.font.color.rgb = RGBColor(0, 0, 0)
        return style

    def create_source_header(self) -> ParagraphStyle:
        if "source_header" in [x.name for x in self.nf.styles]:
            style = self.nf.styles["source_header"]
        else:
            style = self.nf.styles.add_style("source_header", WD_STYLE_TYPE.PARAGRAPH)
        style.quick_style = True
        style.base_style = self.nf.styles['Normal']
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        style.paragraph_format.first_line_indent = Mm(0)
        style.paragraph_format.left_indent = Mm(12.5)
        style.paragraph_format.space_before = Mm(6)
        style.paragraph_format.space_after = Mm(6)
        style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        style.font.name = "Times New Roman"
        style.font.size = Pt(14)
        style.font.bold = False
        style.font.all_caps = True
        style.font.color.rgb = RGBColor(0, 0, 0)
        return style

    def create_picture(self) -> ParagraphStyle:
        if "picture" in [x.name for x in self.nf.styles]:
            style = self.nf.styles["picture"]
        else:
            style = self.nf.styles.add_style("picture", WD_STYLE_TYPE.PARAGRAPH)
        style.quick_style = True
        style.base_style = self.nf.styles['Normal']
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        style.paragraph_format.first_line_indent = Mm(0)
        style.paragraph_format.left_indent = Mm(0)
        style.paragraph_format.space_after = Mm(6)
        style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        style.font.name = "Times New Roman"
        style.font.size = Pt(12)
        style.font.bold = True
        style.font.color.rgb = RGBColor(0, 0, 0)
        return style

    def create_bullet_list(self) -> ParagraphStyle:
        if "1list bullet" in [x.name for x in self.nf.styles]:
            style = self.nf.styles["1list bullet"]
        else:
            style = self.nf.styles.add_style("1list bullet", WD_STYLE_TYPE.PARAGRAPH)
        style.quick_style = True
        style.base_style = self.nf.styles['List Bullet']
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        style.paragraph_format.first_line_indent = Mm(-10)
        style.paragraph_format.left_indent = Mm(22.5)
        style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        style.font.name = "Times New Roman"
        style.font.size = Pt(14)
        style.font.bold = False
        style.font.color.rgb = RGBColor(0, 0, 0)
        return style

    def create_header1(self) -> ParagraphStyle:
        if "header1" in [x.name for x in self.nf.styles]:
            style = self.nf.styles["header1"]
        else:
            style = self.nf.styles.add_style("header1", WD_STYLE_TYPE.PARAGRAPH)
        style.quick_style = True
        style.base_style = self.nf.styles['Heading 1']
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        style.paragraph_format.first_line_indent = Mm(0)
        style.paragraph_format.left_indent = Mm(12.5)
        style.paragraph_format.space_after = Mm(10)
        style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        style.paragraph_format.page_break_before = True
        style.font.name = "Times New Roman"
        style.font.size = Pt(18)
        style.font.bold = True
        style.font.all_caps = True
        style.font.color.rgb = RGBColor(0, 0, 0)
        return style

    def create_header2(self) -> ParagraphStyle:
        if "header2" in [x.name for x in self.nf.styles]:
            style = self.nf.styles["header2"]
        else:
            style = self.nf.styles.add_style("header2", WD_STYLE_TYPE.PARAGRAPH)
        style.quick_style = True
        style.base_style = self.nf.styles['Heading 2']
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        style.paragraph_format.first_line_indent = Mm(0)
        style.paragraph_format.left_indent = Mm(12.5)
        style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        style.paragraph_format.space_before = Mm(15)
        style.paragraph_format.space_after = Mm(10)
        style.paragraph_format.keep_together = True
        style.font.name = "Times New Roman"
        style.font.size = Pt(16)
        style.font.bold = True
        style.font.color.rgb = RGBColor(0, 0, 0)
        return style

    def create_main(self) -> ParagraphStyle:
        if "main" in [x.name for x in self.nf.styles]:
            style = self.nf.styles["main"]
        else:
            style = self.nf.styles.add_style("main", WD_STYLE_TYPE.PARAGRAPH)
        style.quick_style = True
        style.base_style = self.nf.styles['Normal']
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        style.paragraph_format.first_line_indent = Mm(12.5)
        style.paragraph_format.left_indent = Mm(0)
        style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        style.font.name = "Times New Roman"
        style.font.size = Pt(14)
        style.font.color.rgb = RGBColor(0, 0, 0)
        return style
    def create_eq(self):
        if "eq" in [x.name for x in self.nf.styles]:
            style = self.nf.styles["eq"]
        else:
            style = self.nf.styles.add_style("eq", WD_STYLE_TYPE.PARAGRAPH)
        style.quick_style = True
        style.base_style = self.nf.styles['main']
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        style.paragraph_format.first_line_indent = Mm(12.5)
        style.paragraph_format.left_indent = Mm(0)
        style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        style.font.name = "Times New Roman"
        style.font.size = Pt(14)
        style.font.color.rgb = RGBColor(0, 0, 0)
        return style
    def __init__(self, nf: Document):
        self.nf = nf
        Decider._list_ids = self._create_id_map(nf)

        self.main = self.create_main()
        self.header1 = self.create_header1()
        self.header2 = self.create_header2()

        self.source_header = self.create_source_header()

        self.pictures = self.create_picture()
        self.eq = self.create_eq()
        self.lists_nums = self.create_numlist()
        self.list_bullet = self.create_bullet_list()


class Decider:
    _custom_style_names = False
    _csn_dict = {}
    _list_ids = {}

    @staticmethod
    def get_style(p: Paragraph, style: Styles) -> ParagraphStyle | int:
        if Decider._custom_style_names:
            return Decider._get_custom_name_style(p, style)
        # if p.style.type == WD_STYLE_TYPE.PARAGRAPH:
        #     if Decider._list_type(p):
        #         return style.list_bullet if Decider._list_type(p) == "bullet" else style.lists_nums
        #
        #
        #     if "рисунок" in p.text.lower() and "(рисунок " not in p.text.lower():
        #         return style.pictures
        #
        #     if p.style.base_style:
        #         if p.style.base_style.name == "Heading 1" or p.style.name == "Heading 1":
        #             return style.header1
        #         elif p.style.base_style.name == "Heading 2" or p.style.name == "Heading 2":
        #             return style.header2
        #
        #     if "часть" in p.text.lower() and len(p.text.split()) <= 3:
        #         return style.source_header
        #
        #     return style.main
        # else:
        #     return 0

    @staticmethod
    def custom_names(s: Styles):
        Decider._custom_style_names = True
        with open("input/styles.txt", encoding="utf-8") as f:
            stxt = list(map(lambda x: x.replace("\n", ""), f.readlines()))
            style_dict = {
                "main": s.main,
                "header1": s.header1,
                "header2": s.header2,
                "source_header": s.source_header,
                "picture": s.pictures,
                "list bullet": s.list_bullet,
                "list num": "list num"
            }
            list_num = []
            for li in stxt:
                if len(li.split(": ")) > 1:
                    key, value = li.split(": ")
                    if key.startswith("list num"):
                        list_num = value.split(", ")
                        continue
                    for v in value.split(", "):
                        Decider._csn_dict[v] = style_dict[key]
            for name in list_num:
                if name:
                    s._create_numlist_style(name)

    @staticmethod
    def _get_custom_name_style(p: Paragraph, style: Styles) -> ParagraphStyle | int:
        if p.style.type == WD_STYLE_TYPE.PARAGRAPH:
            if Decider._is_eq(p):
                return style.main
            if not p.text:
                return 0
            if Decider._list_type(p):
                return style.list_bullet if Decider._list_type(p) == "bullet" else style.lists_nums

            if f"1list num {p.style.name}" in [x.name for x in style.nf.styles]:
                return style.nf.styles[f"1list num {p.style.name}"]

            if p.style.name in Decider._csn_dict:
                if Decider._csn_dict[p.style.name] != "list num":
                    return Decider._csn_dict[p.style.name]
            else:
                if "рисунок" in p.text.lower() and "(рисунок " not in p.text.lower():
                    return style.pictures
                if "часть" in p.text.lower() and len(p.text.split()) <= 3:
                    return style.source_header
                print(f"Style {p.style.name} не указан в файле styles.txt", p.text[:10])
                return 0
        return 0

    @staticmethod
    def _list_type(paragraph):
        _el = paragraph._element
        p_xml = _el.find('.//w:pPr', namespaces=paragraph.part.element.nsmap)
        if p_xml is None:
            return 0
        a = p_xml.find('.//w:numPr', namespaces=paragraph.part.element.nsmap)
        if a is None:
            if "маркир" in paragraph.style.name.lower() or "bullet" in paragraph.style.name.lower():
                return "bullet"
            if paragraph.style.base_style:
                if "маркир" in paragraph.style.base_style.name.lower() or "bullet" in paragraph.style.base_style.name.lower():
                    return "bullet"
            return 0
        b = a.find('.//w:numId', namespaces=paragraph.part.element.nsmap)
        id = int(b.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val'))
        return Decider._list_ids[id]

    @staticmethod
    def _is_eq(p: Paragraph) -> bool:
        el = p._element
        m = el.find('.//m:oMath', namespaces=p.part.element.nsmap)
        return m is not None