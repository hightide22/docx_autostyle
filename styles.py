from docx.document import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Mm
from docx.styles.style import ParagraphStyle
from docx.text.paragraph import Paragraph
import enum
import docx


f = docx.Document("empty.docx")
listnum = docx.Document("new.docx")
# nf = docx.Document("c.docx")
all_styles = f.styles


def create_main(d: Document) -> ParagraphStyle:
    if "main" in [x.name for x in d.styles]:
        style = d.styles["main"]
    else:
        style = d.styles.add_style("main", WD_STYLE_TYPE.PARAGRAPH)
    style.quick_style = True
    style.base_style = d.styles['Normal']
    style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    style.paragraph_format.first_line_indent = Mm(12.5)
    style.font.name = "Times New Roman"
    style.font.size = Pt(14)
    return style

def create_header1(d: Document) -> ParagraphStyle:
    if "header1" in [x.name for x in d.styles]:
        style = d.styles["header1"]
    else:
        style = d.styles.add_style("header1", WD_STYLE_TYPE.PARAGRAPH)
    style.quick_style = True
    style.base_style = d.styles['Heading 1']
    style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    style.paragraph_format.first_line_indent = Mm(0)
    style.paragraph_format.left_indent = Mm(12.5)
    style.font.name = "Times New Roman"
    style.font.size = Pt(18)
    style.font.bold = True
    style.font.all_caps = True
    return style

def create_header2(d: Document) -> ParagraphStyle:
    if "header2" in [x.name for x in d.styles]:
        style = d.styles["header2"]
    else:
        style = d.styles.add_style("header2", WD_STYLE_TYPE.PARAGRAPH)
    style.quick_style = True
    style.base_style = d.styles['Heading 2']
    style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    style.paragraph_format.first_line_indent = Mm(0)
    style.paragraph_format.left_indent = Mm(12.5)
    style.font.name = "Times New Roman"
    style.font.size = Pt(16)
    style.font.bold = True
    return style

def create_bullet_list(d: Document) -> ParagraphStyle:
    if "list bullet" in [x.name for x in d.styles]:
        style = d.styles["list bullet"]
    else:
        style = d.styles.add_style("list bullet", WD_STYLE_TYPE.PARAGRAPH)
    style.quick_style = True
    style.base_style = d.styles['List Bullet']
    style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    style.paragraph_format.first_line_indent = Mm(-10)
    style.paragraph_format.left_indent = Mm(22.5)
    style.font.name = "Times New Roman"
    style.font.size = Pt(14)
    style.font.bold = False
    return style


def create_num_list(d: Document, n=0) -> ParagraphStyle:
    listnum = docx.Document("new.docx")
    if n > 1:
        if f"list num {n}" in [x.name for x in d.styles]:
            style = d.styles[f"list num {n}"]
        else:
            style = d.styles.add_style(f"list num {n}", WD_STYLE_TYPE.PARAGRAPH)
        style.quick_style = True
        style.base_style = listnum.styles[f'n{n}']
    else:
        if "list num" in [x.name for x in d.styles]:
            style = d.styles["list num"]
        else:
            style = d.styles.add_style("list num", WD_STYLE_TYPE.PARAGRAPH)
        style.quick_style = True
        style.base_style = listnum.styles['n1']
    style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    style.paragraph_format.first_line_indent = Mm(-10)
    style.paragraph_format.left_indent = Mm(22.5)
    style.font.name = "Times New Roman"
    style.font.size = Pt(14)
    style.font.bold = False
    return style

def create_num_lists(d: Document) -> list[ParagraphStyle]:
    names = [x.name for x in d.styles if "номерованный список" in x.name]
    return [create_num_list(d, x) for x in range(1, 5)]



def create_picture(d: Document) -> ParagraphStyle:
    if "picture" in [x.name for x in d.styles]:
        style = d.styles["picture"]
    else:
        style = d.styles.add_style("picture", WD_STYLE_TYPE.PARAGRAPH)
    style.quick_style = True
    style.base_style = d.styles['Normal']
    style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    style.paragraph_format.first_line_indent = Mm(0)
    style.paragraph_format.left_indent = Mm(0)
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)
    style.font.bold = True
    return style

class Styles:
    def get_numlist_style(self, p: Paragraph):
        if f"list num {p.style.name}" in [x.name for x in self.nf.styles]:
            style = self.nf.styles[f"list num {p.style.name}"]
        else:
            style = self.nf.styles.add_style(f"list num {p.style.name}", WD_STYLE_TYPE.PARAGRAPH)
        style.quick_style = True
        style.base_style = p.style
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        style.paragraph_format.first_line_indent = Mm(-10)
        style.paragraph_format.left_indent = Mm(22.5)
        style.font.name = "Times New Roman"
        style.font.size = Pt(14)
        style.font.bold = False
        return style

    def __init__(self, nf: Document):
        self.nf = nf

        self.main = create_main(nf)
        self.header1 = create_header1(nf)
        self.header2 = create_header2(nf)

        self.code = all_styles["code"]
        self.sources = all_styles["sources"]
        self.pictures = create_picture(nf)

        self.lists_nums = create_num_lists(nf) # list!!

        self.list_bullet = create_bullet_list(nf)