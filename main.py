from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_COLOR_INDEX, WD_COLOR
from styles import Styles
from text import Decider
from docx.shared import Mm, RGBColor
old_document = Document("input/work_c.docx")

styles = Styles(old_document)
docx_iter = old_document.iter_inner_content()

for p in docx_iter:
   r_style = Decider.get_style(p, styles)
   if not r_style:
      continue
   if p.style.type == WD_STYLE_TYPE.PARAGRAPH:
      p.style = r_style
      Decider.normalizer(p)

      if p.paragraph_format.left_indent != p.style.paragraph_format.left_indent and p.paragraph_format.left_indent and p.runs:
         p.paragraph_format.left_indent = p.style.paragraph_format.left_indent
         p.runs[0].font.highlight_color = WD_COLOR_INDEX.RED


old_document.save("output/work_c.docx")

