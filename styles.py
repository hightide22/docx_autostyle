from docx import Document
from docx.enum.style import WD_STYLE

"""
Можно заменить на enum (?)
"""

class Styles:
    def __init__(self):
        f = Document("empty.docx")
        all_styles = f.styles


        self.main = all_styles["main"]
        self.header1 = all_styles["header1"] # biggest
        self.header2 = all_styles["header2"] # smaller

        self.code = all_styles["code"]
        self.sources = all_styles["sources"]
        self.pictures = all_styles["pictures"]

        self.list_num =all_styles["list num"]
        self.list_bullet = all_styles["list bullet"]