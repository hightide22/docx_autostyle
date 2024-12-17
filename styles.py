from docx import Document

"""
Можно заменить на enum (?)
"""

class Styles:
    def __init__(self):
        f = Document("empty.docx")
        all_styles = f.styles
        self.main = all_styles["main"]
        self.list_bullet = all_styles["list bullet"]
        self.code = all_styles["code"]
        self.sources = all_styles["sources"]
        self.pictures = all_styles["pictures"]
        self.list_marked = all_styles["list marked"]
        self.list_num = all_styles["list num"]