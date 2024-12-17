import docx

f = docx.Document("empty.docx")
all_styles = f.styles
print(all_styles["main"])