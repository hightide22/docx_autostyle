from docx import Document


def get_paragraphs_with_numFmt(file_path):
    # Открываем документ
    doc = Document(file_path)

    # Получаем элемент нумерации
    numbering = doc.part.numbering_part

    # Проверяем, есть ли элементы нумерации
    if numbering is not None:
        # Идем по всем абзацам в документе
        for paragraph in doc.paragraphs:
            if paragraph._element.find('.//w:numId', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                num_id = paragraph._element.find('.//w:numId', namespaces={
                    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})[0].text

                # Находим формат нумерации
                abstract_num = numbering.element.find(f'.//w:num[@w:numId="{num_id}"]/w:abstractNumId', namespaces={
                    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                if abstract_num is not None:
                    abstract_num_id = abstract_num.text
                    num_format_elem = numbering.element.find(
                        f'.//w:abstractNum[@w:abstractNumId="{abstract_num_id}"]/w:level/w:numFmt',
                        namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                    if num_format_elem is not None:
                        num_format = num_format_elem.get('w:val')
                        print(f'Paragraph: "{paragraph.text}"')
                        print(f'  Num ID: {num_id}, Num Format: {num_format}')
            else:
                print(f'Paragraph: "{paragraph.text}" does not have numbering.')

file_path = "eq.docx"
get_paragraphs_with_numFmt(file_path)