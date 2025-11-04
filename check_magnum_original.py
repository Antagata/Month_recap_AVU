from docx import Document

doc = Document('month recap.docx')

for i, para in enumerate(doc.paragraphs):
    if 'Magnum' in para.text and '52' in para.text:
        print(f'Paragraph {i}:')
        print(para.text)
        print('---')
