from docx import Document

doc = Document('month recap_EUR.docx')

for i, para in enumerate(doc.paragraphs):
    if 'Aalto' in para.text and '2023' in para.text:
        print(f'Paragraph {i}:')
        print(para.text)
        print('---')
