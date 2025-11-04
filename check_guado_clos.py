from docx import Document

doc = Document('month recap_EUR.docx')

for i, para in enumerate(doc.paragraphs):
    if 'Guado al Tasso' in para.text and '2022' in para.text:
        print(f'Guado al Tasso - Paragraph {i}:')
        print(para.text)
        print('---\n')

    if 'Clos' in para.text and ('glise' in para.text or 'Église' in para.text) and '2009' in para.text:
        print(f"Clos l'Église - Paragraph {i}:")
        print(para.text)
        print('---\n')
