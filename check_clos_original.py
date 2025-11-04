from docx import Document

doc = Document('month recap.docx')

for i, para in enumerate(doc.paragraphs):
    if 'Clos' in para.text and ('glise' in para.text or 'Église' in para.text) and '2009' in para.text:
        print(f"Clos l'Église - Paragraph {i}:")
        print(para.text)
        print('---')
