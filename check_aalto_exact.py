from docx import Document

doc = Document('month recap_EUR.docx')

for i, para in enumerate(doc.paragraphs):
    if 'Aalto 2023' in para.text:
        print(f'Paragraph {i}:')
        print(repr(para.text))  # Use repr to see exact string
        print(para.text)
        print('---')

        # Extract just the numbers
        import re
        prices = re.findall(r'(\d+\.\d+)\s*EUR', para.text)
        print(f"EUR prices found: {prices}")
        break
