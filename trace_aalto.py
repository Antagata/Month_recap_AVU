from docx import Document
import re
from difflib import SequenceMatcher

def normalize_wine_name(name):
    if not isinstance(name, str):
        return ""
    name = name.lower()
    name = re.sub(r'\bch[âa]teau\b', '', name)
    name = re.sub(r'\bdomaine\b', '', name)
    name = re.sub(r'[^\w\s]', ' ', name)
    name = ' '.join(name.split())
    return name.strip()

def calculate_similarity(text1, text2):
    text1_norm = normalize_wine_name(text1)
    text2_norm = normalize_wine_name(text2)

    if not text1_norm or not text2_norm:
        return 0.0

    full_similarity = SequenceMatcher(None, text1_norm, text2_norm).ratio()

    if text1_norm in text2_norm or text2_norm in text1_norm:
        full_similarity = max(full_similarity, 0.7)

    words1 = set(text1_norm.split())
    words2 = set(text2_norm.split())
    filler_words = {'the', 'de', 'di', 'du', 'della', 'des', 'le', 'la', 'del'}
    words1 = words1 - filler_words
    words2 = words2 - filler_words

    if words1 and words2:
        word_overlap = len(words1 & words2) / len(words1 | words2)
        combined_similarity = max(full_similarity, word_overlap * 0.9)
    else:
        combined_similarity = full_similarity

    return combined_similarity

# Load document
doc = Document('month recap.docx')

# Find Aalto paragraph
for i, para in enumerate(doc.paragraphs):
    if 'Aalto 2023' in para.text and '33.00CHF' in para.text:
        print(f"Paragraph {i}:")
        print(para.text)
        print("\n=== Testing wine name matching ===")

        # Test similarity scores
        candidates = [
            "Aalto",
            "Il Pino di Biserno",
            "Phelan Segur",
            "La Croix Ducru-Beaucaillou"
        ]

        for wine in candidates:
            score = calculate_similarity("Aalto 2023", wine)
            print(f"Similarity: 'Aalto 2023' vs '{wine}' = {score:.3f}")
