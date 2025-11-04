import pandas as pd
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

# Load Excel
df = pd.read_excel('Conversion_month.xlsx')

# Get all 33 CHF entries with Min Qty = 36, Size = 75, Campaign Sub-Type = Normal
matches = df[(df['Unit Price'] == 33.0) &
             (df['Minimum Quantity'] == 36) &
             (df['Size'] == 75.0) &
             (df['Campaign Sub-Type'] == 'Normal')]

print("=== All 33 CHF candidates (Min Qty=36, Size=75, Normal) ===")
print(matches[['Wine Name', 'Producer Name', 'Unit Price (EUR)']].to_string())

context_wine = "Aalto 2023"
context_producer = "Aalto"  # Extracted from "Aalto 2023"
chf_price = 33.0

print(f"\nContext wine: '{context_wine}'")
print(f"Context producer: '{context_producer}'")
print(f"\n=== Scoring each candidate ===")

for idx, row in matches.iterrows():
    wine_name = row['Wine Name']
    producer_name = row['Producer Name']
    eur_value = row['Unit Price (EUR)']

    # Calculate wine similarity
    wine_similarity = calculate_similarity(context_wine, wine_name) if context_wine else 0
    wine_score = wine_similarity * 2.0

    # Calculate producer similarity
    producer_similarity = calculate_similarity(context_producer, producer_name) if context_producer and pd.notna(producer_name) else 0
    producer_score = producer_similarity * 1.5

    # Calculate price proximity
    expected_eur = chf_price * 1.08
    price_diff = abs(eur_value - expected_eur)
    price_proximity = max(0, 1.0 - (price_diff / expected_eur))
    price_score = price_proximity * 0.5

    total_score = wine_score + producer_score + price_score

    print(f"\n{wine_name} -> {eur_value} EUR")
    print(f"  Wine similarity: {wine_similarity:.3f} x 2.0 = {wine_score:.3f}")
    print(f"  Producer similarity: {producer_similarity:.3f} x 1.5 = {producer_score:.3f}")
    print(f"  Price proximity: {price_proximity:.3f} x 0.5 = {price_score:.3f}")
    print(f"  TOTAL SCORE: {total_score:.3f}")
