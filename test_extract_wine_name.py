import re

def extract_wine_name_from_context(text, price_match_start):
    context_before = text[max(0, price_match_start - 400):price_match_start]

    wine_candidates = []

    # Pattern 1: FIRST colon
    all_colons = list(re.finditer(r'([A-ZÀ-ÿ][^\n:]{3,60})[:]\s*', context_before))
    if all_colons:
        first_colon = all_colons[0]
        candidate = first_colon.group(1).strip()
        candidate = re.sub(r'\s+(at|from|for|with|the|a|an)$', '', candidate, flags=re.IGNORECASE)
        wine_candidates.append(candidate)

        if len(all_colons) > 1:
            last_colon = all_colons[-1]
            candidate_last = last_colon.group(1).strip()
            candidate_last = re.sub(r'\s+(at|from|for|with|the|a|an)$', '', candidate_last, flags=re.IGNORECASE)
            if candidate_last != candidate and not re.search(r'chf|price|offer', candidate_last, re.IGNORECASE):
                wine_candidates.append(candidate_last)

    # Pattern 2: Quoted text
    quote_matches = re.findall(r'["""]([^"""]{3,60})["""]', context_before)
    for match in quote_matches:
        if re.search(r'[A-ZÀ-ÿ]', match):
            wine_candidates.append(match.strip())

    # Pattern 3: Château/Domaine
    chateau_pattern = re.findall(r'\b([CcDd]h[âa]teau|Domaine|Dom\.)\s+([A-ZÀ-ÿ][^\n:,.]{3,40})', context_before)
    for prefix, name in chateau_pattern:
        wine_candidates.append(f"{prefix} {name}".strip())

    # Pattern 4: Producer name patterns
    producer_pattern = re.findall(
        r'\b([A-ZÀ-ÿ][a-zà-ÿ]+(?:\s+[A-ZÀ-ÿ][a-zà-ÿ]+){0,3})\s+(?:\d{4})?',
        context_before
    )

    for match in producer_pattern:
        if len(match) > 3:
            wine_candidates.append(match.strip())

    # Return the best candidate (prefer Pattern 1, then Pattern 3, then others)
    if wine_candidates:
        return wine_candidates[0]
    return None

# Test with Aalto paragraph
paragraph = "Aalto 2023: Brilliant year for both Aalto and PS Aalto - 34.00 CHF + VAT // 36x 33.00CHF + VAT"

# Find position of "33.00"
position_33 = paragraph.find("33.00")

print(f"Testing extraction for 33.00 at position {position_33}")
print(f"Paragraph: {paragraph}")
print(f"\nContext before 33.00: '{paragraph[max(0, position_33-400):position_33]}'")

wine_name = extract_wine_name_from_context(paragraph, position_33)
print(f"\nExtracted wine name: '{wine_name}'")
