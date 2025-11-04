import re

def extract_vintage_from_context(text, price_match_start):
    # Look in wider context (up to 600 chars before to catch vintage at paragraph start)
    context = text[max(0, price_match_start - 600):min(len(text), price_match_start + 100)]

    # Find 4-digit years (vintage years typically 1990-2030)
    year_matches = re.findall(r'\b(19[9]\d|20[0-3]\d)\b', context)

    if year_matches:
        # Return the most recent/last mentioned year
        try:
            return int(year_matches[-1])
        except:
            return None

    return None

# Test with Clos paragraph
paragraph = "Clos l'église 2009: Ex Chateau Clos L'Eglise is a very small property in Pomerol, producing only about 1,000 cases per year. We tasted the 2009 before purchasing, and it is at a perfect stage-drinking beautifully now, yet, as Robert Parker noted in his review, it will continue to develop magnificently over the next 25 years. Parker rated it 98 points, placing it alongside some of Pomerol's most renowned neighbors that cost two to three times more. We believe this represents a truly exceptional value. Bordeaux's best-kept secrets - 105.00 CHF + VAT // 36x 99.00 + VAT"

# Find position of "99.00"
position_99 = paragraph.find("99.00")

print(f"Testing vintage extraction for 99.00 at position {position_99}")
context_600 = paragraph[max(0, position_99-600):min(len(paragraph), position_99+100)]
print(f"\nContext (600 chars): '{context_600[:100]}... {context_600[-100:]}'")

vintage = extract_vintage_from_context(paragraph, position_99)
print(f"\nExtracted vintage: {vintage}")
