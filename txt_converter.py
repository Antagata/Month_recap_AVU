#!/usr/bin/env python3
"""
Text File CHF to EUR Converter
Converts CHF prices to EUR in text files using Stock Lines.xlsx matching
Also generates filtered Stock Lines Excel with recognized items
"""

import re
import pandas as pd
from datetime import datetime
from pathlib import Path
from collections import defaultdict
import math

# Configuration
BASE_DIR = r"C:\Users\Marco.Africani\Desktop\Month recap"
DATABASE_DIR = r"C:\Users\Marco.Africani\OneDrive - AVU SA\AVU CPI Campaign\Puzzle_control_Reports\SOURCE_FILES"
INPUT_FILE_PATH = rf"{BASE_DIR}\Inputs\Multi.txt"
EXCEL_FILE_PATH = rf"{DATABASE_DIR}\OMT Main Offer List.xlsx"
STOCK_FILE_PATH = rf"{DATABASE_DIR}\Stock Lines.xlsx"
OUTPUTS_DIR = rf"{BASE_DIR}\Outputs"
RECOGNITION_REPORT_DIR = rf"{BASE_DIR}\Outputs\Detailed match results"

# Regex patterns for price detection
# Pattern for EUR prices with or without decimals: "EUR 202.00", "EUR 202", "EUR 30.00"
EUR_PRICE_PATTERN = r'\bEUR\s+(\d+(?:\.\d{2})?)\b'

# Pattern for CHF prices (if any exist in the text)
CHF_PRICE_PATTERN = r'\bCHF\s+(\d+(?:\.\d{2})?)\b'


def load_stock_lines():
    """Load Stock Lines.xlsx for direct Item No. to EUR price matching"""
    try:
        print(f"Loading Stock Lines database from: {STOCK_FILE_PATH}")
        stock_df = pd.read_excel(STOCK_FILE_PATH)
        print(f"[OK] Loaded {len(stock_df)} items from Stock Lines.xlsx")
        return stock_df
    except FileNotFoundError:
        print(f"[WARN] Stock Lines.xlsx not found at {STOCK_FILE_PATH}")
        return None
    except Exception as e:
        print(f"[WARN] Error loading Stock Lines.xlsx: {e}")
        return None


def load_omt_data():
    """Load OMT Main Offer List for EUR price lookups"""
    try:
        print(f"Loading OMT Main Offer List from: {EXCEL_FILE_PATH}")
        df = pd.read_excel(EXCEL_FILE_PATH)
        print(f"[OK] Loaded {len(df)} rows from OMT Main Offer List.xlsx")
        return df
    except FileNotFoundError:
        print(f"[WARN] OMT Main Offer List.xlsx not found at {EXCEL_FILE_PATH}")
        return None
    except Exception as e:
        print(f"[WARN] Error loading OMT Main Offer List.xlsx: {e}")
        return None


def extract_wine_name_and_vintage(text_before_price, text_after_price):
    """
    Extract wine name and vintage from text surrounding a price

    Expected format: "- Wine Name VINTAGE: description EUR price"

    Args:
        text_before_price: Text before the price (150 chars)
        text_after_price: Text after the price (20 chars)

    Returns:
        tuple: (wine_name, vintage)
    """
    # Look for vintage (4-digit year between 1900-2099)
    vintage_match = re.search(r'\b(19\d{2}|20\d{2})\b', text_before_price)
    vintage = int(vintage_match.group(1)) if vintage_match else None

    wine_name = None

    # Pattern 1: Look for "Wine Name VINTAGE:" format (most common)
    # This matches things like "Evangile 2015:" or "- Palmer 2015:"
    if vintage:
        # Look for the wine name before the vintage
        # Pattern: optional dash/whitespace, then wine name, then vintage, then colon
        pattern = r'[-–—\s]*([A-Za-zÀ-ÿ\s\.]+?)\s+' + str(vintage) + r'\s*:'
        match = re.search(pattern, text_before_price)
        if match:
            wine_name = match.group(1).strip()
            # Remove leading dashes
            wine_name = re.sub(r'^[-–—\s]+', '', wine_name)
            # Clean up common prefixes
            wine_name = re.sub(r'^(Château|Chateau|Domaine|Dom\.|Ch\.)\s+', '', wine_name, flags=re.IGNORECASE)
            return wine_name, vintage

    # Pattern 2: No colon, just "Wine Name VINTAGE" at the end of the line
    if vintage:
        # Get the part before the vintage
        before_vintage = text_before_price.split(str(vintage))[0]
        # Get the last 50 chars (or less) to find the wine name
        relevant_part = before_vintage[-50:] if len(before_vintage) > 50 else before_vintage
        # Remove everything before the last dash or newline
        parts = re.split(r'[-–—\n]', relevant_part)
        wine_name = parts[-1].strip() if parts else relevant_part.strip()
        # Clean up
        wine_name = re.sub(r'^(Château|Chateau|Domaine|Dom\.|Ch\.)\s+', '', wine_name, flags=re.IGNORECASE)
        if wine_name:
            return wine_name, vintage

    return wine_name, vintage


def match_price_via_stock_lines(eur_price, wine_name, vintage, stock_df, omt_df):
    """
    Match EUR price to Stock Lines item using wine name and vintage

    Args:
        eur_price: EUR price as float
        wine_name: Wine name extracted from text
        vintage: Vintage year
        stock_df: Stock Lines DataFrame
        omt_df: OMT Main Offer List DataFrame

    Returns:
        dict with item info or None
    """
    if stock_df is None or omt_df is None:
        return None

    try:
        # Step 1: Convert OMT Item No. to int for comparison
        omt_df_copy = omt_df.copy()
        omt_df_copy['Item No. Int'] = pd.to_numeric(omt_df_copy['Item No.'], errors='coerce').astype('Int64')

        # Step 2: Find in OMT by EUR price (with tolerance for rounding)
        eur_float = float(eur_price)
        # Try exact match first
        omt_matches = omt_df_copy[
            omt_df_copy['Unit Price (EUR)'].astype(float).round(0) == round(eur_float, 0)
        ]

        if len(omt_matches) == 0:
            return None

        # Step 3: Filter by wine name similarity and vintage
        best_match = None
        best_score = 0

        for idx, omt_row in omt_matches.iterrows():
            score = 0
            omt_wine = str(omt_row.get('Wine Name', '')).lower().strip()
            omt_vintage = omt_row.get('Vintage Code', None)

            # Wine name similarity (more lenient)
            if wine_name:
                wine_name_lower = wine_name.lower().strip()
                # Direct match or substring match
                if wine_name_lower in omt_wine or omt_wine in wine_name_lower:
                    score += 10
                # Partial word match
                wine_words = wine_name_lower.split()
                for word in wine_words:
                    if len(word) > 3 and word in omt_wine:
                        score += 3

            # Vintage match
            if vintage and omt_vintage:
                try:
                    if int(omt_vintage) == vintage:
                        score += 10
                except:
                    pass
            elif not wine_name:
                # If no wine name, rely more on price+vintage
                score += 5

            # Even without wine name match, if vintage matches and price matches, that's good
            if score > best_score:
                best_score = score
                best_match = omt_row

        # Accept match even with low score if price matches
        if best_match is not None and (best_score > 0 or len(omt_matches) == 1):
            item_no = int(best_match['Item No. Int'])

            # Step 4: Find in Stock Lines
            stock_match = stock_df[stock_df['No.'] == item_no]

            if len(stock_match) > 0:
                return {
                    'item_no': item_no,
                    'wine_name': best_match.get('Wine Name', ''),
                    'vintage': best_match.get('Vintage Code', ''),
                    'size': best_match.get('Size', ''),
                    'producer': best_match.get('Producer Name', ''),
                    'eur_price': eur_price,
                    'stock_row': stock_match.iloc[0]
                }

        return None

    except Exception as e:
        print(f"[WARN] Error matching price {eur_price}: {e}")
        return None


def convert_txt_file():
    """Main conversion function for txt files"""
    print("\n" + "="*80)
    print("TXT File CHF to EUR Converter")
    print("="*80 + "\n")

    # Load databases
    stock_df = load_stock_lines()
    omt_df = load_omt_data()

    if stock_df is None or omt_df is None:
        print("\n[ERROR] Cannot proceed without database files")
        return

    # Read input text file
    try:
        with open(INPUT_FILE_PATH, 'r', encoding='utf-8') as f:
            text = f.read()
        print(f"[OK] Loaded input file: {INPUT_FILE_PATH}")
    except FileNotFoundError:
        print(f"[ERROR] Input file not found: {INPUT_FILE_PATH}")
        return
    except Exception as e:
        print(f"[ERROR] Error reading input file: {e}")
        return

    # Find all EUR prices in the text
    recognized_items = []
    all_matches = []

    for match in re.finditer(EUR_PRICE_PATTERN, text):
        eur_price_str = match.group(1)
        eur_price = float(eur_price_str)
        position = match.start()

        # Extract context (150 chars before, 20 after)
        text_before = text[max(0, position-150):position]
        text_after = text[position:min(len(text), position+20)]

        # Extract wine name and vintage
        wine_name, vintage = extract_wine_name_and_vintage(text_before, text_after)

        # Try to match with Stock Lines
        item_info = match_price_via_stock_lines(eur_price, wine_name, vintage, stock_df, omt_df)

        if item_info:
            recognized_items.append(item_info)
            # Remove non-ASCII characters for printing
            safe_wine_name = wine_name.encode('ascii', errors='ignore').decode('ascii') if wine_name else ''
            print(f"[OK] Matched: {safe_wine_name} {vintage} - EUR {eur_price} -> Item No. {item_info['item_no']}")
        else:
            safe_wine_name = wine_name.encode('ascii', errors='ignore').decode('ascii') if wine_name else ''
            print(f"[WARN] Not matched: {safe_wine_name} {vintage} - EUR {eur_price}")

        all_matches.append({
            'wine_name': wine_name,
            'vintage': vintage,
            'eur_price': eur_price,
            'matched': item_info is not None
        })

    # Generate output txt file (converted prices)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_txt_path = rf"{OUTPUTS_DIR}\Multi_converted_{timestamp}.txt"

    # For now, just copy the text as-is (EUR prices are already in EUR)
    with open(output_txt_path, 'w', encoding='utf-8') as f:
        f.write(text)

    print(f"\n[OK] Output txt file saved: {output_txt_path}")

    # Generate filtered Stock Lines Excel
    if recognized_items:
        filtered_stock_rows = []
        item_nos = set()

        for item in recognized_items:
            item_no = item['item_no']
            if item_no not in item_nos:
                item_nos.add(item_no)
                filtered_stock_rows.append(item['stock_row'])

        if filtered_stock_rows:
            filtered_df = pd.DataFrame(filtered_stock_rows)
            output_excel_path = rf"{OUTPUTS_DIR}\Stock_Lines_Filtered_{timestamp}.xlsx"
            filtered_df.to_excel(output_excel_path, index=False)
            print(f"[OK] Filtered Stock Lines Excel saved: {output_excel_path}")
            print(f"   Contains {len(filtered_df)} recognized items")

    # Generate Lines.xlsx with recognized wine details
    if recognized_items:
        lines_data = []
        for item in recognized_items:
            lines_data.append({
                'Wine Name': item['wine_name'],
                'Vintage': item['vintage'],
                'Size': item['size'],
                'Producer': item['producer'],
                'EUR Price': item['eur_price'],
                'Item No.': item['item_no']
            })

        if lines_data:
            lines_df = pd.DataFrame(lines_data)
            lines_excel_path = rf"{OUTPUTS_DIR}\Lines.xlsx"
            lines_df.to_excel(lines_excel_path, index=False)
            print(f"[OK] Lines.xlsx saved with {len(lines_data)} recognized wines")

    # Generate recognition report in Detailed match results folder
    # Create directory if it doesn't exist
    Path(RECOGNITION_REPORT_DIR).mkdir(parents=True, exist_ok=True)
    report_path = rf"{RECOGNITION_REPORT_DIR}\Recognition_Report_{timestamp}.txt"
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write("="*80 + "\n")
        f.write("WINE RECOGNITION REPORT\n")
        f.write("="*80 + "\n\n")
        f.write(f"Total prices found: {len(all_matches)}\n")
        f.write(f"Successfully matched: {len(recognized_items)}\n")
        f.write(f"Not matched: {len(all_matches) - len(recognized_items)}\n\n")
        f.write("="*80 + "\n")
        f.write("RECOGNIZED WINES:\n")
        f.write("="*80 + "\n\n")

        for item in recognized_items:
            f.write(f"Wine: {item['wine_name']}\n")
            f.write(f"Vintage: {item['vintage']}\n")
            f.write(f"EUR Price: {item['eur_price']}\n")
            f.write(f"Item No.: {item['item_no']}\n")
            f.write(f"Size: {item['size']}\n")
            f.write("-"*80 + "\n")

    print(f"[OK] Recognition report saved: {report_path}")

    # Statistics
    print("\n" + "="*80)
    print("CONVERSION STATISTICS")
    print("="*80)
    print(f"Total EUR prices found: {len(all_matches)}")
    print(f"[OK] Successfully matched to Stock Lines: {len(recognized_items)}")
    print(f"[WARN] Not matched: {len(all_matches) - len(recognized_items)}")
    print("="*80 + "\n")


if __name__ == "__main__":
    convert_txt_file()
