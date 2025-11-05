#!/usr/bin/env python3
"""
Wine Item Number Matcher and Learning System
==============================================

This script:
1. Reads wine names + vintages from ItemNoGenerator.txt
2. Matches them against the Excel database
3. Generates a results table with Item Numbers
4. Updates a learning database to improve future wine name recognition

Usage:
    python wine_item_matcher.py [--input FILE] [--size SIZE]

Arguments:
    --input FILE    Input file with wine names (default: ItemNoGenerator.txt)
    --size SIZE     Filter by bottle size (default: 75.0 for standard bottles)
                    Options: 75.0, 150.0, 300.0, or omit for all sizes
"""

import pandas as pd
import sys
import io
import re
import argparse
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path

# Fix encoding for Windows console
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Configuration
EXCEL_FILE = r"C:\Users\Marco.Africani\Desktop\Month recap\Conversion_month.xlsx"
STOCK_FILE = r"C:\Users\Marco.Africani\Desktop\Month recap\Detailed Stock List.xlsx"  # Fallback database
INPUT_FILE = r"C:\Users\Marco.Africani\Desktop\Month recap\ItemNoGenerator.txt"
OUTPUT_DIR = r"C:\Users\Marco.Africani\Desktop\Month recap"
LEARNING_DB_FILE = r"C:\Users\Marco.Africani\Desktop\Month recap\wine_names_learning_db.txt"

# Excel columns - Conversion_month.xlsx
WINE_NAME_COL = 'Wine Name'
VINTAGE_COL = 'Vintage'
ITEM_NO_COL = 'Item No.'
PRODUCER_COL = 'Producer Name'
SIZE_COL = 'Size'

# Stock database columns - Detailed Stock List.xlsx
STOCK_ID_COL = 'ID'
STOCK_WINE_COL = 'Wine'
STOCK_PRODUCER_COL = 'Producer'
STOCK_VINTAGE_COL = 'Vintage'
STOCK_SIZE_COL = 'Size'


def normalize_wine_name(name):
    """
    Normalize wine name for better matching:
    - Convert to lowercase
    - Remove extra whitespace
    - Remove common punctuation
    - Remove château/domaine prefixes
    """
    if not isinstance(name, str):
        return ""

    # Convert to lowercase
    name = name.lower()

    # Remove château/chateau/domaine variations
    name = re.sub(r'\bch[âa]teau\b', '', name)
    name = re.sub(r'\bdomaine\b', '', name)
    name = re.sub(r'\bch\b\.?', '', name)  # Remove "Ch" or "Ch."

    # Remove special characters but keep spaces
    name = re.sub(r'[^\w\s]', ' ', name)

    # Remove extra whitespace
    name = ' '.join(name.split())

    return name.strip()


def calculate_similarity(text1, text2):
    """
    Calculate similarity ratio between two strings (0-1).
    Uses normalized versions for better matching.
    """
    text1_norm = normalize_wine_name(text1)
    text2_norm = normalize_wine_name(text2)

    if not text1_norm or not text2_norm:
        return 0.0

    # Full string similarity
    full_ratio = SequenceMatcher(None, text1_norm, text2_norm).ratio()

    # Check if one contains the other (partial match)
    if text1_norm in text2_norm or text2_norm in text1_norm:
        return max(full_ratio, 0.8)

    return full_ratio


def extract_vintage_from_text(text):
    """Extract vintage year from text (e.g., 'Lafite 2005' -> 2005)"""
    # Look for 4-digit year (1900-2099)
    match = re.search(r'\b(19\d{2}|20\d{2})\b', text)
    if match:
        return int(match.group(1))
    return None


def parse_input_file(file_path):
    """
    Parse input file to extract wine names and vintages.

    Expected formats:
    - "Wine Name Vintage" (e.g., "Lafite Rothschild 2005")
    - "Wine Name, Vintage" (e.g., "Lafite Rothschild, 2005")
    - "Wine Name | Vintage" (e.g., "Lafite Rothschild | 2005")
    - One entry per line
    """
    wines = []

    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()

        for line_num, line in enumerate(lines, 1):
            line = line.strip()

            # Skip empty lines and comments
            if not line or line.startswith('#'):
                continue

            # Try to extract vintage
            vintage = extract_vintage_from_text(line)

            # Extract wine name (remove vintage from the text)
            if vintage:
                wine_name = re.sub(r'\b' + str(vintage) + r'\b', '', line).strip()
                # Remove separators
                wine_name = re.sub(r'[,|]', ' ', wine_name).strip()
                wine_name = ' '.join(wine_name.split())  # Remove extra spaces
            else:
                wine_name = line

            if wine_name:
                wines.append({
                    'original_text': line,
                    'wine_name': wine_name,
                    'vintage': vintage,
                    'line_number': line_num
                })

    except FileNotFoundError:
        print(f"❌ Error: Input file not found at {file_path}")
        print(f"   Please create the file with wine names, one per line.")
        print(f"   Format: 'Wine Name Vintage' (e.g., 'Château Lafite Rothschild 2005')")
        return None
    except Exception as e:
        print(f"❌ Error reading input file: {e}")
        return None

    return wines


def load_learning_database(learning_db_path):
    """
    Load the learning database for quick lookups.
    Returns a dict: {(wine_name, vintage): item_no}
    """
    learning_map = {}

    if not Path(learning_db_path).exists():
        return learning_map

    try:
        with open(learning_db_path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#'):
                    # Parse: Wine Name | Vintage | Item No. | Timestamp
                    parts = line.split(' | ')
                    if len(parts) >= 3:
                        wine_name = parts[0].strip()
                        vintage_str = parts[1].strip()
                        item_no_str = parts[2].strip()

                        # Only add entries with valid Item Numbers (not NOT_FOUND)
                        if item_no_str and item_no_str != 'NOT_FOUND':
                            try:
                                item_no = int(item_no_str)
                                # Convert vintage to int or None
                                vintage = int(vintage_str) if vintage_str != 'N/A' else None
                                # Create lookup key
                                key = (wine_name.lower(), vintage)
                                learning_map[key] = item_no
                            except ValueError:
                                # Skip invalid entries
                                pass

        if learning_map:
            print(f"✅ Loaded {len(learning_map)} entries from learning database")

    except Exception as e:
        print(f"⚠️  Warning: Could not read learning database: {e}")

    return learning_map


def load_excel_database(excel_path):
    """Load wine database from Excel"""
    try:
        df = pd.read_excel(excel_path)

        # Convert vintage to int for matching
        df['Vintage_Int'] = df[VINTAGE_COL].apply(
            lambda x: int(x) if pd.notna(x) and str(x).isdigit() else None
        )

        # Sort by Schedule DateTime (descending) so latest campaigns are prioritized
        if 'Schedule DateTime' in df.columns:
            df = df.sort_values('Schedule DateTime', ascending=False, na_position='last')
            print(f"✅ Loaded {len(df)} wines from primary database (Conversion_month.xlsx)")
            print(f"   📅 Sorted by Schedule DateTime (latest campaigns first)")
        else:
            print(f"✅ Loaded {len(df)} wines from primary database (Conversion_month.xlsx)")

        return df

    except FileNotFoundError:
        print(f"❌ Error: Excel file not found at {excel_path}")
        return None
    except Exception as e:
        print(f"❌ Error loading Excel: {e}")
        return None


def load_stock_database(stock_path):
    """Load fallback wine stock database from Detailed Stock List.xlsx"""
    try:
        # Skip first 2 rows, header is on row 3 (0-indexed: skiprows=[0,1])
        df = pd.read_excel(stock_path, header=2)

        # Rename columns for consistency
        df = df.rename(columns={
            STOCK_ID_COL: 'Item_No',
            STOCK_WINE_COL: 'Wine_Name',
            STOCK_PRODUCER_COL: 'Producer',
            STOCK_VINTAGE_COL: 'Vintage',
            STOCK_SIZE_COL: 'Size'
        })

        # Convert vintage to int for matching
        df['Vintage_Int'] = df['Vintage'].apply(
            lambda x: int(x) if pd.notna(x) and str(x).replace('.0', '').isdigit() else None
        )

        print(f"✅ Loaded {len(df)} wines from fallback database (Detailed Stock List.xlsx)")
        return df

    except FileNotFoundError:
        print(f"⚠️  Warning: Stock database not found at {stock_path}")
        print(f"   Fallback matching will not be available")
        return None
    except Exception as e:
        print(f"⚠️  Warning: Could not load stock database: {e}")
        print(f"   Fallback matching will not be available")
        return None


def find_best_match(wine_name, vintage, df, threshold=0.6, learning_map=None, stock_df=None, preferred_size=75.0):
    """
    Find best matching wine using learning database first, then Excel database, then stock database.

    Priority:
    1. Learning database (exact match)
    2. Primary Excel database (fuzzy match)
    3. Stock database fallback (fuzzy match) - only if not found in primary

    Args:
        preferred_size: Preferred bottle size (default 75.0 for standard bottles)
                       Set to None to disable size filtering

    Returns:
        dict with match info, or None if no good match found
    """
    # PRIORITY 1: Check learning database first (fast, exact matches from corrections)
    if learning_map:
        # Try exact match (case-insensitive)
        key = (wine_name.lower(), vintage)
        if key in learning_map:
            item_no = learning_map[key]
            # Look up full details from Excel
            # Note: Item No. in Excel is stored as string, so convert to string for comparison
            item_row = df[df[ITEM_NO_COL] == str(item_no)]
            if len(item_row) > 0:
                row = item_row.iloc[0]
                # Check if size matches the preferred size filter
                row_size = row.get(SIZE_COL, '')
                if preferred_size is None or row_size == preferred_size:
                    return {
                        'wine_name': str(row.get(WINE_NAME_COL, '')),
                        'vintage': vintage,
                        'item_no': item_no,
                        'producer': str(row.get(PRODUCER_COL, '')),
                        'size': row_size,
                        'similarity': 1.0,  # Perfect match from learning DB
                        'source': 'learning_database'
                    }
                else:
                    print(f"      ⚠️  Learning DB has Item No. {item_no} (size: {row_size}cl), but size filter is {preferred_size}cl. Searching for correct size...")
            else:
                # Item Number from learning DB not found in primary Excel
                # Check stock database before falling back to fuzzy matching
                if stock_df is not None:
                    stock_row = stock_df[stock_df['Item_No'] == item_no]
                    if len(stock_row) > 0:
                        row = stock_row.iloc[0]
                        row_size = row.get('Size', '')
                        # Check if size matches the preferred size filter
                        if preferred_size is None or row_size == preferred_size:
                            return {
                                'wine_name': str(row.get('Wine_Name', '')),
                                'vintage': vintage,
                                'item_no': item_no,
                                'producer': str(row.get('Producer', '')),
                                'size': row_size,
                                'similarity': 1.0,  # Perfect match from learning DB
                                'source': 'learning_database_stock'
                            }
                        else:
                            print(f"      ⚠️  Learning DB has Item No. {item_no} (size: {row_size}cl), but size filter is {preferred_size}cl. Searching for correct size...")
                print(f"      ⚠️  Learning DB has Item No. {item_no}, but it's not in either database! Falling back to fuzzy matching...")

    # PRIORITY 2: Fuzzy matching in primary Excel (slower, but comprehensive)
    candidates = []

    # Filter by vintage if provided
    if vintage:
        df_filtered = df[df['Vintage_Int'] == vintage].copy()
    else:
        df_filtered = df.copy()

    # Filter by size if preferred_size is specified
    if preferred_size is not None and len(df_filtered) > 0:
        df_filtered = df_filtered[df_filtered[SIZE_COL] == preferred_size].copy()

    if len(df_filtered) > 0:
        # Calculate similarity for each wine
        for _, row in df_filtered.iterrows():
            excel_wine_name = str(row.get(WINE_NAME_COL, ''))
            similarity = calculate_similarity(wine_name, excel_wine_name)

            if similarity >= threshold:
                candidates.append({
                    'wine_name': excel_wine_name,
                    'vintage': row.get('Vintage_Int'),
                    'item_no': row.get(ITEM_NO_COL),
                    'producer': row.get(PRODUCER_COL, ''),
                    'size': row.get(SIZE_COL, ''),
                    'similarity': similarity,
                    'source': 'excel_primary'
                })

    # If we found candidates in primary Excel, return best match
    if candidates:
        best_match = max(candidates, key=lambda x: x['similarity'])
        return best_match

    # PRIORITY 3: Fuzzy matching in stock database (fallback)
    if stock_df is not None:
        print(f"      🔄 Not found in primary database, trying fallback stock database...")

        # Filter by vintage if provided
        if vintage:
            stock_filtered = stock_df[stock_df['Vintage_Int'] == vintage].copy()
        else:
            stock_filtered = stock_df.copy()

        # Filter by size if preferred_size is specified
        if preferred_size is not None and len(stock_filtered) > 0:
            stock_filtered = stock_filtered[stock_filtered['Size'] == preferred_size].copy()

        if len(stock_filtered) > 0:
            # Calculate similarity for each wine
            for _, row in stock_filtered.iterrows():
                stock_wine_name = str(row.get('Wine_Name', ''))
                similarity = calculate_similarity(wine_name, stock_wine_name)

                if similarity >= threshold:
                    candidates.append({
                        'wine_name': stock_wine_name,
                        'vintage': row.get('Vintage_Int'),
                        'item_no': row.get('Item_No'),
                        'producer': row.get('Producer', ''),
                        'size': row.get('Size', ''),
                        'similarity': similarity,
                        'source': 'stock_fallback'
                    })

        # Return best match from stock if found
        if candidates:
            best_match = max(candidates, key=lambda x: x['similarity'])
            return best_match

    # No match found in any database
    return None


def update_learning_database(wine_entries, learning_db_path):
    """
    Update the learning database with processed wine names.
    This database grows over time and helps improve future matching.
    Prevents duplicate entries (same wine + vintage + item_no).
    """
    # Load existing database as unique keys (wine|vintage|item_no)
    existing_keys = set()
    existing_lines = []

    if Path(learning_db_path).exists():
        try:
            with open(learning_db_path, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith('#'):
                        existing_lines.append(line)
                        # Extract key: wine_name|vintage|item_no (ignore timestamp)
                        parts = line.split(' | ')
                        if len(parts) >= 3:
                            key = f"{parts[0]}|{parts[1]}|{parts[2]}"
                            existing_keys.add(key)
        except Exception as e:
            print(f"⚠️  Warning: Could not read learning database: {e}")

    # Add new entries (only if unique)
    new_entries = []
    duplicate_count = 0

    for entry in wine_entries:
        wine_name = entry.get('wine_name', '')
        vintage = entry.get('vintage', '')
        item_no = entry.get('matched_item_no', '')

        if wine_name:
            # Create unique key
            vintage_str = str(vintage) if vintage else 'N/A'
            item_no_str = str(item_no) if item_no else 'NOT_FOUND'
            key = f"{wine_name}|{vintage_str}|{item_no_str}"

            # Only add if not already in database
            if key not in existing_keys:
                # Format: Wine Name | Vintage | Item No. | Timestamp
                entry_line = f"{wine_name} | {vintage_str} | {item_no_str} | {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
                new_entries.append(entry_line)
                existing_keys.add(key)
            else:
                duplicate_count += 1

    # Write updated database
    try:
        with open(learning_db_path, 'a', encoding='utf-8') as f:
            if not Path(learning_db_path).exists() or Path(learning_db_path).stat().st_size == 0:
                # Write header for new file
                f.write("# Wine Names Learning Database\n")
                f.write("# Format: Wine Name | Vintage | Item No. | Timestamp\n")
                f.write("# This file grows over time and helps improve wine name recognition\n")
                f.write("#" + "="*80 + "\n\n")

            for entry in new_entries:
                f.write(entry + "\n")

        if new_entries:
            print(f"\n✅ Added {len(new_entries)} new entries to learning database")
            if duplicate_count > 0:
                print(f"   ⏭️  Skipped {duplicate_count} duplicate entries")
            print(f"   Total unique entries in database: {len(existing_keys)}")
        else:
            print(f"\n✅ Learning database already up to date ({len(existing_keys)} unique entries)")
            if duplicate_count > 0:
                print(f"   ⏭️  All {duplicate_count} entries were duplicates (already in database)")

    except Exception as e:
        print(f"\n❌ Error updating learning database: {e}")


def create_correction_file(wine_entries, output_dir):
    """
    Create a correction file for wines that were NOT FOUND or have low similarity.
    This allows manual correction by adding the correct Item No.

    Format: Wine Name | Vintage | ITEM_NO_TO_ADD | Notes
    """
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    correction_file = Path(output_dir) / f"CORRECTIONS_NEEDED_{timestamp}.txt"

    # Find wines needing correction
    wines_needing_correction = []

    for entry in wine_entries:
        # Include if NOT FOUND or low similarity (< 80%)
        item_no = entry.get('matched_item_no')
        similarity = entry.get('similarity', 0)

        if not item_no or similarity < 0.8:
            wines_needing_correction.append(entry)

    if not wines_needing_correction:
        # No corrections needed
        return None

    try:
        with open(correction_file, 'w', encoding='utf-8') as f:
            # Header
            f.write("="*100 + "\n")
            f.write("WINE CORRECTIONS FILE\n")
            f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("="*100 + "\n\n")

            f.write("INSTRUCTIONS:\n")
            f.write("-"*100 + "\n")
            f.write("1. For each wine below, look up the correct Item Number in your Excel\n")
            f.write("2. Replace 'YOUR_ITEM_NO_HERE' with the actual Item Number\n")
            f.write("3. Save this file as 'CORRECTIONS_[timestamp].txt'\n")
            f.write("4. Run: python apply_corrections.py CORRECTIONS_[timestamp].txt\n")
            f.write("   (This will update the learning database with your corrections)\n")
            f.write("\n")
            f.write("Format: Wine Name | Vintage | Item No. | Notes\n")
            f.write("="*100 + "\n\n")

            # List wines needing correction
            for i, entry in enumerate(wines_needing_correction, 1):
                wine_name = entry.get('wine_name', '')
                vintage = str(entry.get('vintage', 'N/A'))
                original_text = entry.get('original_text', '')
                item_no = entry.get('matched_item_no', '')
                similarity = entry.get('similarity', 0)

                f.write(f"[{i}] {original_text}\n")

                if not item_no:
                    # Not found - needs manual lookup
                    f.write(f"{wine_name} | {vintage} | YOUR_ITEM_NO_HERE | NOT FOUND - Please add correct Item No.\n")
                else:
                    # Low similarity - needs verification
                    excel_name = entry.get('excel_wine_name', '')
                    f.write(f"{wine_name} | {vintage} | {item_no} | LOW SIMILARITY ({similarity:.1%}) - Matched to '{excel_name}' - Verify if correct\n")

                f.write("\n")

        print(f"\n⚠️  CORRECTIONS FILE CREATED: {correction_file}")
        print(f"   {len(wines_needing_correction)} wines need manual review")
        print(f"   Please review and correct Item Numbers in this file")

        return correction_file

    except Exception as e:
        print(f"\n❌ Error creating correction file: {e}")
        return None


def generate_output_report(wine_entries, output_dir):
    """Generate a formatted output report with results"""
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = Path(output_dir) / f"ItemNo_Results_{timestamp}.txt"

    try:
        with open(output_file, 'w', encoding='utf-8') as f:
            # Header
            f.write("="*100 + "\n")
            f.write("WINE ITEM NUMBER MATCHING RESULTS\n")
            f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("="*100 + "\n\n")

            # Summary statistics
            total = len(wine_entries)
            matched = sum(1 for e in wine_entries if e.get('matched_item_no'))
            not_matched = total - matched

            f.write("SUMMARY\n")
            f.write("-"*100 + "\n")
            f.write(f"Total wines processed: {total}\n")
            f.write(f"Successfully matched: {matched} ({matched/total*100:.1f}%)\n")
            f.write(f"Not matched: {not_matched} ({not_matched/total*100:.1f}%)\n")
            f.write("\n\n")

            # Results table
            f.write("RESULTS TABLE\n")
            f.write("-"*100 + "\n")
            f.write(f"{'Wine Name':<40} {'Vintage':<10} {'Item No.':<12} {'Similarity':<12} {'Status'}\n")
            f.write("-"*100 + "\n")

            for entry in wine_entries:
                wine_name = entry.get('wine_name', '')[:38]
                vintage = str(entry.get('vintage', 'N/A'))
                item_no = str(entry.get('matched_item_no', ''))
                similarity = entry.get('similarity', 0.0)
                status = '✅ MATCHED' if item_no else '❌ NOT FOUND'

                similarity_str = f"{similarity:.1%}" if similarity > 0 else "N/A"

                f.write(f"{wine_name:<40} {vintage:<10} {item_no:<12} {similarity_str:<12} {status}\n")

            f.write("\n\n")

            # Detailed results
            f.write("DETAILED RESULTS\n")
            f.write("="*100 + "\n\n")

            for i, entry in enumerate(wine_entries, 1):
                f.write(f"[{i}] {entry.get('original_text', '')}\n")
                f.write(f"    Parsed: {entry.get('wine_name', '')} | Vintage: {entry.get('vintage', 'N/A')}\n")

                if entry.get('matched_item_no'):
                    f.write(f"    ✅ MATCHED: Item No. {entry.get('matched_item_no')}\n")
                    f.write(f"       Excel Name: {entry.get('excel_wine_name', '')}\n")
                    f.write(f"       Producer: {entry.get('producer', 'N/A')}\n")
                    f.write(f"       Size: {entry.get('size', 'N/A')} cl\n")
                    f.write(f"       Similarity: {entry.get('similarity', 0):.1%}\n")
                else:
                    f.write(f"    ❌ NOT FOUND in database\n")

                f.write("\n")

        print(f"\n✅ Results report saved: {output_file}")
        return output_file

    except Exception as e:
        print(f"\n❌ Error generating report: {e}")
        return None


def main():
    """Main execution function"""
    # Parse command line arguments
    parser = argparse.ArgumentParser(description='Wine Item Number Matcher')
    parser.add_argument('--input', type=str, default=INPUT_FILE,
                       help='Input file with wine names')
    parser.add_argument('--size', type=float, default=75.0,
                       help='Preferred bottle size (75.0, 150.0, 300.0). Omit for all sizes.')
    args = parser.parse_args()

    input_file = args.input
    preferred_size = args.size

    print("="*100)
    print("Wine Item Number Matcher and Learning System")
    print("="*100 + "\n")

    # Show configuration
    print(f"📝 Input file: {input_file}")
    print(f"🍾 Bottle size filter: {preferred_size}cl (standard bottle)\n")

    # Load learning database first
    print("Step 1: Loading learning database...")
    learning_map = load_learning_database(LEARNING_DB_FILE)
    print()

    # Load primary Excel database
    print("Step 2: Loading primary database...")
    df = load_excel_database(EXCEL_FILE)
    if df is None:
        return
    print()

    # Load fallback stock database
    print("Step 3: Loading fallback stock database...")
    stock_df = load_stock_database(STOCK_FILE)
    print()

    # Parse input file
    print("Step 4: Parsing input file...")
    wines = parse_input_file(input_file)
    if wines is None or len(wines) == 0:
        print("❌ No wines found in input file")
        return
    print(f"✅ Found {len(wines)} wines to process\n")

    # Match wines
    print("Step 5: Matching wines against databases...")
    print("-"*100)

    wine_entries = []

    for wine in wines:
        wine_name = wine['wine_name']
        vintage = wine['vintage']
        original_text = wine['original_text']

        print(f"\n🔍 Processing: {original_text}")
        print(f"   Wine: {wine_name} | Vintage: {vintage if vintage else 'N/A'}")

        # Find best match (checks learning database first, then primary Excel, then stock)
        match = find_best_match(wine_name, vintage, df, learning_map=learning_map,
                               stock_df=stock_df, preferred_size=preferred_size)

        entry = {
            'original_text': original_text,
            'wine_name': wine_name,
            'vintage': vintage,
        }

        if match:
            # Show source of match with appropriate emoji
            source = match.get('source', '')
            if source == 'learning_database':
                source_label = "📚 Learning DB"
            elif source == 'learning_database_stock':
                source_label = "📚 Learning DB (Stock)"
            elif source == 'stock_fallback':
                source_label = "📦 Stock DB (Fallback)"
            else:
                source_label = "📊 Primary DB"

            print(f"   ✅ MATCHED: Item No. {match['item_no']} ({source_label})")
            print(f"      Wine: {match['wine_name']} {match['vintage']}")
            print(f"      Similarity: {match['similarity']:.1%}")

            entry.update({
                'matched_item_no': match['item_no'],
                'excel_wine_name': match['wine_name'],
                'producer': match['producer'],
                'size': match['size'],
                'similarity': match['similarity']
            })
        else:
            print(f"   ❌ NOT FOUND (no match above 60% similarity)")

        wine_entries.append(entry)

    print("\n" + "="*100)
    print("\nRESULTS TABLE")
    print("-"*100)
    print(f"{'Wine Name':<40} {'Vintage':<10} {'Item No.':<12} {'Similarity':<12} {'Status'}")
    print("-"*100)

    for entry in wine_entries:
        wine_name = entry.get('wine_name', '')[:38]
        vintage = str(entry.get('vintage', 'N/A'))
        item_no = str(entry.get('matched_item_no', ''))
        similarity = entry.get('similarity', 0.0)
        status = '✅ MATCHED' if item_no else '❌ NOT FOUND'

        similarity_str = f"{similarity:.1%}" if similarity > 0 else "N/A"

        print(f"{wine_name:<40} {vintage:<10} {item_no:<12} {similarity_str:<12} {status}")

    print("="*100)

    # Generate output report
    print("\nStep 6: Generating results report...")
    generate_output_report(wine_entries, OUTPUT_DIR)

    # Update learning database
    print("\nStep 7: Updating learning database...")
    update_learning_database(wine_entries, LEARNING_DB_FILE)

    # Create correction file for wines needing review
    print("\nStep 8: Checking for wines needing correction...")
    correction_file = create_correction_file(wine_entries, OUTPUT_DIR)

    print("\n" + "="*100)
    print("✅ PROCESSING COMPLETE")
    print("="*100)

    # Summary
    total = len(wine_entries)
    matched = sum(1 for e in wine_entries if e.get('matched_item_no'))
    needs_review = sum(1 for e in wine_entries if not e.get('matched_item_no') or e.get('similarity', 0) < 0.8)

    print(f"\n📊 Summary: {matched}/{total} wines matched ({matched/total*100:.1f}%)")
    if needs_review > 0:
        print(f"   ⚠️  {needs_review} wines need manual review")

    print(f"\n💾 Files created:")
    print(f"   - Results report: ItemNo_Results_[timestamp].txt")
    print(f"   - Learning database: {LEARNING_DB_FILE}")
    if correction_file:
        print(f"   - Corrections needed: {correction_file.name}")


if __name__ == "__main__":
    main()
