#!/usr/bin/env python3
"""
Integrated Wine Converter
=========================
1. Extract wine names + vintages from Multi.txt
2. Write to ItemNoGenerator.txt
3. Run wine_item_matcher.py to get Item Numbers
4. Use matched Item Numbers for conversion
5. Generate Lines.xlsx in correct order
"""

import re
import subprocess
import sys
from pathlib import Path
from datetime import datetime

# Configuration
BASE_DIR = r"C:\Users\Marco.Africani\Desktop\Month recap"
INPUT_FILE = rf"{BASE_DIR}\Inputs\Multi.txt"
ITEMNO_GEN_FILE = rf"{BASE_DIR}\Inputs\ItemNoGenerator.txt"
LEARNING_DB = rf"{BASE_DIR}\wine_names_learning_db.txt"


def extract_wine_names_from_multi(file_path):
    """Extract wine names and vintages from Multi.txt"""
    wines = []
    seen_wines = set()  # Track duplicates

    with open(file_path, 'r', encoding='utf-8') as f:
        text = f.read()

    # Pattern 1: Wine Name VINTAGE : description
    # Pattern 2: Wine Name VINTAGE – something : description
    # Pattern 3: Wine Name (no vintage) : description (for non-vintage wines)

    # Try to match wines with vintage first (with optional separator before colon)
    pattern_with_vintage = r'([A-Z\u00C0-\u017F][^\n\d]*?)\s+(19\d{2}|20\d{2})\s*[–—-]*\s*:'

    for match in re.finditer(pattern_with_vintage, text, re.MULTILINE):
        wine_name = match.group(1).strip()
        vintage = match.group(2)

        # Clean up emojis and special chars
        wine_name = re.sub(r'[✨💎💼🍷🏆⭐🎯]', '', wine_name).strip()

        # Skip duplicates (same wine + vintage)
        wine_key = f"{wine_name}|{vintage}"
        if wine_key in seen_wines:
            continue
        seen_wines.add(wine_key)

        wines.append((wine_name, vintage))

    # Now try to match non-vintage wines (no year before colon)
    # But skip lines that already matched above
    pattern_non_vintage = r'([A-Z\u00C0-\u017F][^\n:]*?)\s*:\s*[a-z]'

    for match in re.finditer(pattern_non_vintage, text, re.MULTILINE):
        wine_name = match.group(1).strip()

        # Skip if it has a 4-digit year (already matched above)
        if re.search(r'(19\d{2}|20\d{2})', wine_name):
            continue

        # Clean up emojis, special chars, and edition numbers
        wine_name = re.sub(r'[✨💎💼🍷🏆⭐🎯]', '', wine_name).strip()
        wine_name = re.sub(r'\s+\d+(?:ème|eme|th|nd|rd|st)\s+(?:Édition|Edition)', '', wine_name, flags=re.IGNORECASE).strip()

        # Skip if wine_name is too short or generic
        if len(wine_name) < 5 or wine_name.lower() in ['top wines', 'top selling', 'more expensive']:
            continue

        # Skip duplicates
        wine_key = f"{wine_name}|NV"
        if wine_key in seen_wines:
            continue
        seen_wines.add(wine_key)

        wines.append((wine_name, 'NV'))

    return wines


def write_to_itemno_generator(wines, output_file):
    """Write wine names and vintages to ItemNoGenerator.txt"""
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write("# Wine List for Item Number Matching\n")
        f.write(f"# Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write("#\n")
        f.write("# Format: Wine Name | Vintage\n")
        f.write("# One wine per line\n")
        f.write("#\n\n")

        for wine_name, vintage in wines:
            f.write(f"{wine_name} | {vintage}\n")

    return len(wines)


def run_wine_matcher(size="75.0"):
    """Run wine_item_matcher.py to match Item Numbers"""
    print("\n" + "="*80)
    print("STEP 2: Running Wine Item Matcher...")
    print("="*80)

    cmd = [sys.executable, "wine_item_matcher.py", "--size", size]

    result = subprocess.run(
        cmd,
        cwd=Path(__file__).parent,
        capture_output=True,
        text=True,
        encoding='utf-8',
        errors='replace'
    )

    # Print output with error handling for emojis
    try:
        print(result.stdout)
        if result.stderr:
            print(result.stderr)
    except UnicodeEncodeError:
        # If Windows console can't handle UTF-8, replace problematic chars
        print(result.stdout.encode('ascii', errors='replace').decode('ascii'))
        if result.stderr:
            print(result.stderr.encode('ascii', errors='replace').decode('ascii'))

    return result.returncode == 0


def check_matching_quality(learning_db_file):
    """Check if matching quality is >= 70%"""
    print("\n" + "="*80)
    print("STEP 3: Checking Matching Quality...")
    print("="*80)

    if not Path(learning_db_file).exists():
        print("[WARN] Learning database not found. Cannot check quality.")
        return True  # Allow to proceed

    with open(learning_db_file, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    total = 0
    matched = 0
    not_found = 0

    for line in lines:
        line = line.strip()
        if line and not line.startswith('#'):
            parts = line.split(' | ')
            if len(parts) >= 3:
                total += 1
                item_no = parts[2].split()[0]

                if item_no == 'NOT_FOUND':
                    not_found += 1
                else:
                    matched += 1

    if total == 0:
        print("[WARN] No entries found in learning database.")
        return True

    success_rate = (matched / total) * 100

    print(f"\nMatching Statistics:")
    print(f"  Total wines: {total}")
    print(f"  Successfully matched: {matched}")
    print(f"  NOT FOUND: {not_found}")
    print(f"  Success rate: {success_rate:.1f}%")

    if success_rate < 70:
        print(f"\n[ERROR] Success rate ({success_rate:.1f}%) is below 70%!")
        print(f"[ERROR] Please review the ItemNo_Results_*.txt file and apply corrections.")
        print(f"[ERROR] Use 'Apply Corrections' button in the GUI to update the learning database.")
        return False
    else:
        print(f"\n[OK] Success rate ({success_rate:.1f}%) is >= 70%. Proceeding with conversion...")
        return True


def main():
    """Main workflow"""
    print("="*80)
    print("INTEGRATED WINE CONVERTER")
    print("="*80)
    print()

    # STEP 1: Extract wine names from Multi.txt
    print("STEP 1: Extracting wine names and vintages from Multi.txt...")
    try:
        wines = extract_wine_names_from_multi(INPUT_FILE)
        print(f"[OK] Extracted {len(wines)} unique wines")

        # Show first 10 wines
        print(f"\nFirst 10 wines:")
        for i, (name, vintage) in enumerate(wines[:10], 1):
            print(f"  {i}. {name} {vintage}")

        if len(wines) > 10:
            print(f"  ... and {len(wines) - 10} more")

    except Exception as e:
        print(f"[ERROR] Failed to extract wines: {e}")
        return 1

    # Write to ItemNoGenerator.txt
    print(f"\nWriting to {Path(ITEMNO_GEN_FILE).name}...")
    try:
        count = write_to_itemno_generator(wines, ITEMNO_GEN_FILE)
        print(f"[OK] Wrote {count} wines to ItemNoGenerator.txt")
    except Exception as e:
        print(f"[ERROR] Failed to write file: {e}")
        return 1

    # STEP 2: Run wine_item_matcher.py
    try:
        success = run_wine_matcher(size="75.0")
        if not success:
            print("[ERROR] Wine matching failed")
            return 1
    except Exception as e:
        print(f"[ERROR] Failed to run wine matcher: {e}")
        return 1

    # STEP 3: Check matching quality
    if not check_matching_quality(LEARNING_DB):
        print("\n" + "="*80)
        print("STOPPING HERE - MATCHING QUALITY TOO LOW")
        print("="*80)
        print("\nPlease:")
        print("1. Review the ItemNo_Results_*.txt file")
        print("2. Correct any wrong Item Numbers")
        print("3. Click 'Apply Corrections' in the GUI")
        print("4. Run this script again")
        return 2  # Special return code for low quality

    # STEP 4: Wine matching completed - corrections will be shown in GUI
    print("\n" + "="*80)
    print("WINE MATCHING COMPLETED!")
    print("="*80)
    print("\nNext steps:")
    print("  1. Review corrections in the popup window")
    print("  2. Edit vintages and Item Numbers as needed")
    print("  3. Delete any unwanted wines")
    print("  4. Click 'Apply All Corrections'")
    print("  5. Then click 'Generate Lines.xlsx' button in GUI")
    print("\nNOTE: Lines.xlsx will be generated AFTER you apply corrections!")
    print("      This ensures all your corrections are included in the final output.")

    return 0  # Success - corrections file was created in txt_converter


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
