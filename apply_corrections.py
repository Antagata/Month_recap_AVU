#!/usr/bin/env python3
"""
Apply Manual Corrections to Learning Database
==============================================

This script reads a corrections file created by wine_item_matcher.py
and applies the manual corrections to the learning database.

Usage:
    python apply_corrections.py CORRECTIONS_[timestamp].txt
"""

import sys
import io
from datetime import datetime
from pathlib import Path

# Fix encoding for Windows console
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Configuration
LEARNING_DB_FILE = r"C:\Users\Marco.Africani\Desktop\Month recap\wine_names_learning_db.txt"


def parse_corrections_file(corrections_file):
    """
    Parse the corrections file and extract wine entries with corrected Item Numbers.

    Format expected:
    Wine Name | Vintage | Item No. | Notes
    """
    corrections = []

    try:
        with open(corrections_file, 'r', encoding='utf-8') as f:
            lines = f.readlines()

        for line in lines:
            line = line.strip()

            # Skip empty lines, comments, headers, and instruction lines
            if not line or line.startswith('#') or line.startswith('=') or line.startswith('-'):
                continue
            if line.startswith('[') and ']' in line:
                # Skip entry headers like "[1] Original wine name"
                continue
            if 'INSTRUCTIONS' in line or 'Format:' in line or 'Generated:' in line:
                continue
            if 'WINE CORRECTIONS FILE' in line or 'PROCESSING COMPLETE' in line:
                continue

            # Parse wine entry
            # Format: Wine Name | Vintage | Item No. | Notes
            parts = line.split(' | ')

            if len(parts) >= 3:
                wine_name = parts[0].strip()
                vintage = parts[1].strip()
                item_no = parts[2].strip()

                # Skip entries that haven't been corrected
                if item_no == 'YOUR_ITEM_NO_HERE' or not item_no:
                    print(f"⏭️  Skipping uncorrected entry: {wine_name} {vintage}")
                    continue

                # Validate Item No is numeric
                try:
                    int(item_no)
                except ValueError:
                    print(f"⚠️  Invalid Item No '{item_no}' for {wine_name} {vintage} - skipping")
                    continue

                corrections.append({
                    'wine_name': wine_name,
                    'vintage': vintage,
                    'item_no': item_no
                })

    except FileNotFoundError:
        print(f"❌ Error: Corrections file not found: {corrections_file}")
        return None
    except Exception as e:
        print(f"❌ Error reading corrections file: {e}")
        return None

    return corrections


def apply_corrections_to_learning_db(corrections, learning_db_path):
    """
    Apply corrections to the learning database.
    Only adds entries if they don't already exist.
    """
    # Load existing database keys
    existing_keys = set()

    if Path(learning_db_path).exists():
        try:
            with open(learning_db_path, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith('#'):
                        # Extract key: wine_name|vintage|item_no
                        parts = line.split(' | ')
                        if len(parts) >= 3:
                            key = f"{parts[0]}|{parts[1]}|{parts[2]}"
                            existing_keys.add(key)
        except Exception as e:
            print(f"⚠️  Warning: Could not read learning database: {e}")

    # Apply corrections
    new_entries = []
    duplicate_count = 0

    for correction in corrections:
        wine_name = correction['wine_name']
        vintage = correction['vintage']
        item_no = correction['item_no']

        # Create unique key
        key = f"{wine_name}|{vintage}|{item_no}"

        # Only add if not already in database
        if key not in existing_keys:
            # Format: Wine Name | Vintage | Item No. | Timestamp
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            entry_line = f"{wine_name} | {vintage} | {item_no} | {timestamp} (manual correction)"
            new_entries.append(entry_line)
            existing_keys.add(key)
            print(f"✅ Adding correction: {wine_name} {vintage} → Item No. {item_no}")
        else:
            duplicate_count += 1
            print(f"⏭️  Skipping duplicate: {wine_name} {vintage} → Item No. {item_no}")

    # Write new entries to database
    if new_entries:
        try:
            with open(learning_db_path, 'a', encoding='utf-8') as f:
                for entry in new_entries:
                    f.write(entry + "\n")

            print(f"\n✅ Successfully added {len(new_entries)} corrections to learning database")
            if duplicate_count > 0:
                print(f"   ⏭️  Skipped {duplicate_count} duplicates")
            print(f"   Total unique entries in database: {len(existing_keys)}")

        except Exception as e:
            print(f"\n❌ Error writing to learning database: {e}")
            return False
    else:
        print(f"\n✅ No new corrections to add")
        if duplicate_count > 0:
            print(f"   All {duplicate_count} entries were already in database")

    return True


def main():
    """Main execution function"""
    print("="*100)
    print("Apply Manual Corrections to Learning Database")
    print("="*100 + "\n")

    # Check command line arguments
    if len(sys.argv) < 2:
        print("❌ Error: No corrections file specified")
        print("\nUsage:")
        print("    python apply_corrections.py CORRECTIONS_[timestamp].txt")
        print("\nExample:")
        print("    python apply_corrections.py CORRECTIONS_NEEDED_20251104_195500.txt")
        return

    corrections_file = sys.argv[1]

    # Check if file exists
    if not Path(corrections_file).exists():
        print(f"❌ Error: File not found: {corrections_file}")
        return

    print(f"📄 Corrections file: {corrections_file}")
    print(f"📚 Learning database: {LEARNING_DB_FILE}\n")

    # Parse corrections file
    print("Step 1: Parsing corrections file...")
    corrections = parse_corrections_file(corrections_file)

    if corrections is None:
        return

    if len(corrections) == 0:
        print("⚠️  No valid corrections found in file")
        print("   Make sure you replaced 'YOUR_ITEM_NO_HERE' with actual Item Numbers")
        return

    print(f"✅ Found {len(corrections)} valid corrections\n")

    # Apply corrections
    print("Step 2: Applying corrections to learning database...")
    success = apply_corrections_to_learning_db(corrections, LEARNING_DB_FILE)

    print("\n" + "="*100)
    if success:
        print("✅ CORRECTIONS APPLIED SUCCESSFULLY")
    else:
        print("❌ CORRECTIONS FAILED")
    print("="*100)


if __name__ == "__main__":
    main()
