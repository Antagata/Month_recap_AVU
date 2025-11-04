# Wine Item Number Matcher - User Guide

## Overview
The Wine Item Number Matcher is a standalone tool that matches wine names + vintages to Item Numbers from your Excel database. It also builds a **learning database** that grows over time to improve wine name recognition.

## Quick Start

### 1. Edit the Input File
Open `ItemNoGenerator.txt` and add your wines, one per line:

```
Château Lafite Rothschild 2005
Château Lafleur 2022
Barolo Monfortino Riserva 2004
Pétrus 2010
```

**Supported Formats:**
- `Wine Name Vintage` (e.g., "Lafite 2005")
- `Wine Name, Vintage` (e.g., "Lafite, 2005")
- `Wine Name | Vintage` (e.g., "Lafite | 2005")

### 2. Run the Matcher
```bash
python wine_item_matcher.py
```

### 3. Check the Results
The script creates two files:

#### A. Results Report (`ItemNo_Results_[timestamp].txt`)
Contains:
- Summary statistics (matched vs. not matched)
- Results table with Wine Name, Vintage, Item No., Similarity
- Detailed results with Excel name, producer, size

**Example Table:**
```
Wine Name                      Vintage    Item No.     Similarity   Status
--------------------------------------------------------------------------------
Barolo Monfortino Riserva      2004       15971        100.0%       ✅ MATCHED
Lafleur                        2022       63011        100.0%       ✅ MATCHED
Pavie                          2021                    N/A          ❌ NOT FOUND
```

#### B. Learning Database (`wine_names_learning_db.txt`)
A cumulative database that records ALL wines processed over time:

```
Barolo Monfortino Riserva | 2004 | 15971 | 2025-11-04 19:45:55
Lafleur | 2022 | 63011 | 2025-11-04 19:45:55
Pavie | 2021 | NOT_FOUND | 2025-11-04 19:45:55
```

**This database:**
- ✅ Grows with each run (new entries added automatically)
- ✅ Never deletes old entries
- ✅ Tracks when each wine was processed
- ✅ Records both successful matches AND failures

---

## How It Works

### Matching Algorithm
1. **Exact Vintage Match**: First filters by vintage (if provided)
2. **Fuzzy Name Matching**: Uses normalized wine names to handle:
   - Different spellings (Château vs Chateau)
   - With/without prefixes (Château Lafite vs Lafite)
   - Partial names (Lafite matches Lafite Rothschild)
3. **Similarity Score**: Only matches above 60% similarity
4. **Best Match**: Returns highest similarity score

### Normalization Process
The system normalizes wine names by:
- Converting to lowercase
- Removing "Château", "Domaine", "Ch", "Ch."
- Removing special characters
- Removing extra whitespace

**Example:**
```
Input:  "Château Lafite-Rothschild 2005"
Normalized: "lafite rothschild"
Matches: "Lafite Rothschild" (100%)
```

---

## Use Cases

### Use Case 1: Generate Item Numbers for New Wines
**Scenario:** You have a list of wines and need Item Numbers quickly

1. Add wines to `ItemNoGenerator.txt`
2. Run `python wine_item_matcher.py`
3. Check results report for Item Numbers

### Use Case 2: Build Training Database
**Scenario:** Train the system to recognize wine names better over time

1. Process wines regularly (weekly/monthly)
2. Learning database grows automatically
3. Future: Main converter can use this database for better matching

### Use Case 3: Verify Wine Names
**Scenario:** Check if wines exist in your Excel database

1. Add suspected wine names to input file
2. Run matcher
3. Check "NOT FOUND" entries - these wines don't exist in database

---

## Future Integration

### How This Helps the Main Converter

The learning database (`wine_names_learning_db.txt`) can be used to:

1. **Pre-match Common Wines**: Before fuzzy matching, check learning database
2. **Learn Name Variations**: System learns that "Lafite" = "Lafite Rothschild"
3. **Faster Processing**: Skip Excel lookup for known wines
4. **Better Accuracy**: Use historical data to improve matching

**Future Enhancement:**
```python
# In word_converter_improved.py, add:
def load_learning_database():
    """Load pre-learned wine name mappings"""
    # Parse wine_names_learning_db.txt
    # Create quick lookup: wine_name+vintage -> item_no
    # Use this BEFORE fuzzy matching for faster results
```

---

## Configuration

### Change Input/Output Locations
Edit these variables in `wine_item_matcher.py`:

```python
EXCEL_FILE = r"C:\Users\...\Conversion_month.xlsx"
INPUT_FILE = r"C:\Users\...\ItemNoGenerator.txt"
OUTPUT_DIR = r"C:\Users\...\Month recap"
LEARNING_DB_FILE = r"C:\Users\...\wine_names_learning_db.txt"
```

### Adjust Matching Threshold
To make matching more/less strict:

```python
def find_best_match(wine_name, vintage, df, threshold=0.6):  # Change 0.6
```

- **0.6** (60%) = Default (balanced)
- **0.8** (80%) = Stricter (fewer but more accurate matches)
- **0.4** (40%) = Looser (more matches but less accurate)

---

## Tips & Best Practices

### ✅ DO:
- Add wines gradually and check results
- Review "NOT FOUND" entries - they might have typos
- Keep learning database backed up (it's valuable over time)
- Run regularly to build comprehensive database

### ❌ DON'T:
- Delete `wine_names_learning_db.txt` - it's your training data!
- Add wines without vintages (matching will be less accurate)
- Ignore low similarity scores (< 80%) - verify these manually

---

## Troubleshooting

### "NOT FOUND" for Known Wines
**Possible reasons:**
1. Wine name spelling is very different from Excel
2. Vintage doesn't exist in Excel
3. Wine genuinely not in database

**Solutions:**
- Try simpler name (e.g., "Lafite 2005" instead of "Château Lafite Rothschild Premier Grand Cru Classé 2005")
- Check Excel for exact spelling
- Try without "Château" or "Domaine"

### Low Similarity Scores (< 80%)
**What it means:** Match found but names are quite different

**Actions:**
- Check if Item Number is correct
- Verify in Excel
- Consider adding alias to learning database

### Multiple Matches Possible
**Current behavior:** Returns best match (highest similarity)

**If wrong match:**
- Be more specific in input (add producer name)
- Check Excel - might have duplicate wine names

---

## Files Generated

| File | Purpose | Keep? |
|------|---------|-------|
| `ItemNo_Results_[timestamp].txt` | Current run results | Archive old ones |
| `wine_names_learning_db.txt` | Cumulative database | ✅ KEEP ALWAYS |
| `.gitignore` (updated) | Ignores temporary files | ✅ Keep |

---

## Example Workflow

### Weekly Processing:
```bash
# Monday: Add wines from clients
echo "Lafite 2010" >> ItemNoGenerator.txt
echo "Margaux 2015" >> ItemNoGenerator.txt

# Tuesday: Run matcher
python wine_item_matcher.py

# Wednesday: Review results, use Item Numbers
# Learning database grows automatically!

# Next week: Repeat with new wines
# Database keeps growing, system gets smarter!
```

---

## Support

### Common Questions

**Q: Can I process the same wine twice?**
A: Yes! The learning database tracks all entries with timestamps.

**Q: What if vintage is missing?**
A: System will try to match by name only (less accurate).

**Q: Can I edit the learning database manually?**
A: Yes, but be careful with the format: `Wine | Vintage | ItemNo | Timestamp`

**Q: How big can the learning database get?**
A: Unlimited! It's a text file that grows over time. 10,000 entries ≈ 1MB.

---

## Version Info

- **Version:** 1.0
- **Created:** 2025-11-04
- **Requires:** Python 3.7+, pandas, openpyxl
- **Compatible with:** word_converter_improved.py v2.0+
