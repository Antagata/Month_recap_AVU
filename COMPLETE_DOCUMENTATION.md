# AVU Echo Spinner - Complete Documentation

## Overview

AVU Echo Spinner is a desktop application for wine document processing with two main functions:
1. **SPIN Button** - CHF to EUR price converter with wine recognition
2. **MATCH Button** - Wine name to Item Number matcher

---

## SPIN BUTTON - CHF→EUR Converter & Wine Recognition

### Purpose
Processes wine offer documents to:
- Recognize wines from EUR prices in text
- Match wines to Stock Lines database
- Generate filtered Stock Lines Excel with only recognized items
- Create detailed recognition reports

### Input File
**Location**: `Inputs/Multi.txt`

**Format**: Text file containing wine descriptions with EUR prices

**Example**:
```
- Evangile 2015: "Rich, juicy, and spectacular wine." EUR 202.00 + VAT
- Palmer 2015: "Exceptionally complex wine." EUR 287.00 + VAT
- Marojallia 2015: Great wine from Margaux. EUR 30.00 + VAT
```

### How It Works

#### Step 1: Load Databases
- Loads **Stock Lines.xlsx** from OneDrive SOURCE_FILES
- Loads **OMT Main Offer List.xlsx** from OneDrive SOURCE_FILES

#### Step 2: Extract Wine Information
For each EUR price found in the text:
1. Extracts context (150 characters before the price)
2. Looks for vintage (4-digit year like 2015)
3. Extracts wine name using patterns:
   - Pattern 1: `Wine Name VINTAGE:` (e.g., "Evangile 2015:")
   - Pattern 2: `Wine Name VINTAGE` without colon

#### Step 3: Match to Database
Matching algorithm (in order of preference):
1. **Exact EUR price match** - Finds all items with the same EUR price (rounded)
2. **Wine name similarity** - Scores matches based on:
   - Direct substring match: +10 points
   - Partial word match (words >3 chars): +3 points per word
3. **Vintage match** - If vintage year matches: +10 points
4. **Best match selection** - Chooses highest scoring match
5. **Stock Lines lookup** - Finds the item in Stock Lines using Item No.

#### Step 4: Generate Outputs

**1. Multi_converted_YYYYMMDD_HHMMSS.txt**
- Location: `Outputs/`
- Content: Copy of input text (prices already in EUR)

**2. Lines.xlsx**
- Location: `Outputs/`
- Content: Recognized wines with details
- Columns:
  - Wine Name (from OMT database)
  - Vintage (from OMT database)
  - Size (bottle size in cl)
  - Producer (producer name)
  - EUR Price
  - Item No.

**3. Stock_Lines_Filtered_YYYYMMDD_HHMMSS.xlsx**
- Location: `Outputs/`
- Content: Complete Stock Lines data for ONLY recognized items
- Purpose: Quick verification of recognized wines
- Usage:
  - Review all matched items
  - Remove false positives (incorrectly identified wines)
  - Verify all intended wines were detected

**4. Recognition_Report_YYYYMMDD_HHMMSS.txt**
- Location: `Outputs/Detailed match results/`
- Content: Detailed matching statistics
- Sections:
  - Total prices found
  - Successfully matched count
  - Not matched count
  - List of all recognized wines with details

### Performance Metrics

**Typical Results**:
- **Success Rate**: 77% (17 out of 22 wines matched)
- **Processing Time**: 5-10 seconds
- **Accuracy**: Depends on EUR price uniqueness and database completeness

**Why Some Wines Don't Match**:
1. **Wine not in database** - EUR price doesn't exist in OMT Main Offer List
2. **Multiple wines same price** - Ambiguous match, algorithm picks best candidate
3. **Price not in Stock Lines** - Item exists in OMT but not in current Stock Lines inventory

### Usage Instructions

1. **Prepare Input**:
   - Edit `Inputs/Multi.txt` with your wine offer text
   - Ensure EUR prices are included (format: "EUR XX.00")
   - Include wine names and vintages when possible

2. **Run Converter**:
   - Launch application: Run `Launch AVU Echo Spinner.bat`
   - Click **"🔄 SPIN!"** button
   - Wait for processing (5-10 seconds)

3. **Review Results**:
   - Check `Outputs/Lines.xlsx` for quick summary
   - Open `Outputs/Stock_Lines_Filtered_XXX.xlsx` to verify recognized items
   - Review `Outputs/Detailed match results/Recognition_Report_XXX.txt` for details

4. **Verify Accuracy**:
   - Open Stock_Lines_Filtered Excel
   - Check if all rows are correct wines
   - Remove any false positives
   - Note any missing wines in the Recognition Report

### Direct Paragraph Conversion (Alternative Method)

**Purpose**: Quick conversion of small text snippets

**How to Use**:
1. Type or paste text directly into the paragraph input box (GUI)
2. Click **"🔄 SPIN!"** button
3. View converted output in results panel

**Output**:
- Saved to `Outputs/Converted_Paragraph_YYYYMMDD_HHMMSS.txt`
- Shows original and converted text side-by-side

---

## MATCH BUTTON - Wine Name to Item Number Matcher

### Purpose
Matches wine names to Item Numbers using fuzzy matching and learning database

### Input File
**Location**: `Inputs/ItemNoGenerator.txt`

**Format**: Text file with wine names (one per line or comma-separated)

**Example**:
```
Château Margaux 2015
Petrus 2018, Lafite 2016
Domaine de la Romanée-Conti
```

### How It Works

#### Step 1: Load Databases
- Loads **OMT Main Offer List.xlsx** for Item Numbers
- Loads **Stock Lines.xlsx** for current inventory
- Loads **wine_names_learning_db.txt** for previous corrections

#### Step 2: Parse Wine Names
Extracts from input:
- Wine name (with château/domaine prefix handling)
- Vintage year (4-digit number)
- Variations and alternate spellings

#### Step 3: Fuzzy Matching
For each wine name:
1. **Exact match** - Looks for exact name in databases
2. **Fuzzy match** - Uses similarity scoring:
   - Levenshtein distance
   - Partial string matching
   - Word-by-word comparison
3. **Learning database check** - Uses previous manual corrections
4. **Size filtering** - Prioritizes standard 75cl bottles
5. **Latest campaign** - Picks most recent offer if multiple matches

#### Step 4: Generate Outputs

**1. ItemNo_Results_YYYYMMDD_HHMMSS.txt**
- Location: `Outputs/Detailed match results/`
- Content: Detailed matching results with:
  - Search query
  - Match status (Found/Not Found/Multiple)
  - Item Number
  - Full wine details from database
  - Similarity score

**2. Learning Database Display**
- Shows in GUI results panel
- Lists all previous wine name corrections
- Format: `Original Name → Corrected Name`

### Learning Database

**Purpose**: Improve future matching by storing corrections

**How It Works**:
1. When you correct a wine name match manually
2. System stores: Original Name → Correct Name
3. Future searches automatically use corrections
4. Builds over time to improve accuracy

**File**: `wine_names_learning_db.txt`

**Format**:
```
incorrect name|correct name|timestamp
Petru|Petrus|2025-01-18_14:30:00
Margau|Margaux|2025-01-18_14:35:00
```

### Usage Instructions

1. **Prepare Input**:
   - Edit `Inputs/ItemNoGenerator.txt`
   - Add wine names (one per line or comma-separated)
   - Include vintages when known

2. **Run Matcher**:
   - Launch application
   - Click **"🔍 MATCH"** button
   - Wait for processing

3. **Review Results**:
   - Check GUI results panel for summary
   - Open `Outputs/Detailed match results/ItemNo_Results_XXX.txt` for full details
   - Note similarity scores (>70% is usually accurate)

4. **Improve Learning Database**:
   - For incorrect matches, note the correct wine name
   - Manually add to learning database if needed
   - Future searches will use the correction

---

## File Structure

```
C:\Users\Marco.Africani\Desktop\Month recap\
├── Inputs/
│   ├── Multi.txt                    # Input for SPIN button
│   ├── ItemNoGenerator.txt          # Input for MATCH button
│   └── Example style.txt            # Style template (future AI use)
│
├── Outputs/
│   ├── Lines.xlsx                   # SPIN: Recognized wines summary
│   ├── Multi_converted_*.txt        # SPIN: Converted text output
│   ├── Stock_Lines_Filtered_*.xlsx  # SPIN: Filtered Stock Lines
│   └── Detailed match results/
│       ├── Recognition_Report_*.txt # SPIN: Detailed recognition report
│       └── ItemNo_Results_*.txt     # MATCH: Detailed matching results
│
├── Database/ (OneDrive)
│   ├── Stock Lines.xlsx
│   └── OMT Main Offer List.xlsx
│
├── wine_names_learning_db.txt       # Learning database
├── avu_echo_spinner.py              # GUI application
├── txt_converter.py                 # SPIN engine
└── wine_item_matcher.py             # MATCH engine
```

---

## Database Requirements

### Stock Lines.xlsx
**Location**: OneDrive SOURCE_FILES folder

**Required Columns**:
- Column A: `No.` (Item Number)
- Column AQ: `OMT Last Private Offer Price` (CHF price)
- Column AU: `OMT Last Offer Date` (Campaign date)
- Additional columns: Wine Name, Size, Vintage Code, Producer Name, etc.

### OMT Main Offer List.xlsx
**Location**: OneDrive SOURCE_FILES folder

**Required Columns**:
- `Item No.` (matches Stock Lines No.)
- `Unit Price` (CHF price)
- `Unit Price (EUR)` (EUR price)
- `Wine Name`
- `Vintage Code`
- `Producer Name`
- `Size`
- `Minimum Quantity`
- `Schedule DateTime` (Campaign date/time)

---

## Troubleshooting

### SPIN Button Issues

**Problem**: No wines matched (0/22)
- **Cause**: Database files not found or EUR prices don't exist in database
- **Solution**:
  1. Check OneDrive path is correct
  2. Verify EUR prices in Multi.txt exist in OMT database
  3. Check Recognition Report for specific errors

**Problem**: Low match rate (<50%)
- **Cause**: EUR prices are too common or wine names poorly formatted
- **Solution**:
  1. Add vintage years to text (improves matching)
  2. Use standard wine name format: "Wine Name VINTAGE:"
  3. Check if wines exist in database with those exact EUR prices

**Problem**: Wrong wines matched
- **Cause**: Multiple wines with same EUR price
- **Solution**:
  1. Review Stock_Lines_Filtered Excel
  2. Remove incorrect rows manually
  3. Check Recognition Report similarity scores

### MATCH Button Issues

**Problem**: Wine not found
- **Cause**: Name misspelled or not in database
- **Solution**:
  1. Check spelling against OMT database
  2. Try alternate spellings
  3. Add to learning database manually

**Problem**: Multiple matches
- **Cause**: Ambiguous wine name (e.g., "Margaux" matches many wines)
- **Solution**:
  1. Add more specific details (vintage, producer)
  2. Review all matches in results file
  3. Pick correct one based on context

---

## Advanced Features

### Price Rounding (SPIN)
- Prices above CHF 300 are rounded to nearest 5 or 0
- Example: EUR 1162 → EUR 1165, EUR 1163 → EUR 1165

### 36-Bottle Pricing (SPIN)
- Detects "36x" or "36 bottles" pattern in text
- Matches with `Minimum Quantity = 36` in database
- Separate pricing from regular bottles

### Campaign Date Matching (SPIN)
- Matches items to specific campaign dates
- Stock Lines "OMT Last Offer Date" → OMT "Schedule DateTime"
- Ensures historical accuracy

---

## Tips for Best Results

### SPIN Button
1. **Include vintages** - "Wine Name 2015" improves matching
2. **Use standard format** - "Wine Name VINTAGE: description EUR price"
3. **Check EUR prices** - Verify prices exist in database before running
4. **Review filtered Excel** - Always verify recognized items
5. **Update databases** - Keep Stock Lines and OMT current

### MATCH Button
1. **Be specific** - "Château Margaux 2015" better than "Margaux"
2. **Use learning database** - Builds accuracy over time
3. **Check similarity scores** - >70% is usually correct
4. **Include producer** - Helps with common wine names
5. **Review all matches** - Don't rely only on first match

---

## Version History

**Current Version**: 2.0

**Recent Updates**:
- Added txt file support (replacing Word documents)
- Implemented Stock Lines filtering
- Enhanced wine recognition algorithm
- Added Lines.xlsx generation
- Improved matching success rate to 77%
- Updated GUI labels and defaults

---

## Support

For issues or questions:
1. Check this documentation
2. Review Recognition Reports for specific errors
3. Verify database files are accessible
4. Check file paths in configuration

---

Generated: November 2025
