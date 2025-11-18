# Simplified Conversion Logic Using Stock Lines.xlsx

## Overview

The new conversion approach uses **Stock Lines.xlsx** to directly match items and retrieve EUR prices from the exact campaign where they were last offered.

## Database Structure

### Stock Lines.xlsx
Located: `C:\Users\Marco.Africani\OneDrive - AVU SA\AVU CPI Campaign\Puzzle_control_Reports\SOURCE_FILES\Stock Lines.xlsx`

**Key Columns:**
- **Column J**: Item No. (matching key)
- **Column AU**: OMT Last Offer Date (campaign identifier)
- **Column ?**: OMT Last Private Offer Price (CHF price in Word document)
- **Column ?**: EUR price or will be looked up in OMT Main Offer List

### OMT Main Offer List.xlsx
Located: `C:\Users\Marco.Africani\OneDrive - AVU SA\AVU CPI Campaign\Puzzle_control_Reports\SOURCE_FILES\OMT Main Offer List.xlsx`

Contains historical campaign data with EUR prices.

## New Conversion Process

### Step 1: Match Item in Stock Lines.xlsx
```python
# Extract Item No. from context (wine name, vintage)
item_no = extract_item_from_context(wine_name, vintage)

# Look up in Stock Lines.xlsx using Column J
stock_row = stock_df[stock_df['Column_J'] == item_no].iloc[0]
```

### Step 2: Get Campaign Date
```python
# Get OMT Last Offer Date from Column AU
campaign_date = stock_row['Column_AU']
```

### Step 3: Match Campaign in OMT Main Offer List
```python
# Find the exact campaign with matching date
campaign_row = omt_df[
    (omt_df['Item No.'] == item_no) &
    (omt_df['Schedule DateTime'] == campaign_date)
].iloc[0]
```

### Step 4: Return EUR Price
```python
# Get EUR price from that specific campaign
eur_price = campaign_row['Unit Price (EUR)']
```

## Benefits Over Old Approach

### Old Approach ❌
1. Complex fuzzy matching of wine names
2. Multiple filtering steps (Campaign Sub-Type, Size, Min Quantity)
3. Schedule DateTime sorting to pick latest
4. Fallback to 1.08 conversion if not found
5. Potential for wrong matches

### New Approach ✅
1. **Direct match** using Item No. from Stock Lines.xlsx
2. **Exact campaign** using OMT Last Offer Date
3. **Guaranteed correct price** from historical campaign
4. **Simpler logic** - just two lookups
5. **Faster execution**

## Implementation Status

### ✅ Completed
- Database paths updated to OneDrive SOURCE_FILES
- STOCK_FILE_PATH added to word_converter_improved.py
- Stock Lines.xlsx referenced in configuration

### ⏳ To Implement
- [ ] Load Stock Lines.xlsx structure (pending file access)
- [ ] Identify exact column names for:
  - Item No. (assumed Column J)
  - OMT Last Offer Date (assumed Column AU)
  - OMT Last Private Offer Price
  - EUR price (if present)
- [ ] Update `load_excel_data()` function to load Stock Lines
- [ ] Create `match_item_via_stock()` function
- [ ] Update `convert_chf_to_eur()` to use new logic
- [ ] Add error handling for missing items/campaigns

## Example

**Word Document Text:**
```
Château Margaux 2018 - CHF 300.00 + VAT
```

**Conversion Steps:**
1. Detect CHF 300.00
2. Extract context: Château Margaux, 2018
3. Match in Stock Lines → Item No. 12345
4. Get OMT Last Offer Date → "2025-09-15"
5. Find campaign in OMT Main Offer List with Item 12345 on 2025-09-15
6. Return EUR price → EUR 325.00

**Result:**
```
Château Margaux 2018 - EUR 325.00 + VAT
```

## Notes

- Stock Lines.xlsx file was locked during initial check (likely open in Excel)
- Need to verify exact column names and structure
- This approach assumes Stock Lines.xlsx is always up-to-date with last offers
- If item not in Stock Lines, fallback to old matching logic (optional)

---

Generated: 2025-01-10
Status: Documentation complete, implementation pending file access
