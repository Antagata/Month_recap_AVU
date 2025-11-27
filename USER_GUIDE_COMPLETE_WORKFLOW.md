# AVU Wine Price Converter - Complete User Guide

## Overview
This application converts wine lists from CHF (Swiss Francs) to EUR (Euros), matching wine names to your database and generating an Excel file with the correct order and pricing.

---

## Step-by-Step Workflow

### Step 1: Prepare Your Input File

**File Location**: `C:\Users\Marco.Africani\Desktop\Month recap\Inputs\Multi.txt`

**What to Include**:
- Wine names with vintages
- CHF prices (for 1 bottle or 36 bottles)
- Market price references (will be converted by 1.08)
- Write naturally in essay/paragraph format - the app extracts the data automatically

**Example Format**:
```
Magnum Cristal Rosé 2014: This exceptional champagne is available at CHF 900.00

Bollinger RD 2008: A prestigious vintage at CHF 230.00 per bottle,
or CHF 220.00 for 36 bottles

Château Haut-Marbuzet 2019: Market price CHF 35.00 (converted automatically)
```

**Important Notes**:
- Wine names can include producer names (e.g., "Krug Rosé 29ème Édition")
- Champagne wines may not have vintages (NV = Non-Vintage)
- The word "Magnum" in the text tells the app it's a 150cl bottle (not standard 75cl)
- Order matters! The output will match your input order exactly

---

### Step 2: Launch the Application

1. Open **AVU Echo Spinner** (the GUI application)
2. The default input file should already be set to `Multi.txt`
3. If not, click **"Browse Document"** and select your Multi.txt file

---

### Step 3: Run the Conversion

1. Click the **"SPIN (CHF → EUR)"** button (green button)
2. The application will:
   - Extract wine names and vintages from Multi.txt
   - Match them against the database using:
     - Learning database (wines you've corrected before)
     - OMT Main Offer List (current prices)
     - Stock Lines database (wine details)
   - Apply filters:
     - Only prices where "Competitor Code" is empty
     - Only campaigns with "Campaign Status" = "Sent"
     - Correct bottle size (75cl standard, 150cl Magnum)
   - Convert CHF prices to EUR
   - Preserve the exact order from your Multi.txt file

**Wait Time**: Usually 30-60 seconds depending on file size

---

### Step 4: Review Corrections (If Needed)

#### What Happens:
If the application can't find a perfect match for some wines, it will:
1. **Automatically show the "Wine Corrections - Interactive" panel** at the bottom
2. Display a table with wines that need your review

#### What You'll See in the Corrections Table:

| Wine Name | Vintage | Price | Suggested Item No. | **Correct Item No.** | Reason |
|-----------|---------|-------|--------------------|---------------------|---------|
| Krug Rosé 29ème Édition | NV | CHF 300.0 | 65274 | *(editable field)* | Price-only match |

**Columns Explained**:
- **Wine Name**: The wine from your Multi.txt
- **Vintage**: Year (or NV for non-vintage)
- **Price**: CHF price from your file
- **Suggested Item No.**: What the system thinks is correct (may be right or wrong)
- **Correct Item No.**: **This is where YOU type the correct Item Number**
- **Reason**: Why it needs review (e.g., "Price-only match" means the price matched but not the wine name)

#### How to Correct:

1. **Review each wine** in the table
2. **Check if the suggested Item No. looks correct**
   - If correct: Leave it as-is (it's pre-filled)
   - If wrong: Type the correct Item Number in the "Correct Item No." field
   - If unsure: Look up the wine in your database
3. **Click "Apply All Corrections"** button

#### What Happens Next:
- The system validates all Item Numbers (must be numeric)
- Saves corrections to the **Learning Database** (`wine_names_learning_db.txt`)
- Shows success message: "Applied X corrections to learning database"
- The corrections panel closes automatically
- **Next time you run a conversion, these wines will be matched automatically!**

#### Alternative: Load Corrections Manually
If you closed the panel or want to review corrections from a previous run:
1. Click **"📝 Load Corrections"** button (orange button)
2. Browse to `Outputs\Detailed match results\`
3. Select any `CORRECTIONS_NEEDED_*.txt` file
4. Make corrections and apply

---

### Step 5: Review the Results

After clicking "Apply All Corrections" (or if no corrections were needed), the system generates output files:

#### Output Files Generated:

**1. ItemNo_Results File** (Match Report)
- **Location**: `C:\Users\Marco.Africani\Desktop\Month recap\Outputs\Detailed match results\ItemNo_Results_20251126_082759.txt`
- **Purpose**: Shows which wines were matched and how (Learning DB, Name+Vintage, Price-only)
- **Check this to verify**: 99% of wines should show as "MATCHED"

Example content:
```
✓ Processing: Bollinger RD | 2008
   Wine: Bollinger RD | Vintage: 2008
   ✓ MATCHED: Item No. 56997 (✓ Learning DB)
      Wine: Champagne Extra Brut R.D. 2008
      Similarity: 100.0%
```

**2. Lines.xlsx File** (MAIN OUTPUT - Your Final Deliverable)
- **Location**: `C:\Users\Marco.Africani\Desktop\Month recap\Outputs\Detailed match results\Main offer\Lines_20251126_082806.xlsx`
- **Purpose**: Excel file with all matched wines in the **exact same order** as your Multi.txt
- **Columns**: Item Number, Wine Name, Vintage, CHF Price, EUR Price, etc.

**This is the file you use for your work!**

**3. Multi_converted File** (Converted Text)
- **Location**: `C:\Users\Marco.Africani\Desktop\Month recap\Outputs\Multi_converted_20251126_082806.txt`
- **Purpose**: Your original text with CHF prices converted to EUR
- **Use case**: If you need the text format with EUR prices

**4. Recognition Report** (Optional - for debugging)
- Shows detailed matching statistics and any issues

---

## Understanding the Matching Process

The application uses a **3-tier matching system** to find the correct Item Number and EUR price:

### Tier 1: Learning Database (Highest Priority)
- Uses wines you've previously corrected
- **100% accurate** because you verified them
- Stored in `wine_names_learning_db.txt`
- Format: `Wine Name | Vintage | Item No. | Timestamp`

### Tier 2: Name + Vintage + Price Match
- Matches wine name AND vintage AND CHF price
- Applies filters:
  - Competitor Code must be empty
  - Campaign Status must be "Sent"
  - Correct bottle size (75cl or 150cl)
- High confidence matches

### Tier 3: Price-Only Match (Needs Review)
- Only the CHF price matches
- Wine name didn't match the database
- **These go to the Corrections Panel** for your review
- Could be correct or could need manual correction

---

## Special Cases Handled Automatically

### Champagne Wines
- **Non-vintage (NV)**: The app recognizes Champagne without years
- **Producer in name**: "Krug Rosé 29ème Édition" is handled correctly (Krug is the producer)
- **Magnum detection**: If "Magnum" appears in the text, it filters for 150cl size

### 36-Bottle Pricing
- If your text shows two prices (e.g., "CHF 230.00 per bottle, or CHF 220.00 for 36 bottles")
- The app automatically detects the 36-bottle price
- Creates a second line in Lines.xlsx with "Min Qty: 36"

### Market Price Conversion
- Prices labeled as "market price" are automatically multiplied by 1.08
- No manual calculation needed

### Order Preservation
- **Critical feature**: Lines.xlsx maintains the exact order from Multi.txt
- Even if the matching process shuffles data internally, the final output is in your original order

---

## Troubleshooting

### Problem: No corrections panel appears
**Solution**: This means all wines matched successfully! Check Lines.xlsx - it should have all your wines.

### Problem: Too many corrections needed (>20%)
**Solutions**:
1. Check if wine names in Multi.txt are spelled correctly
2. Verify vintages are correct
3. Make sure CHF prices match the database
4. After correcting once, the Learning Database will remember them

### Problem: Wrong Item Numbers in corrections panel
**Solution**:
1. Look up the correct Item Number in your source database
2. Enter it in the "Correct Item No." field
3. Click "Apply All Corrections"
4. The system will remember this for future runs

### Problem: Wine appears twice in Lines.xlsx
**Explanation**: This is normal if you have both single-bottle and 36-bottle pricing in your Multi.txt
- First line: Single bottle price (Min Qty: 0)
- Second line: 36-bottle price (Min Qty: 36)

### Problem: Conversion failed with return code 1
**Common causes**:
1. Database files not found (check paths)
2. Multi.txt file is empty or has no recognizable wines
3. Unicode/encoding issues with special characters

**Solution**: Check the error message in the GUI results window

---

## Best Practices

### For Better Matching:
1. **Use consistent wine naming** in Multi.txt
   - "Château Margaux 2022" is better than "Margaux 22"
   - Include full vintage year (2022, not '22)

2. **Check the Learning Database regularly**
   - Click **"🔄 Refresh DB"** to see what's been learned
   - Shows the last 50 corrections you've made

3. **Review CORRECTIONS_NEEDED files**
   - These files are saved in `Outputs\Detailed match results\`
   - Keep them as a reference for patterns of corrections needed

4. **Run a test with a small file first**
   - Before processing a large Multi.txt, test with 5-10 wines
   - Verify the output format is what you expect

### For Faster Processing:
1. **Build your Learning Database** over time
   - The more corrections you make, the fewer you'll need to make in the future
   - After a few runs, most wines will match automatically

2. **Keep database files up to date**
   - Ensure OMT Main Offer List.xlsx is the latest version
   - Check that Stock Lines.xlsx is current

---

## File Structure Reference

```
Month recap/
├── Inputs/
│   ├── Multi.txt                          ← YOUR INPUT FILE (edit this)
│   └── ItemNoGenerator.txt                ← Auto-generated wine list
│
├── Outputs/
│   ├── Detailed match results/
│   │   ├── Main offer/
│   │   │   └── Lines_TIMESTAMP.xlsx       ← MAIN OUTPUT (your deliverable)
│   │   ├── ItemNo_Results_TIMESTAMP.txt   ← Match report (review this)
│   │   └── CORRECTIONS_NEEDED_TIMESTAMP.txt ← If corrections needed
│   │
│   └── Multi_converted_TIMESTAMP.txt      ← Converted text with EUR prices
│
├── Database/
│   ├── Conversion_month.xlsx              ← OMT database (read-only)
│   └── Stock Lines.xlsx                   ← Stock database (read-only)
│
└── wine_names_learning_db.txt             ← Learning database (grows over time)
```

---

## Quick Reference: Button Guide

| Button | What It Does |
|--------|--------------|
| **SPIN (CHF → EUR)** | Main conversion button - runs the entire process |
| **📝 Load Corrections** | Manually load a CORRECTIONS_NEEDED file |
| **✔️ Apply Corrections** | Legacy button (use "Apply All Corrections" in panel instead) |
| **🔄 Refresh DB** | Reload and display the Learning Database |
| **Apply All Corrections** | (In corrections panel) Save your corrections |
| **Hide Corrections** | Close the corrections panel without saving |

---

## Typical Session Example

```
1. User edits Multi.txt with 50 wines
   ↓
2. User clicks "SPIN"
   ↓
3. System processes 50 wines in 45 seconds
   ↓
4. Corrections panel appears showing 3 wines needing review
   ↓
5. User reviews suggested Item Numbers
   - Wine 1: Suggested 65274 ✓ Looks correct (leave as-is)
   - Wine 2: Suggested 62045 ✗ Wrong! Should be 62046 (user edits)
   - Wine 3: Suggested 62116 ✓ Looks correct (leave as-is)
   ↓
6. User clicks "Apply All Corrections"
   ↓
7. System saves 3 corrections to Learning Database
   ↓
8. User reviews Lines.xlsx - all 50 wines present in correct order
   ↓
9. User uses Lines.xlsx for their work
   ↓
10. Next time: Those 3 wines will match automatically!
```

---

## Success Indicators

✅ **Good conversion:**
- 95%+ wines matched automatically
- Few or no corrections needed
- Lines.xlsx has all wines in correct order
- EUR prices look reasonable (roughly CHF × 1.08)

❌ **Needs attention:**
- Many corrections needed (>20%)
- Wines missing from Lines.xlsx
- Order is scrambled
- Prices look way off

---

## Getting Help

If you encounter issues:
1. Check the **ItemNo_Results** file for detailed match information
2. Review the **Recognition Report** for error messages
3. Verify your database files are current
4. Check that Multi.txt is properly formatted
5. Review this guide's Troubleshooting section

---

## Summary

**In short**: Edit Multi.txt → Click SPIN → Review corrections if needed → Get Lines.xlsx

The application does all the heavy lifting:
- Extracts wines from your text
- Matches to database
- Applies business rules (filters, sizes, campaigns)
- Learns from your corrections
- Maintains order
- Generates Excel output

After a few uses, the Learning Database will know your wines, and most conversions will be fully automatic!
