# AVU Echo Spinner - GUI Application

A stylish desktop application with a Winamp-inspired chrome/black aesthetic for wine document processing.

## Features

### 🔄 CHF → EUR Converter
- Upload Word document with CHF prices
- One-click conversion with the **SPIN** button
- Automatic price matching and highlighting
- Generates converted document with color-coded results

### 🔍 Wine Item Number Matcher
- Upload or edit wine list (text file)
- Match wine names to Item Numbers
- Uses 3-tier matching system:
  - Learning Database (instant)
  - Primary Database (27,577 wines)
  - Fallback Stock Database (4,790 wines)
- Apply corrections with one click

### 📊 Learning Database Viewer
- Real-time display of wine → Item No. mappings
- Shows statistics and recent entries
- Auto-refreshes after operations

## How to Use

### 1. Launch the Application
Double-click: `Launch AVU Echo Spinner.bat`

Or run directly:
```bash
python avu_echo_spinner.py
```

### 2. Convert CHF to EUR
1. Browse or confirm Word document path
2. Click the **🔄 SPIN** button
3. Wait for completion
4. Check output: `month recap_EUR.docx`

### 3. Match Wine Names
1. Edit or confirm wine list path (ItemNoGenerator.txt)
2. Click **🔍 Match Wines**
3. View results in the display panel
4. If corrections needed, edit the CORRECTIONS_NEEDED_*.txt file
5. Click **✔️ Apply Corrections**

### 4. View Learning Database
- Click **🔄 Refresh DB** to reload
- Shows all learned wine → Item No. mappings
- Displays statistics and recent entries

## Interface Elements

### Top Section - CHF → EUR Converter
- **Word Document Field**: Path to input Word file
- **📁 Browse**: Select Word document
- **🔄 SPIN Button**: Run converter (orange, large)

### Bottom Section - Wine Matcher
- **Wine List Field**: Path to wine names text file
- **📁 Browse**: Select wine list
- **✏️ Edit**: Open wine list in text editor
- **🔍 Match Wines**: Run matcher (purple)
- **✔️ Apply Corrections**: Apply manual corrections (green)
- **🔄 Refresh DB**: Reload learning database (gray)
- **Results Panel**: Displays output and learning database

### Status Bar
- Shows current operation status
- Updates in real-time

## Keyboard Shortcuts

*Future enhancement - currently click-based interface*

## Technical Details

### Requirements
- Python 3.8+
- tkinter (included with Python)
- Pillow (for logo display)
- All backend scripts must be in the same directory

### File Paths
Default paths can be changed in the application:
- Word File: `month recap.docx`
- Wine List: `ItemNoGenerator.txt`
- Logo: `static/images/spinner.jpg`
- Learning DB: `wine_names_learning_db.txt`

### Threading
All operations run in background threads to keep the UI responsive.
This prevents freezing during long operations.

### Timeout Settings
- Converter: 120 seconds (2 minutes)
- Matcher: 60 seconds (1 minute)
- Apply Corrections: 30 seconds

## Color Scheme

Inspired by Winamp's classic chrome aesthetic:

- **Background**: `#1a1a1a` (almost black)
- **Frames**: `#2d2d2d` (dark gray)
- **Text**: `#ffffff` (white)
- **Primary Accent**: `#00ff00` (neon green)
- **Secondary Accent**: `#00ccff` (cyan)
- **Tertiary Accent**: `#ff00ff` (magenta)
- **Action Button**: `#ff6600` (orange)
- **Success**: `#00aa00` (green)
- **Matcher**: `#6600ff` (purple)

## Troubleshooting

### Application won't start
```bash
# Check Python is installed
python --version

# Install Pillow if missing
pip install Pillow

# Run directly to see errors
python avu_echo_spinner.py
```

### Logo not displaying
- Check if `static/images/spinner.jpg` exists
- App will run without logo if file is missing

### Scripts not running
- Ensure all scripts are in the same directory:
  - `word_converter_improved.py`
  - `wine_item_matcher.py`
  - `apply_corrections.py`
  - `avu_echo_spinner.py`

### Conversion/Matching timeout
- Increase timeout values in `avu_echo_spinner.py`
- Check system resources (CPU/memory)

## Future Enhancements

- [ ] Drag-and-drop file upload
- [ ] Keyboard shortcuts
- [ ] Settings panel
- [ ] Export results to PDF
- [ ] Real-time progress bars
- [ ] Sound effects (optional)
- [ ] Minimize to system tray
- [ ] Dark/Light theme toggle

---

**Version**: 2.0
**Last Updated**: 2025-11-05
**Author**: AVU Echo Spinner Team
**License**: Internal Use
