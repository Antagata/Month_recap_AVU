#!/usr/bin/env python3
"""
AVU Echo Spinner - Wine Price Converter & Item Number Matcher
A stylish desktop application for wine document processing
"""

import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
from pathlib import Path
import subprocess
import threading
import sys
import os
from PIL import Image, ImageTk

# Configuration
BASE_DIR = r"C:\Users\Marco.Africani\Desktop\Month recap"
DATABASE_DIR = r"C:\Users\Marco.Africani\OneDrive - AVU SA\AVU CPI Campaign\Puzzle_control_Reports\SOURCE_FILES"
DEFAULT_MULTI_FILE = rf"{BASE_DIR}\Inputs\Multi.txt"  # Changed from month recap.docx to Multi.txt
DEFAULT_WINE_LIST = rf"{BASE_DIR}\Inputs\ItemNoGenerator.txt"
LOGO_PATH = rf"{BASE_DIR}\static\images\spinner.jpg"
LEARNING_DB = rf"{BASE_DIR}\wine_names_learning_db.txt"
OUTPUTS_DIR = rf"{BASE_DIR}\Outputs"


class AVUEchoSpinner(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("AVU Echo Spinner")
        self.geometry("900x700")
        self.configure(bg="#1a1a1a")

        # Enable resizing - set minimum size
        self.resizable(True, True)
        self.minsize(900, 700)  # Minimum size to ensure readability

        # Variables
        self.word_file_path = tk.StringVar(value=DEFAULT_MULTI_FILE)
        self.wine_list_path = tk.StringVar(value=DEFAULT_WINE_LIST)
        self.enable_translations = tk.BooleanVar(value=True)  # Translation checkbox

        self.setup_ui()

    def setup_ui(self):
        """Setup the main UI components"""

        # ============ TITLE BAR ============
        title_frame = tk.Frame(self, bg="#2d2d2d", height=40)
        title_frame.pack(fill=tk.X, padx=2, pady=2)
        title_frame.pack_propagate(False)

        # Logo
        try:
            logo_img = Image.open(LOGO_PATH)
            logo_img = logo_img.resize((32, 32), Image.Resampling.LANCZOS)
            self.logo_photo = ImageTk.PhotoImage(logo_img)
            logo_label = tk.Label(title_frame, image=self.logo_photo, bg="#2d2d2d")
            logo_label.pack(side=tk.LEFT, padx=5)
        except Exception as e:
            print(f"Could not load logo: {e}")

        title_label = tk.Label(
            title_frame,
            text="AVU Echo Spinner",
            font=("Arial", 16, "bold"),
            fg="#00ff00",
            bg="#2d2d2d"
        )
        title_label.pack(side=tk.LEFT, padx=10)

        version_label = tk.Label(
            title_frame,
            text="v2.0",
            font=("Arial", 9),
            fg="#888888",
            bg="#2d2d2d"
        )
        version_label.pack(side=tk.LEFT, padx=5)

        # ============ WORD CONVERTER SECTION ============
        converter_frame = tk.LabelFrame(
            self,
            text=" CHF → EUR Converter ",
            font=("Arial", 11, "bold"),
            fg="#00ccff",
            bg="#1a1a1a",
            relief=tk.RIDGE,
            bd=3
        )
        converter_frame.pack(fill=tk.X, padx=10, pady=5)

        # Word file input
        word_input_frame = tk.Frame(converter_frame, bg="#1a1a1a")
        word_input_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(
            word_input_frame,
            text="Document to Convert:",
            font=("Arial", 10),
            fg="#ffffff",
            bg="#1a1a1a",
            width=18,
            anchor="w"
        ).pack(side=tk.LEFT)

        word_entry = tk.Entry(
            word_input_frame,
            textvariable=self.word_file_path,
            font=("Consolas", 9),
            bg="#2d2d2d",
            fg="#00ff00",
            insertbackground="#00ff00",
            relief=tk.FLAT,
            bd=2
        )
        word_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        browse_btn = tk.Button(
            word_input_frame,
            text="📁 Browse",
            command=self.browse_word_file,
            font=("Arial", 9),
            bg="#3d3d3d",
            fg="#ffffff",
            activebackground="#4d4d4d",
            activeforeground="#00ff00",
            relief=tk.RAISED,
            bd=2,
            cursor="hand2"
        )
        browse_btn.pack(side=tk.LEFT, padx=2)

        # Direct paragraph input (above SPIN button)
        direct_para_frame = tk.Frame(converter_frame, bg="#1a1a1a")
        direct_para_frame.pack(fill=tk.BOTH, padx=10, pady=5)

        # Create a frame for label and checkbox on same line
        para_label_frame = tk.Frame(direct_para_frame, bg="#1a1a1a")
        para_label_frame.pack(fill=tk.X, anchor="w")

        tk.Label(
            para_label_frame,
            text="Or Convert Paragraph Directly:",
            font=("Arial", 10),
            fg="#ffff00",
            bg="#1a1a1a"
        ).pack(side=tk.LEFT)

        # Translation checkbox next to paragraph label
        translation_checkbox = tk.Checkbutton(
            para_label_frame,
            text="Generate Translations (DE/FR)",
            variable=self.enable_translations,
            font=("Arial", 9),
            fg="#00ccff",
            bg="#1a1a1a",
            selectcolor="#2d2d2d",
            activebackground="#1a1a1a",
            activeforeground="#00ff00",
            cursor="hand2"
        )
        translation_checkbox.pack(side=tk.LEFT, padx=20)

        tk.Label(
            direct_para_frame,
            text="(Paste text with CHF prices, click SPIN to convert to EUR)",
            font=("Arial", 8),
            fg="#888888",
            bg="#1a1a1a"
        ).pack(anchor="w")

        self.direct_para_text = scrolledtext.ScrolledText(
            direct_para_frame,
            height=4,
            font=("Consolas", 9),
            bg="#2d2d2d",
            fg="#00ccff",
            insertbackground="#00ccff",
            relief=tk.FLAT,
            bd=2
        )
        self.direct_para_text.pack(fill=tk.BOTH, expand=True, pady=5)

        # SPIN button (main converter button)
        spin_btn = tk.Button(
            converter_frame,
            text="🔄 SPIN",
            command=self.run_converter,
            font=("Arial", 14, "bold"),
            bg="#ff6600",
            fg="#ffffff",
            activebackground="#ff8833",
            activeforeground="#ffffff",
            relief=tk.RAISED,
            bd=3,
            cursor="hand2",
            height=2
        )
        spin_btn.pack(pady=10, padx=20, fill=tk.X)

        # ============ WINE MATCHER SECTION ============
        matcher_frame = tk.LabelFrame(
            self,
            text=" Wine Item Number Matcher ",
            font=("Arial", 11, "bold"),
            fg="#ff00ff",
            bg="#1a1a1a",
            relief=tk.RIDGE,
            bd=3
        )
        matcher_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Wine list file input (original single-line)
        wine_input_frame = tk.Frame(matcher_frame, bg="#1a1a1a")
        wine_input_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(
            wine_input_frame,
            text="Wine List File:",
            font=("Arial", 10),
            fg="#ffffff",
            bg="#1a1a1a",
            width=15,
            anchor="w"
        ).pack(side=tk.LEFT)

        wine_entry = tk.Entry(
            wine_input_frame,
            textvariable=self.wine_list_path,
            font=("Consolas", 9),
            bg="#2d2d2d",
            fg="#ff00ff",
            insertbackground="#ff00ff",
            relief=tk.FLAT,
            bd=2
        )
        wine_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        browse_wine_btn = tk.Button(
            wine_input_frame,
            text="📁 Browse",
            command=self.browse_wine_list,
            font=("Arial", 9),
            bg="#3d3d3d",
            fg="#ffffff",
            activebackground="#4d4d4d",
            activeforeground="#ff00ff",
            relief=tk.RAISED,
            bd=2,
            cursor="hand2"
        )
        browse_wine_btn.pack(side=tk.LEFT, padx=2)

        edit_wine_btn = tk.Button(
            wine_input_frame,
            text="✏️ Edit",
            command=self.edit_wine_list,
            font=("Arial", 9),
            bg="#3d3d3d",
            fg="#ffffff",
            activebackground="#4d4d4d",
            activeforeground="#ffff00",
            relief=tk.RAISED,
            bd=2,
            cursor="hand2"
        )
        edit_wine_btn.pack(side=tk.LEFT, padx=2)

        # Direct wine input (multi-line text area)
        direct_input_frame = tk.Frame(matcher_frame, bg="#1a1a1a")
        direct_input_frame.pack(fill=tk.BOTH, padx=10, pady=5)

        tk.Label(
            direct_input_frame,
            text="Or Enter Directly:",
            font=("Arial", 10),
            fg="#ffff00",
            bg="#1a1a1a"
        ).pack(anchor="w")

        tk.Label(
            direct_input_frame,
            text="(Format: Wine Name | Vintage, one per line)",
            font=("Arial", 8),
            fg="#888888",
            bg="#1a1a1a"
        ).pack(anchor="w")

        self.direct_wine_text = scrolledtext.ScrolledText(
            direct_input_frame,
            height=4,
            font=("Consolas", 9),
            bg="#2d2d2d",
            fg="#00ff00",
            insertbackground="#00ff00",
            relief=tk.FLAT,
            bd=2
        )
        self.direct_wine_text.pack(fill=tk.BOTH, expand=True, pady=5)

        # Size filter
        size_frame = tk.Frame(matcher_frame, bg="#1a1a1a")
        size_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(
            size_frame,
            text="Bottle Size:",
            font=("Arial", 10),
            fg="#ffffff",
            bg="#1a1a1a",
            width=15,
            anchor="w"
        ).pack(side=tk.LEFT)

        self.size_filter = tk.StringVar(value="75.0")
        size_dropdown = ttk.Combobox(
            size_frame,
            textvariable=self.size_filter,
            values=["75.0", "150.0", "300.0", "All sizes"],
            state="readonly",
            font=("Arial", 9),
            width=15
        )
        size_dropdown.pack(side=tk.LEFT, padx=5)

        tk.Label(
            size_frame,
            text="(Default: 75cl - standard bottle)",
            font=("Arial", 8),
            fg="#888888",
            bg="#1a1a1a"
        ).pack(side=tk.LEFT, padx=5)

        # Control buttons
        control_frame = tk.Frame(matcher_frame, bg="#1a1a1a")
        control_frame.pack(fill=tk.X, padx=10, pady=5)

        match_btn = tk.Button(
            control_frame,
            text="🔍 Match Wines",
            command=self.run_matcher,
            font=("Arial", 10, "bold"),
            bg="#6600ff",
            fg="#ffffff",
            activebackground="#7711ff",
            activeforeground="#ffffff",
            relief=tk.RAISED,
            bd=2,
            cursor="hand2",
            width=20
        )
        match_btn.pack(side=tk.LEFT, padx=5)

        correct_btn = tk.Button(
            control_frame,
            text="✔️ Apply Corrections",
            command=self.apply_corrections,
            font=("Arial", 10, "bold"),
            bg="#00aa00",
            fg="#ffffff",
            activebackground="#00cc00",
            activeforeground="#ffffff",
            relief=tk.RAISED,
            bd=2,
            cursor="hand2",
            width=20
        )
        correct_btn.pack(side=tk.LEFT, padx=5)

        refresh_btn = tk.Button(
            control_frame,
            text="🔄 Refresh DB",
            command=self.refresh_learning_db,
            font=("Arial", 10),
            bg="#3d3d3d",
            fg="#ffffff",
            activebackground="#4d4d4d",
            activeforeground="#00ff00",
            relief=tk.RAISED,
            bd=2,
            cursor="hand2",
            width=15
        )
        refresh_btn.pack(side=tk.LEFT, padx=5)

        load_corrections_btn = tk.Button(
            control_frame,
            text="📝 Load Corrections",
            command=self.load_corrections_manually,
            font=("Arial", 10),
            bg="#ff6600",
            fg="#ffffff",
            activebackground="#ff8800",
            activeforeground="#ffffff",
            relief=tk.RAISED,
            bd=2,
            cursor="hand2",
            width=18
        )
        load_corrections_btn.pack(side=tk.LEFT, padx=5)

        # Results display
        results_label = tk.Label(
            matcher_frame,
            text="Learning Database Results:",
            font=("Arial", 10, "bold"),
            fg="#ffff00",
            bg="#1a1a1a",
            anchor="w"
        )
        results_label.pack(fill=tk.X, padx=10, pady=(10, 2))

        # Scrolled text for results
        self.results_text = scrolledtext.ScrolledText(
            matcher_frame,
            font=("Consolas", 9),
            bg="#0d0d0d",
            fg="#00ff00",
            insertbackground="#00ff00",
            relief=tk.SUNKEN,
            bd=2,
            wrap=tk.WORD,
            height=15
        )
        self.results_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # ============ CORRECTIONS PANEL (INTERACTIVE) ============
        self.corrections_frame = tk.LabelFrame(
            self,
            text=" Wine Corrections - Interactive ",
            font=("Arial", 11, "bold"),
            fg="#ff9900",
            bg="#1a1a1a",
            relief=tk.RIDGE,
            bd=3
        )
        # Initially hidden - will show when corrections are needed

        corrections_info = tk.Label(
            self.corrections_frame,
            text="Wines needing correction will appear here. Enter the correct Item Number for each wine.",
            font=("Arial", 9),
            fg="#ffcc00",
            bg="#1a1a1a",
            wraplength=850,
            justify=tk.LEFT
        )
        corrections_info.pack(fill=tk.X, padx=10, pady=5)

        # Scrollable frame for corrections table
        corrections_canvas = tk.Canvas(self.corrections_frame, bg="#1a1a1a", height=200)
        corrections_scrollbar = ttk.Scrollbar(self.corrections_frame, orient="vertical", command=corrections_canvas.yview)
        self.corrections_table_frame = tk.Frame(corrections_canvas, bg="#1a1a1a")

        self.corrections_table_frame.bind(
            "<Configure>",
            lambda e: corrections_canvas.configure(scrollregion=corrections_canvas.bbox("all"))
        )

        corrections_canvas.create_window((0, 0), window=self.corrections_table_frame, anchor="nw")
        corrections_canvas.configure(yscrollcommand=corrections_scrollbar.set)

        corrections_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        corrections_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Action buttons for corrections
        corrections_btn_frame = tk.Frame(self.corrections_frame, bg="#1a1a1a")
        corrections_btn_frame.pack(fill=tk.X, padx=10, pady=5)

        apply_corrections_btn = tk.Button(
            corrections_btn_frame,
            text="Apply All Corrections",
            command=self.apply_interactive_corrections,
            font=("Arial", 10, "bold"),
            bg="#00aa00",
            fg="#ffffff",
            activebackground="#00cc00",
            relief=tk.RAISED,
            bd=2,
            cursor="hand2",
            width=20
        )
        apply_corrections_btn.pack(side=tk.LEFT, padx=5)

        hide_corrections_btn = tk.Button(
            corrections_btn_frame,
            text="Hide Corrections",
            command=self.hide_corrections_panel,
            font=("Arial", 10),
            bg="#666666",
            fg="#ffffff",
            relief=tk.RAISED,
            bd=2,
            cursor="hand2",
            width=15
        )
        hide_corrections_btn.pack(side=tk.LEFT, padx=5)

        # Storage for correction entry widgets
        self.correction_entries = []

        # ============ STATUS BAR ============
        status_frame = tk.Frame(self, bg="#2d2d2d", height=25)
        status_frame.pack(fill=tk.X, padx=2, pady=2)
        status_frame.pack_propagate(False)

        self.status_label = tk.Label(
            status_frame,
            text="Ready",
            font=("Arial", 9),
            fg="#00ff00",
            bg="#2d2d2d",
            anchor="w"
        )
        self.status_label.pack(side=tk.LEFT, padx=10)

        # Load initial learning database
        self.refresh_learning_db()

    def browse_word_file(self):
        """Browse for document to convert"""
        filename = filedialog.askopenfilename(
            title="Select Document to Convert",
            filetypes=[("Text Files", "*.txt"), ("Word Documents", "*.docx"), ("All Files", "*.*")],
            initialdir=Path(DEFAULT_MULTI_FILE).parent
        )
        if filename:
            self.word_file_path.set(filename)
            self.update_status(f"Selected: {Path(filename).name}")

    def browse_wine_list(self):
        """Browse for wine list text file"""
        filename = filedialog.askopenfilename(
            title="Select Wine List",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")],
            initialdir=Path(DEFAULT_WINE_LIST).parent
        )
        if filename:
            self.wine_list_path.set(filename)
            self.update_status(f"Selected: {Path(filename).name}")

    def edit_wine_list(self):
        """Open wine list in default text editor"""
        try:
            wine_file = self.wine_list_path.get()
            if sys.platform == "win32":
                os.startfile(wine_file)
            else:
                subprocess.run(["xdg-open", wine_file])
            self.update_status("Opened wine list in editor")
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file:\n{e}")

    def update_status(self, message):
        """Update status bar"""
        self.status_label.config(text=message)
        self.update_idletasks()

    def run_converter(self):
        """Run word_converter_improved.py in a separate thread"""
        def run():
            try:
                self.update_status("🔄 Converting CHF to EUR...")
                self.results_text.delete(1.0, tk.END)
                self.results_text.insert(tk.END, "Running CHF → EUR Converter...\n\n")
                self.results_text.update()

                # Check if user entered text directly
                direct_input = self.direct_para_text.get("1.0", tk.END).strip()

                if direct_input:
                    # Direct paragraph conversion
                    self.results_text.insert(tk.END, "Converting paragraph directly...\n\n")
                    converted_text = self.convert_paragraph_direct(direct_input)

                    if converted_text:
                        # Save to output file
                        from datetime import datetime
                        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                        output_file = Path(OUTPUTS_DIR) / f"Converted_Paragraph_{timestamp}.txt"

                        with open(output_file, 'w', encoding='utf-8') as f:
                            f.write("="*80 + "\n")
                            f.write("DIRECT PARAGRAPH CONVERSION - CHF to EUR\n")
                            f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                            f.write("="*80 + "\n\n")
                            f.write("ORIGINAL TEXT:\n")
                            f.write("-"*80 + "\n")
                            f.write(direct_input + "\n\n")
                            f.write("CONVERTED TEXT:\n")
                            f.write("-"*80 + "\n")
                            f.write(converted_text + "\n")

                        # Display result
                        self.results_text.insert(tk.END, f"✅ Conversion completed!\n\n")
                        self.results_text.insert(tk.END, f"ORIGINAL:\n{direct_input}\n\n")
                        self.results_text.insert(tk.END, f"CONVERTED:\n{converted_text}\n\n")
                        self.results_text.insert(tk.END, f"📁 Saved to: {output_file.name}\n")

                        self.update_status("✅ Paragraph conversion completed!")
                        messagebox.showinfo("Success", f"Paragraph converted successfully!\n\nSaved to:\n{output_file}")
                    else:
                        self.update_status("❌ Conversion failed")
                        messagebox.showerror("Error", "Could not convert paragraph")
                else:
                    # Integrated workflow: Match wines then convert
                    # Now using integrated_converter.py instead of txt_converter.py
                    result = subprocess.run(
                        [sys.executable, "integrated_converter.py"],
                        cwd=Path(__file__).parent,
                        capture_output=True,
                        text=True,
                        encoding='utf-8',
                        errors='replace',
                        timeout=300  # Increased timeout for full workflow
                    )

                    output = result.stdout + result.stderr
                    self.results_text.insert(tk.END, output)

                    if result.returncode == 0:
                        self.update_status("✅ Conversion completed successfully!")
                        messagebox.showinfo("Success", "Wine recognition and conversion completed!\n\nCheck:\n- Outputs/Multi_converted_XXX.txt\n- Outputs/Stock_Lines_Filtered_XXX.xlsx\n- Outputs/Detailed match results/Recognition_Report_XXX.txt")

                        # Check for corrections file and show panel if needed
                        self.check_for_corrections_file()
                    else:
                        self.update_status("❌ Conversion failed")
                        messagebox.showerror("Error", f"Conversion failed with return code {result.returncode}")

                        # Even if conversion failed, check for corrections that might have been generated
                        self.check_for_corrections_file()

            except subprocess.TimeoutExpired:
                self.update_status("⏱️ Conversion timeout")
                self.results_text.insert(tk.END, "\n\n⚠️ Process timed out after 2 minutes")
                messagebox.showwarning("Timeout", "Conversion took too long and was terminated")
            except Exception as e:
                self.update_status(f"❌ Error: {str(e)}")
                self.results_text.insert(tk.END, f"\n\n❌ ERROR: {e}")
                messagebox.showerror("Error", f"An error occurred:\n{e}")

        # Run in thread to avoid freezing UI
        thread = threading.Thread(target=run, daemon=True)
        thread.start()

    def convert_paragraph_direct(self, text):
        """Convert a paragraph of text from CHF to EUR using a simplified approach"""
        try:
            import pandas as pd
            import re

            # Load the Excel database
            excel_file = rf"{DATABASE_DIR}\OMT Main Offer List.xlsx"
            df = pd.read_excel(excel_file)

            # Create conversion map (CHF -> EUR)
            df['CHF_KEY'] = df['Unit Price'].astype(float).round(2).apply(lambda x: f'{x:.2f}')
            df['EUR_VALUE'] = df['Unit Price (EUR)'].astype(float).round(0).apply(lambda x: f'{int(x)}.00')

            conversion_map = {}
            for _, row in df.iterrows():
                chf = row['CHF_KEY']
                eur = row['EUR_VALUE']
                if chf not in conversion_map:
                    conversion_map[chf] = eur

            # Find and replace CHF prices
            converted = text

            # Pattern 1: "CHF XX.XX"
            for match in re.finditer(r'[Cc][Hh][Ff]\s+(\d+(?:[\']\d{3})*\.?\d{0,2})', converted):
                chf_str = match.group(1).replace("'", "")
                if '.' not in chf_str:
                    chf_str += '.00'
                elif len(chf_str.split('.')[1]) == 1:
                    chf_str += '0'

                eur_value = conversion_map.get(chf_str, f"{int(float(chf_str) * 1.08)}.00")
                converted = converted.replace(match.group(0), f'EUR {eur_value}')

            # Pattern 2: "XX.XX CHF"
            for match in re.finditer(r'(\d+(?:[\']\d{3})*\.?\d{0,2})\s+[Cc][Hh][Ff]', converted):
                chf_str = match.group(1).replace("'", "")
                if '.' not in chf_str:
                    chf_str += '.00'
                elif len(chf_str.split('.')[1]) == 1:
                    chf_str += '0'

                eur_value = conversion_map.get(chf_str, f"{int(float(chf_str) * 1.08)}.00")
                converted = converted.replace(match.group(0), f'{eur_value} EUR')

            # Replace any remaining CHF with EUR
            converted = re.sub(r'\b[Cc][Hh][Ff]\b', 'EUR', converted)

            return converted

        except Exception as e:
            self.results_text.insert(tk.END, f"\n❌ Error during conversion: {e}\n")
            return None

    def run_matcher(self):
        """Run wine_item_matcher.py in a separate thread"""
        def run():
            try:
                self.update_status("🔍 Matching wine names to Item Numbers...")
                self.results_text.delete(1.0, tk.END)
                self.results_text.insert(tk.END, "Running Wine Item Matcher...\n\n")
                self.results_text.update()

                # Check if user entered wines directly
                direct_input = self.direct_wine_text.get("1.0", tk.END).strip()

                # Get size filter
                size_filter = self.size_filter.get()

                # Prepare command
                cmd = [sys.executable, "wine_item_matcher.py"]

                # Add size parameter
                if size_filter != "All sizes":
                    cmd.extend(["--size", size_filter])

                # If direct input provided, create temporary file
                temp_input_file = None
                if direct_input:
                    self.results_text.insert(tk.END, f"Using direct input with {len(direct_input.splitlines())} wines\n")
                    self.results_text.insert(tk.END, f"Size filter: {size_filter}\n\n")

                    # Create temporary input file
                    import tempfile
                    temp_input_file = Path(tempfile.gettempdir()) / "avu_temp_wine_input.txt"
                    with open(temp_input_file, 'w', encoding='utf-8') as f:
                        f.write(direct_input)

                    cmd.extend(["--input", str(temp_input_file)])
                else:
                    # Use file path
                    wine_file = self.wine_list_path.get()
                    if wine_file:
                        cmd.extend(["--input", wine_file])
                    self.results_text.insert(tk.END, f"Using wine list file: {Path(wine_file).name}\n")
                    self.results_text.insert(tk.END, f"Size filter: {size_filter}\n\n")

                # Run the matcher with explicit UTF-8 encoding
                result = subprocess.run(
                    cmd,
                    cwd=Path(__file__).parent,
                    capture_output=True,
                    text=True,
                    encoding='utf-8',
                    errors='replace',
                    timeout=60
                )

                # Clean up temp file
                if temp_input_file and temp_input_file.exists():
                    temp_input_file.unlink()

                output = result.stdout + result.stderr
                self.results_text.insert(tk.END, output)

                if result.returncode == 0:
                    self.update_status("✅ Wine matching completed!")
                    # Refresh learning database display
                    self.after(100, self.refresh_learning_db)
                    messagebox.showinfo("Success", "Wine matching completed!\n\nCheck results in ItemNo_Results_[timestamp].txt")
                else:
                    self.update_status("❌ Matching failed")
                    messagebox.showerror("Error", f"Matching failed with return code {result.returncode}")

            except subprocess.TimeoutExpired:
                self.update_status("⏱️ Matching timeout")
                self.results_text.insert(tk.END, "\n\n⚠️ Process timed out after 1 minute")
                messagebox.showwarning("Timeout", "Matching took too long and was terminated")
            except Exception as e:
                self.update_status(f"❌ Error: {str(e)}")
                self.results_text.insert(tk.END, f"\n\n❌ ERROR: {e}")
                messagebox.showerror("Error", f"An error occurred:\n{e}")

        thread = threading.Thread(target=run, daemon=True)
        thread.start()

    def apply_corrections(self):
        """Run apply_corrections.py in a separate thread"""
        def run():
            try:
                self.update_status("✔️ Applying corrections to learning database...")
                self.results_text.delete(1.0, tk.END)
                self.results_text.insert(tk.END, "Applying Corrections...\n\n")
                self.results_text.update()

                # Run apply corrections with explicit UTF-8 encoding
                result = subprocess.run(
                    [sys.executable, "apply_corrections.py"],
                    cwd=Path(__file__).parent,
                    capture_output=True,
                    text=True,
                    encoding='utf-8',
                    errors='replace',
                    timeout=30
                )

                output = result.stdout + result.stderr
                self.results_text.insert(tk.END, output)

                if result.returncode == 0:
                    self.update_status("✅ Corrections applied successfully!")
                    # Refresh learning database display
                    self.after(100, self.refresh_learning_db)
                    messagebox.showinfo("Success", "Corrections applied to learning database!")
                else:
                    self.update_status("❌ Apply corrections failed")
                    messagebox.showerror("Error", f"Apply corrections failed with return code {result.returncode}")

            except subprocess.TimeoutExpired:
                self.update_status("⏱️ Apply corrections timeout")
                self.results_text.insert(tk.END, "\n\n⚠️ Process timed out after 30 seconds")
            except Exception as e:
                self.update_status(f"❌ Error: {str(e)}")
                self.results_text.insert(tk.END, f"\n\n❌ ERROR: {e}")
                messagebox.showerror("Error", f"An error occurred:\n{e}")

        thread = threading.Thread(target=run, daemon=True)
        thread.start()

    def refresh_learning_db(self):
        """Load and display learning database contents"""
        try:
            self.results_text.delete(1.0, tk.END)

            if not Path(LEARNING_DB).exists():
                self.results_text.insert(tk.END, "⚠️ Learning database not found\n")
                self.update_status("Learning database not found")
                return

            with open(LEARNING_DB, 'r', encoding='utf-8') as f:
                lines = f.readlines()

            # Count entries
            valid_entries = []
            not_found_count = 0

            for line in lines:
                line = line.strip()
                if line and not line.startswith('#'):
                    parts = line.split(' | ')
                    if len(parts) >= 3:
                        wine_name = parts[0]
                        vintage = parts[1]
                        item_no = parts[2].split()[0]  # Get just the item number

                        if item_no == 'NOT_FOUND':
                            not_found_count += 1
                        else:
                            valid_entries.append((wine_name, vintage, item_no))

            # Display header
            self.results_text.insert(tk.END, "="*80 + "\n")
            self.results_text.insert(tk.END, "LEARNING DATABASE - WINE → ITEM NUMBER MAPPINGS\n", "header")
            self.results_text.insert(tk.END, "="*80 + "\n\n")

            self.results_text.insert(tk.END, f"📊 Statistics:\n")
            self.results_text.insert(tk.END, f"  • Total Valid Mappings: {len(valid_entries)}\n", "success")
            self.results_text.insert(tk.END, f"  • NOT FOUND Entries: {not_found_count}\n", "error")
            self.results_text.insert(tk.END, "\n" + "-"*80 + "\n\n")

            # Display entries in a formatted table (latest first)
            if valid_entries:
                self.results_text.insert(tk.END, f"{'Wine Name':<40} {'Vintage':<10} {'Item No.':<10}\n", "header")
                self.results_text.insert(tk.END, "-"*80 + "\n")

                # Show last 50 entries in reverse order (latest first)
                display_entries = valid_entries[-50:][::-1]
                for wine, vintage, item in display_entries:
                    wine_short = wine[:38] + ".." if len(wine) > 40 else wine
                    self.results_text.insert(tk.END, f"{wine_short:<40} {vintage:<10} {item:<10}\n")

                if len(valid_entries) > 50:
                    self.results_text.insert(tk.END, f"\n... showing latest 50 of {len(valid_entries)} entries\n", "info")
            else:
                self.results_text.insert(tk.END, "No valid entries found in learning database\n", "error")

            # Configure text tags for colors
            self.results_text.tag_config("header", foreground="#ffff00", font=("Consolas", 9, "bold"))
            self.results_text.tag_config("success", foreground="#00ff00")
            self.results_text.tag_config("error", foreground="#ff6666")
            self.results_text.tag_config("info", foreground="#00ccff")

            self.update_status(f"✅ Learning DB loaded: {len(valid_entries)} valid mappings")

        except Exception as e:
            self.results_text.insert(tk.END, f"❌ Error loading learning database:\n{e}\n")
            self.update_status(f"❌ Error loading DB: {str(e)}")

    def load_corrections_file(self, corrections_file_path):
        """Parse CORRECTIONS_NEEDED file and populate interactive corrections table"""
        try:
            corrections = []
            current_entry = {}

            with open(corrections_file_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()

            for line in lines:
                line_stripped = line.strip()

                if not line_stripped or line_stripped.startswith('='):
                    continue

                # Parse entry fields
                if line_stripped.startswith('Name:'):
                    current_entry['wine_name'] = line_stripped.replace('Name:', '').strip()
                elif line_stripped.startswith('Vintage:') and 'wine_name' in current_entry:
                    current_entry['vintage'] = line_stripped.replace('Vintage:', '').strip()
                elif line_stripped.startswith('CHF Price:'):
                    current_entry['chf_price'] = line_stripped.replace('CHF Price:', '').strip()
                elif 'MATCHED TO DATABASE:' in line:
                    # Next section starts
                    pass
                elif line_stripped.startswith('Wine:') and 'chf_price' in current_entry:
                    current_entry['matched_wine'] = line_stripped.replace('Wine:', '').strip()
                elif line_stripped.startswith('Item No.:'):
                    current_entry['item_no'] = line_stripped.replace('Item No.:', '').strip()
                elif line_stripped.startswith('REASON:'):
                    current_entry['reason'] = line_stripped.replace('REASON:', '').strip()

                    # Entry complete
                    if current_entry.get('wine_name') and current_entry.get('vintage'):
                        corrections.append(current_entry.copy())
                    current_entry = {}

            return corrections

        except Exception as e:
            print(f"Error loading corrections file: {e}")
            return []

    def show_corrections_panel(self, corrections_file_path):
        """Display the corrections panel with wines needing correction"""
        corrections = self.load_corrections_file(corrections_file_path)

        if not corrections:
            messagebox.showinfo("No Corrections", "No wines found needing correction.")
            return

        # Clear existing table
        for widget in self.corrections_table_frame.winfo_children():
            widget.destroy()
        self.correction_entries.clear()

        # Create header
        header_frame = tk.Frame(self.corrections_table_frame, bg="#2d2d2d", relief=tk.RAISED, bd=2)
        header_frame.pack(fill=tk.X, padx=5, pady=5)

        tk.Label(header_frame, text="Wine Name", font=("Arial", 9, "bold"), fg="#ffff00", bg="#2d2d2d", width=25, anchor="w").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        tk.Label(header_frame, text="Vintage", font=("Arial", 9, "bold"), fg="#ffff00", bg="#2d2d2d", width=8).grid(row=0, column=1, padx=5, pady=5)
        tk.Label(header_frame, text="Price", font=("Arial", 9, "bold"), fg="#ffff00", bg="#2d2d2d", width=8).grid(row=0, column=2, padx=5, pady=5)
        tk.Label(header_frame, text="Suggested Item No.", font=("Arial", 9, "bold"), fg="#ffff00", bg="#2d2d2d", width=15).grid(row=0, column=3, padx=5, pady=5)
        tk.Label(header_frame, text="Correct Item No.", font=("Arial", 9, "bold"), fg="#00ff00", bg="#2d2d2d", width=15).grid(row=0, column=4, padx=5, pady=5)
        tk.Label(header_frame, text="Reason", font=("Arial", 9, "bold"), fg="#ffff00", bg="#2d2d2d", width=25, anchor="w").grid(row=0, column=5, padx=5, pady=5, sticky="w")

        # Create rows for each correction
        for idx, corr in enumerate(corrections):
            row_frame = tk.Frame(self.corrections_table_frame, bg="#1a1a1a", relief=tk.GROOVE, bd=1)
            row_frame.pack(fill=tk.X, padx=5, pady=2)

            wine_short = corr['wine_name'][:22] + "..." if len(corr['wine_name']) > 25 else corr['wine_name']
            reason_short = corr.get('reason', 'N/A')[:22] + "..." if len(corr.get('reason', '')) > 25 else corr.get('reason', 'N/A')

            tk.Label(row_frame, text=wine_short, font=("Arial", 9), fg="#ffffff", bg="#1a1a1a", width=25, anchor="w").grid(row=0, column=0, padx=5, pady=3, sticky="w")
            tk.Label(row_frame, text=corr['vintage'], font=("Arial", 9), fg="#ffffff", bg="#1a1a1a", width=8).grid(row=0, column=1, padx=5, pady=3)
            tk.Label(row_frame, text=corr['chf_price'], font=("Arial", 9), fg="#ffffff", bg="#1a1a1a", width=8).grid(row=0, column=2, padx=5, pady=3)
            tk.Label(row_frame, text=corr.get('item_no', 'N/A'), font=("Arial", 9), fg="#ff9900", bg="#1a1a1a", width=15).grid(row=0, column=3, padx=5, pady=3)

            # Entry for user to input correct Item No.
            item_entry = tk.Entry(row_frame, font=("Arial", 10), bg="#2d2d2d", fg="#00ff00", insertbackground="#00ff00", width=15)
            item_entry.grid(row=0, column=4, padx=5, pady=3)

            # Pre-fill with suggested Item No if it's not MANUAL_ENTRY_NEEDED
            if corr.get('item_no') and corr['item_no'] not in ['MANUAL_ENTRY_NEEDED', 'NOT_FOUND']:
                item_entry.insert(0, corr['item_no'])

            tk.Label(row_frame, text=reason_short, font=("Arial", 8), fg="#cccccc", bg="#1a1a1a", width=25, anchor="w").grid(row=0, column=5, padx=5, pady=3, sticky="w")

            # Store entry widget with wine info
            self.correction_entries.append({
                'wine_name': corr['wine_name'],
                'vintage': corr['vintage'],
                'entry': item_entry
            })

        # Show the panel
        self.corrections_frame.pack(fill=tk.BOTH, expand=False, padx=10, pady=10, before=self.winfo_children()[-1])
        self.update_status(f"Showing {len(corrections)} wines needing correction")

    def hide_corrections_panel(self):
        """Hide the corrections panel"""
        self.corrections_frame.pack_forget()
        self.update_status("Corrections panel hidden")

    def apply_interactive_corrections(self):
        """Apply corrections entered by user in the interactive panel"""
        if not self.correction_entries:
            messagebox.showwarning("No Corrections", "No corrections to apply.")
            return

        corrections = []
        invalid_count = 0

        for entry_info in self.correction_entries:
            item_no = entry_info['entry'].get().strip()

            if item_no:
                # Validate Item No is numeric
                try:
                    int(item_no)
                    corrections.append({
                        'wine_name': entry_info['wine_name'],
                        'vintage': entry_info['vintage'],
                        'item_no': item_no
                    })
                except ValueError:
                    invalid_count += 1

        if invalid_count > 0:
            messagebox.showwarning("Invalid Entries", f"{invalid_count} entries have invalid Item Numbers (must be numeric).\nThey will be skipped.")

        if not corrections:
            messagebox.showwarning("No Valid Corrections", "No valid corrections to apply.")
            return

        # Apply corrections to learning database
        try:
            from datetime import datetime

            # Load existing database keys to avoid duplicates
            existing_keys = set()
            if Path(LEARNING_DB).exists():
                with open(LEARNING_DB, 'r', encoding='utf-8') as f:
                    for line in f:
                        line = line.strip()
                        if line and not line.startswith('#'):
                            parts = line.split(' | ')
                            if len(parts) >= 3:
                                key = f"{parts[0]}|{parts[1]}|{parts[2]}"
                                existing_keys.add(key)

            # Write new corrections
            new_count = 0
            duplicate_count = 0

            with open(LEARNING_DB, 'a', encoding='utf-8') as f:
                for corr in corrections:
                    key = f"{corr['wine_name']}|{corr['vintage']}|{corr['item_no']}"

                    if key not in existing_keys:
                        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        entry_line = f"{corr['wine_name']} | {corr['vintage']} | {corr['item_no']} | {timestamp} (GUI correction)\n"
                        f.write(entry_line)
                        existing_keys.add(key)
                        new_count += 1
                    else:
                        duplicate_count += 1

            # Show success message
            msg = f"Applied {new_count} corrections to learning database"
            if duplicate_count > 0:
                msg += f"\nSkipped {duplicate_count} duplicates"

            messagebox.showinfo("Success", msg)

            # Hide panel and refresh database display
            self.hide_corrections_panel()
            self.refresh_learning_db()

            self.update_status(f"Applied {new_count} corrections successfully")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to apply corrections:\n{e}")
            self.update_status(f"Error applying corrections: {str(e)}")

    def load_corrections_manually(self):
        """Allow user to manually select a CORRECTIONS_NEEDED file to load"""
        corrections_dir = rf"{BASE_DIR}\Outputs\Detailed match results"
        filename = filedialog.askopenfilename(
            title="Select Corrections File",
            filetypes=[("Text Files", "CORRECTIONS_NEEDED_*.txt"), ("All Files", "*.*")],
            initialdir=corrections_dir
        )

        if filename:
            self.show_corrections_panel(filename)

    def check_for_corrections_file(self):
        """Check if a new CORRECTIONS_NEEDED file was created and show panel"""
        import glob

        corrections_dir = rf"{BASE_DIR}\Outputs\Detailed match results"
        pattern = str(Path(corrections_dir) / "CORRECTIONS_NEEDED_*.txt")
        corrections_files = glob.glob(pattern)

        if corrections_files:
            # Sort by modification time (most recent first)
            corrections_files.sort(key=lambda x: Path(x).stat().st_mtime, reverse=True)
            latest_file = corrections_files[0]

            # Check if file is less than 5 minutes old
            file_age = Path(latest_file).stat().st_mtime
            current_time = Path(latest_file).stat().st_mtime

            # Show corrections panel
            self.show_corrections_panel(latest_file)


def main():
    """Launch the application"""
    app = AVUEchoSpinner()
    app.mainloop()


if __name__ == "__main__":
    main()
