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
DEFAULT_WORD_FILE = r"C:\Users\Marco.Africani\Desktop\Month recap\month recap.docx"
DEFAULT_WINE_LIST = r"C:\Users\Marco.Africani\Desktop\Month recap\ItemNoGenerator.txt"
LOGO_PATH = r"C:\Users\Marco.Africani\Desktop\Month recap\static\images\spinner.jpg"
LEARNING_DB = r"C:\Users\Marco.Africani\Desktop\Month recap\wine_names_learning_db.txt"


class AVUEchoSpinner(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("AVU Echo Spinner")
        self.geometry("900x700")
        self.configure(bg="#1a1a1a")

        # Disable resizing for that classic Winamp look
        self.resizable(False, False)

        # Variables
        self.word_file_path = tk.StringVar(value=DEFAULT_WORD_FILE)
        self.wine_list_path = tk.StringVar(value=DEFAULT_WINE_LIST)

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
            text="Word Document:",
            font=("Arial", 10),
            fg="#ffffff",
            bg="#1a1a1a",
            width=15,
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

        # Wine list input
        wine_input_frame = tk.Frame(matcher_frame, bg="#1a1a1a")
        wine_input_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(
            wine_input_frame,
            text="Wine List:",
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
        """Browse for Word document"""
        filename = filedialog.askopenfilename(
            title="Select Word Document",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
            initialdir=Path(DEFAULT_WORD_FILE).parent
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

                # Run the converter with explicit UTF-8 encoding
                result = subprocess.run(
                    [sys.executable, "word_converter_improved.py"],
                    cwd=Path(__file__).parent,
                    capture_output=True,
                    text=True,
                    encoding='utf-8',
                    errors='replace',  # Replace invalid characters instead of crashing
                    timeout=120
                )

                output = result.stdout + result.stderr
                self.results_text.insert(tk.END, output)

                if result.returncode == 0:
                    self.update_status("✅ Conversion completed successfully!")
                    messagebox.showinfo("Success", "CHF → EUR conversion completed!\n\nCheck:\n- month recap_EUR.docx\n- Lines.xlsx")
                else:
                    self.update_status("❌ Conversion failed")
                    messagebox.showerror("Error", f"Conversion failed with return code {result.returncode}")

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

    def run_matcher(self):
        """Run wine_item_matcher.py in a separate thread"""
        def run():
            try:
                self.update_status("🔍 Matching wine names to Item Numbers...")
                self.results_text.delete(1.0, tk.END)
                self.results_text.insert(tk.END, "Running Wine Item Matcher...\n\n")
                self.results_text.update()

                # Run the matcher with explicit UTF-8 encoding
                result = subprocess.run(
                    [sys.executable, "wine_item_matcher.py"],
                    cwd=Path(__file__).parent,
                    capture_output=True,
                    text=True,
                    encoding='utf-8',
                    errors='replace',
                    timeout=60
                )

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

            # Display entries in a formatted table
            if valid_entries:
                self.results_text.insert(tk.END, f"{'Wine Name':<40} {'Vintage':<10} {'Item No.':<10}\n", "header")
                self.results_text.insert(tk.END, "-"*80 + "\n")

                for wine, vintage, item in valid_entries[-50:]:  # Show last 50 entries
                    wine_short = wine[:38] + ".." if len(wine) > 40 else wine
                    self.results_text.insert(tk.END, f"{wine_short:<40} {vintage:<10} {item:<10}\n")

                if len(valid_entries) > 50:
                    self.results_text.insert(tk.END, f"\n... showing last 50 of {len(valid_entries)} entries\n", "info")
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


def main():
    """Launch the application"""
    app = AVUEchoSpinner()
    app.mainloop()


if __name__ == "__main__":
    main()
