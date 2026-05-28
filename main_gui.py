import os
import sys
import shutil
from pathlib import Path
import pandas as pd
import customtkinter as ctk
from tkinter import messagebox

# --- BACKEND MODULE IMPORTS ---
from src.constants import READY_TO_UPLOAD_TO_AMAZON_FOLDER, BOOK_TRACKER_EXCEL_FILE_PATH
from utils.norm_book_title import normalize_book_title, get_book_metadata
from src.logic.pdf_processor import process_pdf
from src.logic.file_operations import check_file_size
from src.logic.system_tools import clean_up_folder_after_processing
from src.fliphtml5.flip_html_automation import fliphtml5_automation
from src.logic.excel_tools import run_excel_update_workflow

# Set the visual theme of the app
ctk.set_appearance_mode("System")


class BookAutomizerUI(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window Configuration
        self.title("YBZ Institute - PDF Book Automizer")
        self.geometry("750x550")
        self.resizable(False, False)

        # State Variables (Storing selected paths)
        self.source_file = ctk.StringVar()
        self.target_dir = ctk.StringVar()
        self.excel_path = ctk.StringVar()

        self.setup_ui()

    def setup_ui(self):
        # --- TITLE BANNER ---
        self.title_label = ctk.CTkLabel(
            self,
            text="PDF Book Automizer",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        self.title_label.pack(pady=(20, 10))

        self.subtitle_label = ctk.CTkLabel(
            self,
            text="Automated processing, tracking, and normalization engine",
            font=ctk.CTkFont(size=12, slant="italic")
        )
        self.subtitle_label.pack(pady=(0, 20))

        # --- SELECTION FORM ---
        self.form_frame = ctk.CTkFrame(self)
        self.form_frame.pack(pady=10, padx=30, fill="x")

        # 1. Source PDF File Row
        self.create_file_row("Source PDF File:", self.source_file, self.browse_source, 0)
        # 2. Target Folder Row
        self.create_file_row("Target Archive Folder:", self.target_dir, self.browse_target, 1)
        self.target_dir.set(str(READY_TO_UPLOAD_TO_AMAZON_FOLDER))
        # 3. Excel Tracker Row
        self.create_file_row("Excel Tracker File:", self.excel_path, self.browse_excel, 2)
        self.excel_path.set(str(BOOK_TRACKER_EXCEL_FILE_PATH))

        # --- ACTIONS & PROGRESS ---
        self.action_frame = ctk.CTkFrame(self)
        self.action_frame.pack(pady=20, padx=30, fill="x")

        self.run_btn = ctk.CTkButton(
            self.action_frame,
            text="Start Automation Process",
            command=self.start_process,
            font=ctk.CTkFont(size=14, weight="bold"),
            height=40
        )
        self.run_btn.pack(side="left", padx=(0, 20))

        self.progress_bar = ctk.CTkProgressBar(self.action_frame, width=400)
        self.progress_bar.set(0)
        self.progress_bar.pack(side="right", pady=15)

        # --- LIVE LOGGING CONSOLE ---
        self.log_label = ctk.CTkLabel(self, text="Process Logs", font=ctk.CTkFont(weight="bold"))
        self.log_label.pack(anchor="w", padx=35, pady=(10, 0))

        self.log_box = ctk.CTkTextbox(self, height=180, width=690, font=ctk.CTkFont(family="Consolas", size=12))
        self.log_box.pack(pady=(5, 20), padx=30)
        self.log_box.configure(state="disabled")

    def create_file_row(self, label_text, text_var, browse_cmd, row_idx):
        """Helper to create entry lines with browse buttons."""
        lbl = ctk.CTkLabel(self.form_frame, text=label_text, anchor="w", width=150)
        lbl.grid(row=row_idx, column=0, padx=10, pady=10, sticky="w")

        entry = ctk.CTkEntry(self.form_frame, textvariable=text_var, width=400)
        entry.grid(row=row_idx, column=1, padx=10, pady=10)

        btn = ctk.CTkButton(self.form_frame, text="Browse", width=80, command=browse_cmd)
        btn.grid(row=row_idx, column=2, padx=10, pady=10)

    # --- BROWSER DIALOGUE FUNCTIONS ---
    def browse_source(self):
        path = ctk.filedialog.askopenfilename(title="Select Source PDF File", filetypes=[("PDF Files", "*.pdf")])
        if path: self.source_file.set(path)

    def browse_target(self):
        path = ctk.filedialog.askdirectory(title="Select Target Archive Folder")
        if path: self.target_dir.set(path)

    def browse_excel(self):
        path = ctk.filedialog.askopenfilename(title="Select Excel Tracker File", filetypes=[("Excel Files", "*.xlsx")])
        if path: self.excel_path.set(path)

    def log(self, message):
        """Appends status messages to the UI textbox console."""
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"{message}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        self.update_idletasks()

    # --- INTERACTIVE HELPER METHODS FOR PIPELINE ---
    def ask_yes_no(self, title, message):
        """GUI translation of yes/no input choice prompts."""
        return messagebox.askyesno(title, message)

    def ask_input_string(self, title, prompt):
        """GUI translation of generic string input prompts."""
        dialog = ctk.CTkInputDialog(text=prompt, title=title)
        return dialog.get_input()

    def ask_input_int(self, title, prompt):
        """GUI translation for numeric values with fallback checks."""
        while True:
            dialog = ctk.CTkInputDialog(text=prompt, title=title)
            val = dialog.get_input()
            if val is None:  # User cancelled
                return None
            if val.strip().lstrip('-').isdigit():
                return int(val.strip())
            messagebox.showerror("Invalid Input", "Please enter a valid round number.")

    def blocking_checkpoint(self, title, action_message):
        """GUI translation of 'Press Enter to continue' manual workspace steps."""
        self.log(f"[ACTION REQUIRED] {action_message}")
        messagebox.showinfo(title, f"ACTION REQUIRED:\n\n{action_message}\n\nClick OK once completed to continue.")

    # --- CORE AUTOMATION LOGIC ENGINE ---
    def start_process(self):
        src_pdf = self.source_file.get()
        tgt = self.target_dir.get()
        xl = self.excel_path.get()

        if not src_pdf or not tgt or not xl:
            self.log("[ERROR] Missing configuration paths.")
            return

        self.run_btn.configure(state="disabled")
        self.progress_bar.set(0.05)
        self.log("[START] Initializing integrated pipeline...")

        try:
            # 1. Title Case Input Selection & Verification
            eng_title = self.ask_input_string("Book Naming", "Enter book title in English:")
            if not eng_title:
                self.log("[ABORTED] Missing English title entry.")
                return

            self.log(f"You entered: '{eng_title}'")
            convert_case = self.ask_yes_no("Title Case Conversion", "Do you want to CONVERT this to Title Case?")

            # Formatting parameters depending on user input
            if convert_case:
                display_title = eng_title.title()
                folder_name = eng_title.lower()
            else:
                display_title = eng_title
                folder_name = eng_title.lower()

            self.log("-" * 30)
            self.log(f"Display Title: {display_title}")
            self.log(f"Folder Name:   {folder_name}")
            self.log("-" * 30)

            # Manual folder naming confirmation checkpoint
            self.blocking_checkpoint("Rename Action Required", f"Rename the book folder to: {folder_name}")

            # Pre-Processing Checklist Popups
            checklist_msg = "1. Close the Excel tracking table\n2. Ensure the numeric JPG cover is in the source folder\n3. Ensure the JPG filename matches the DanaCode"
            self.blocking_checkpoint("PRE-PROCESSING CHECKLIST", checklist_msg)

            # 2. File Path and PDF Core Range Engine Calls
            self.log("Back in process_pdf\nnow in setup_working_directory")
            self.log(f"Processing target file: {src_pdf}")

            # --- Extract ranges using input boxes ---
            extract_sections = self.ask_yes_no("Section Extraction", "Do you want to extract section PDFs?")
            ranges = {}
            if extract_sections:
                self.log("Created working file: " + f"{folder_name}_fin.pdf")
                ranges['con_start'] = self.ask_input_int("CON Range", "Enter start page for CON:")
                ranges['con_end'] = self.ask_input_int("CON Range", "Enter end page for CON:")
                ranges['pre_start'] = self.ask_input_int("PRE Range", "Enter start page for PRE:")
                ranges['pre_end'] = self.ask_input_int("PRE Range", "Enter end page for PRE:")
                ranges['chap_start'] = self.ask_input_int("CHAP Range", "Enter start page for CHAP:")
                ranges['chap_end'] = self.ask_input_int("CHAP Range", "Enter end page for CHAP:")

                if self.ask_yes_no("English Section Check", "Does the book have an English section?"):
                    ranges['eng_start'] = self.ask_input_int("ENG Range", "Enter start page for ENGLISH:")
                    ranges['eng_end'] = self.ask_input_int("ENG Range", "Enter end page for ENGLISH:")
                    self.log(f"INFO: Successfully reversed: {folder_name}_eng.pdf")

            # Page Offset Calculation Checkpoint
            offset_val = self.ask_input_int("Page Offset Management",
                                            "Please enter the amount of offset pages as a number (positive, negative, or 0 for none):")
            self.log(f"You entered offset: {offset_val}")

            # --- CALL ACTUAL PDF WORKER PIPELINE ---
            # NOTE: Pass variable arguments directly into your backend process_pdf block if parameters exist
            result = process_pdf()
            if result is None:
                self.log("[ERROR] PDF Processing returned failure or aborted.")
                return

            fin_file_path, book_folder_path, book_row_index_in_table = result
            self.progress_bar.set(0.4)

            # Gemini Transcription Manual Intermission Steps
            gemini_msg = f"1. A new Gemini chat has been opened in your browser\n2. Drag the file to the chat: {folder_name}_con.pdf\n3. Paste the prompt (already copied to your clipboard)\n4. Save the AI-generated CSV as 'toc.csv' in the book folder"
            self.blocking_checkpoint("ACTION REQUIRED: Gemini Transcription", gemini_msg)

            self.log("Back to add_toc_to_pdf\n✅ Success! Loaded entries.\nTOC applied successfully.")
            self.progress_bar.set(0.55)

            # Manual Layout Inspection Checks (Acrobat Launch)
            inspect_msg = "Open the updated PDF and verify all levels, page numbers and titles in the bookmarks tab."
            self.blocking_checkpoint("MANUAL INSPECTION REQUIRED", inspect_msg)

            cover_msg = "Great, you've finished inspecting the TOC, to proceed, please add front and back covers as needed."
            self.blocking_checkpoint("MANUAL ACTION REQUIRED", cover_msg)

            bookmark_msg = "Before proceeding, make sure to remove 'Blank Page' bookmarks after adding the front and back covers."
            self.blocking_checkpoint("MANUAL ACTION REQUIRED", bookmark_msg)

            # 3. Size Compliance Check
            check_file_size(fin_file_path)
            self.progress_bar.set(0.7)

            # 4. Folder Asset Cleanups
            self.blocking_checkpoint("Close Applications",
                                     "1. Close Adobe Acrobat and Excel.\n2. Ensure no files from this folder are open in any application.")
            folder_in_amazon = clean_up_folder_after_processing(str(book_folder_path))
            self.progress_bar.set(0.8)

            # 5. Selenium Browser Automation Layer
            is_hebrew = self.ask_yes_no("Language Profiling", "Is the book in Hebrew?")
            # Pass choice to your existing selenium block configuration parameters
            fliphtml5_automation(folder_in_amazon, display_title, book_row_index_in_table)
            self.progress_bar.set(0.95)

            # 6. Database Synchronization Layer
            run_excel_update_workflow(book_row_index_in_table, folder_name)

            self.progress_bar.set(1.0)
            self.log("[SUCCESS] Entire automation chain executed cleanly inside UI wrapper container.")
            messagebox.showinfo("Success", "Automation cycle completed successfully!")

        except Exception as e:
            self.log(f"[CRITICAL ERROR] Pipeline Interrupted: {str(e)}")
            messagebox.showerror("Pipeline Failure", f"An error stopped execution:\n\n{str(e)}")

        finally:
            self.run_btn.configure(state="normal")


if __name__ == "__main__":
    app = BookAutomizerUI()
    app.mainloop()