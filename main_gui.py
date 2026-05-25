import os
import sys
import shutil
from pathlib import Path
import pandas as pd
import customtkinter as ctk

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
        self.source_dir = ctk.StringVar()
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

        # 1. Source Folder Row
        self.create_file_row("Source PDF Folder:", self.source_dir, self.browse_source, 0)
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
        path = ctk.filedialog.askdirectory(title="Select Folder Containing Source PDFs")
        if path: self.source_dir.set(path)

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

    # --- CORE AUTOMATION CORE LOGIC ---
    def start_process(self):
        src = self.source_dir.get()
        tgt = self.target_dir.get()
        xl = self.excel_path.get()

        # Input Validation
        if not src or not tgt or not xl:
            self.log("[ERROR] Missing configuration paths.")
            return

        self.run_btn.configure(state="disabled")
        self.progress_bar.set(0.1)
        self.log("[START] Initializing integrated pipeline...")

        try:
            # 1. Title Normalization & Formatting Logic
            # Uses your original interactive name formatting engine logic
            self.log("[INFO] Normalizing book titles and naming structures...")
            # Note: Since normalize_book_title reads from standard CLI input, you can pass titles via
            # modifying its behavior or wrapping it. Here we trace using your metadata tool safely:
            raw_folder_name = Path(src).name
            book_titles = get_book_metadata(raw_folder_name)  # Safely extracts title case + web safe formats

            self.log(f"[INFO] Normalized Display Title: {book_titles['display_title']}")
            self.log(f"[INFO] Safe Folder Name assigned: {book_titles['folder_name']}")
            self.progress_bar.set(0.3)

            # 2. Advanced Book Splitting & TOC Bookmarking Engine
            self.log("[INFO] Executing PyPDF2 engine processing loop...")
            # We call your actual processing function which orchestrates the underlying heavy lifting
            result = process_pdf()  # Inherits full CLI step handling for ranges, offsets, and splits

            if result is None:
                self.log("[ERROR] PDF Processing returned failure or aborted.")
                self.run_btn.configure(state="normal")
                return

            fin_file_path, book_folder_path, book_row_index_in_table = result
            self.progress_bar.set(0.5)

            # 3. File System Integrity Verifications
            self.log("[INFO] Auditing document generation size constraints...")
            check_file_size(fin_file_path)

            # 4. Cleanup & Move to Amazon Workspace Folder
            self.log("[INFO] Cleaning workspace and migrating archives to target...")
            folder_in_amazon = clean_up_folder_after_processing(str(book_folder_path))
            self.progress_bar.set(0.7)

            # 5. Selenium Browser Script Upload & Design Mapping
            self.log("[INFO] Starting FlipHTML5 browser engine script...")
            fliphtml5_automation(folder_in_amazon, book_titles['display_title'], book_row_index_in_table)
            self.progress_bar.set(0.9)

            # 6. Synchronizing Custom Database Ledger Updates
            self.log("[INFO] Updating database Excel ledger columns via workflow engine...")
            run_excel_update_workflow(book_row_index_in_table, book_titles['folder_name'])

            self.progress_bar.set(1.0)
            self.log("[SUCCESS] Entire automation chain executed with zero errors.")

        except Exception as e:
            self.log(f"[CRITICAL ERROR] Pipeline Interrupted: {str(e)}")

        finally:
            self.run_btn.configure(state="normal")


if __name__ == "__main__":
    app = BookAutomizerUI()
    app.mainloop()