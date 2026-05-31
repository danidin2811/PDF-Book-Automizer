import os
import sys
import shutil
import threading
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

ctk.set_appearance_mode("System")


class BookAutomizerUI(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window Configuration
        self.title("YBZ Institute - PDF Book Automizer")
        self.geometry("780x650")
        self.resizable(False, False)

        # State Variables
        self.source_file = ctk.StringVar()
        self.target_dir = ctk.StringVar()
        self.excel_path = ctk.StringVar()
        self.generated_folder_var = ctk.StringVar(value="[Run pipeline to generate]")

        # Asynchronous Flow Context Controls
        self.is_canceling = False
        self.checkpoint_event = threading.Event()
        self.current_dialog_response = None

        self.setup_ui()

    def setup_ui(self):
        # --- TITLE BANNER ---
        self.title_label = ctk.CTkLabel(
            self,
            text="PDF Book Automizer",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        self.title_label.pack(pady=(15, 5))

        self.subtitle_label = ctk.CTkLabel(
            self,
            text="Automated processing, tracking, and normalization engine",
            font=ctk.CTkFont(size=12, slant="italic")
        )
        self.subtitle_label.pack(pady=(0, 15))

        # --- SELECTION FORM ---
        self.form_frame = ctk.CTkFrame(self)
        self.form_frame.pack(pady=5, padx=30, fill="x")

        self.create_file_row("Source PDF File:", self.source_file, self.browse_source, 0)
        self.create_file_row("Target Archive Folder:", self.target_dir, self.browse_target, 1)
        self.target_dir.set(str(READY_TO_UPLOAD_TO_AMAZON_FOLDER))
        self.create_file_row("Excel Tracker File:", self.excel_path, self.browse_excel, 2)
        self.excel_path.set(str(BOOK_TRACKER_EXCEL_FILE_PATH))

        # Copyable Generated Folder Name Row
        lbl = ctk.CTkLabel(self.form_frame, text="Generated Folder Name:", anchor="w", width=150)
        lbl.grid(row=3, column=0, padx=10, pady=10, sticky="w")

        self.folder_entry = ctk.CTkEntry(
            self.form_frame,
            textvariable=self.generated_folder_var,
            width=400,
            fg_color="#2A2A2A",
            state="readonly"
        )
        self.folder_entry.grid(row=3, column=1, padx=10, pady=10)

        copy_btn = ctk.CTkButton(self.form_frame, text="Copy Text", width=80, command=self.copy_folder_to_clipboard)
        copy_btn.grid(row=3, column=2, padx=10, pady=10)

        # --- LIVE INTERACTIVE OVERLAY PANEL ---
        # This component replaces blocking system modal alerts so the Stop button stays alert
        self.overlay_frame = ctk.CTkFrame(self, fg_color="#1E1E1E", border_width=1, border_color="#333333")
        self.overlay_frame.pack(pady=10, padx=30, fill="x")

        self.overlay_title = ctk.CTkLabel(self.overlay_frame, text="Pipeline Status: Idle",
                                          font=ctk.CTkFont(weight="bold", size=13))
        self.overlay_title.pack(pady=(8, 4), padx=15, anchor="w")

        self.overlay_msg = ctk.CTkLabel(self.overlay_frame,
                                        text="Start the operation pipeline to stream live checkpoints.", justify="left",
                                        anchor="w", font=ctk.CTkFont(size=12))
        self.overlay_msg.pack(pady=(0, 8), padx=15, fill="x")

        self.overlay_btn_frame = ctk.CTkFrame(self.overlay_frame, fg_color="transparent")
        self.overlay_btn_frame.pack(pady=(0, 8), padx=15, fill="x")

        self.confirm_btn = ctk.CTkButton(self.overlay_btn_frame, text="Confirm Execution Step", width=160,
                                         command=self._release_checkpoint)
        self.confirm_btn.pack(side="left", padx=(0, 10))
        self.confirm_btn.configure(state="disabled")

        # --- ACTIONS & PROGRESS ---
        self.action_frame = ctk.CTkFrame(self)
        self.action_frame.pack(pady=10, padx=30, fill="x")

        self.run_btn = ctk.CTkButton(
            self.action_frame,
            text="Start Automation Process",
            command=self.start_process,
            font=ctk.CTkFont(size=14, weight="bold"),
            height=40
        )
        self.run_btn.pack(side="left", padx=(0, 20))

        self.progress_bar = ctk.CTkProgressBar(self.action_frame, width=420)
        self.progress_bar.set(0)
        self.progress_bar.pack(side="right", pady=15)

        # --- LIVE LOGGING CONSOLE ---
        self.log_box = ctk.CTkTextbox(self, height=130, width=720, font=ctk.CTkFont(family="Consolas", size=11))
        self.log_box.pack(pady=(5, 15), padx=30)
        self.log_box.configure(state="disabled")

    def create_file_row(self, label_text, text_var, browse_cmd, row_idx):
        lbl = ctk.CTkLabel(self.form_frame, text=label_text, anchor="w", width=150)
        lbl.grid(row=row_idx, column=0, padx=10, pady=8, sticky="w")
        entry = ctk.CTkEntry(self.form_frame, textvariable=text_var, width=400)
        entry.grid(row=row_idx, column=1, padx=10, pady=8)
        btn = ctk.CTkButton(self.form_frame, text="Browse", width=80, command=browse_cmd)
        btn.grid(row=row_idx, column=2, padx=10, pady=8)

    def browse_source(self):
        path = ctk.filedialog.askopenfilename(title="Select Source PDF File", filetypes=[("PDF Files", "*.pdf")])
        if path: self.source_file.set(path)

    def browse_target(self):
        path = ctk.filedialog.askdirectory(title="Select Target Archive Folder")
        if path: self.target_dir.set(path)

    def browse_excel(self):
        path = ctk.filedialog.askopenfilename(title="Select Excel Tracker File", filetypes=[("Excel Files", "*.xlsx")])
        if path: self.excel_path.set(path)

    def copy_folder_to_clipboard(self):
        self.clipboard_clear()
        self.clipboard_append(self.generated_folder_var.get())
        self.log("[INFO] Folder name copied to clipboard!")

    def log(self, message):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"{message}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        self.update_idletasks()

    # --- ASYNCHRONOUS UI SAFE POPUP WRAPPERS ---
    def async_ask_yes_no(self, title, message):
        self.checkpoint_event.clear()
        self.overlay_title.configure(text=f"❓ Setup Decision: {title}", text_color="#3498DB")
        self.overlay_msg.configure(text=message)

        # Clear existing layout components inside the overlay button matrix
        for widget in self.overlay_btn_frame.winfo_children():
            widget.pack_forget()

        btn_yes = ctk.CTkButton(self.overlay_btn_frame, text="Yes", width=100, fg_color="#2ECC71",
                                hover_color="#27AE60", command=lambda: self._resolve_async_dialog(True))
        btn_yes.pack(side="left", padx=(0, 10))
        btn_no = ctk.CTkButton(self.overlay_btn_frame, text="No", width=100, fg_color="#E74C3C", hover_color="#C0392B",
                               command=lambda: self._resolve_async_dialog(False))
        btn_no.pack(side="left")

        self.checkpoint_event.wait()
        return self.current_dialog_response

    def async_ask_string(self, title, prompt):
        self.checkpoint_event.clear()
        self.overlay_title.configure(text=f"✏️ Input Required: {title}", text_color="#F1C40F")
        self.overlay_msg.configure(text=prompt)

        for widget in self.overlay_btn_frame.winfo_children():
            widget.pack_forget()

        entry_val = ctk.CTkEntry(self.overlay_btn_frame, width=300)
        entry_val.pack(side="left", padx=(0, 10))
        entry_val.focus()

        btn_submit = ctk.CTkButton(self.overlay_btn_frame, text="Submit", width=100,
                                   command=lambda: self._resolve_async_dialog(entry_val.get()))
        btn_submit.pack(side="left")

        self.checkpoint_event.wait()
        return self.current_dialog_response

    def async_ask_int(self, title, prompt):
        while True:
            val = self.async_ask_string(title, prompt)
            if val is None or val == "":
                return None
            if val.strip().lstrip('-').isdigit():
                return int(val.strip())
            self.log("[WARN] Non-integer detected. Re-prompting input variables...")

    def async_blocking_checkpoint(self, title, action_message):
        self.checkpoint_event.clear()
        self.overlay_title.configure(text=f"⚠️ Action Required: {title}", text_color="#E67E22")
        self.overlay_msg.configure(text=action_message)

        for widget in self.overlay_btn_frame.winfo_children():
            widget.pack_forget()

        btn_confirm = ctk.CTkButton(self.overlay_btn_frame, text="I Have Completed This Step", width=220,
                                    fg_color="#2980B9", hover_color="#2471A3", command=self._release_checkpoint)
        btn_confirm.pack(side="left")

        self.log(f"[ACTION REQUIRED] {title} - Complete step in background.")
        self.checkpoint_event.wait()

    def _resolve_async_dialog(self, response_value):
        self.current_dialog_response = response_value
        self._release_checkpoint()

    def _release_checkpoint(self):
        self.checkpoint_event.set()

    def stop_process(self):
        """Immediately flag cancellation and break threads waiting on user checkpoints."""
        self.is_canceling = True
        self.log("[CANCELING] Cancel click registered. Freeing interface pipelines immediately...")
        self.run_btn.configure(text="Canceling...", state="disabled")
        # Instantly unblocks the background worker context if it was waiting on a checkpoint confirmation
        self._release_checkpoint()

    def start_process(self):
        src_pdf = self.source_file.get()
        tgt = self.target_dir.get()
        xl = self.excel_path.get()

        if not src_pdf or not tgt or not xl:
            self.log("[ERROR] Missing configuration paths.")
            return

        self.is_canceling = False
        self.run_btn.configure(
            text="Stop Process",
            fg_color="#C0392B",
            hover_color="#922B21",
            command=self.stop_process
        )

        self.progress_bar.set(0.05)

        worker_thread = threading.Thread(target=self._run_pipeline_worker, daemon=True)
        worker_thread.start()

    # --- PIPELINE WORKER ENGINE ---
    def _run_pipeline_worker(self):
        self.log("[START] Initializing asynchronous integrated pipeline...")

        try:
            # 1. Title Case Input Selection & Verification
            eng_title = self.async_ask_string("Book Naming", "Enter book title in English:")
            if not eng_title or self.is_canceling:
                self.log("[ABORTED] Pipeline execution stopped.")
                return

            convert_case = self.async_ask_yes_no("Title Case Conversion", "Do you want to CONVERT this to Title Case?")
            if self.is_canceling: return

            if convert_case:
                display_title = eng_title.title()
                folder_name = eng_title.lower()
            else:
                display_title = eng_title
                folder_name = eng_title.lower()

            self.generated_folder_var.set(folder_name)
            self.copy_folder_to_clipboard()

            self.log("-" * 30)
            self.log(f"Display Title: {display_title}")
            self.log(f"Folder Name:   {folder_name} [COPIED TO CLIPBOARD]")
            self.log("-" * 30)

            if self.is_canceling: return

            self.async_blocking_checkpoint(
                "Rename Action Required",
                f"Rename the book folder to: '{folder_name}'\n(This text is ready on your system clipboard!)"
            )
            if self.is_canceling: return

            checklist_msg = "1. Close the Excel tracking table\n2. Ensure the numeric JPG cover is in the source folder\n3. Ensure the JPG filename matches the DanaCode"
            self.async_blocking_checkpoint("PRE-PROCESSING CHECKLIST", checklist_msg)
            if self.is_canceling: return

            # 2. PDF Processing & Splits
            self.log("Back in process_pdf\nnow in setup_working_directory")

            extract_sections = self.async_ask_yes_no("Section Extraction", "Do you want to extract section PDFs?")
            if self.is_canceling: return

            ranges = {}
            if extract_sections:
                self.log(f"Created working file: {folder_name}_fin.pdf")
                ranges['con_start'] = self.async_ask_int("CON Range", "Enter start page for CON:")
                if self.is_canceling: return
                ranges['con_end'] = self.async_ask_int("CON Range", "Enter end page for CON:")
                if self.is_canceling: return
                ranges['pre_start'] = self.async_ask_int("PRE Range", "Enter start page for PRE:")
                if self.is_canceling: return
                ranges['pre_end'] = self.async_ask_int("PRE Range", "Enter end page for PRE:")
                if self.is_canceling: return
                ranges['chap_start'] = self.async_ask_int("CHAP Range", "Enter start page for CHAP:")
                if self.is_canceling: return
                ranges['chap_end'] = self.async_ask_int("CHAP Range", "Enter end page for CHAP:")
                if self.is_canceling: return

                if self.async_ask_yes_no("English Section Check", "Does the book have an English section?"):
                    if self.is_canceling: return
                    ranges['eng_start'] = self.async_ask_int("ENG Range", "Enter start page for ENGLISH:")
                    if self.is_canceling: return
                    ranges['eng_end'] = self.async_ask_int("ENG Range", "Enter end page for ENGLISH:")
                    if self.is_canceling: return

            offset_val = self.async_ask_int("Page Offset Management", "Please enter the amount of offset pages:")
            if self.is_canceling: return

            # Backend worker processing
            result = process_pdf()
            if result is None:
                self.log("[ERROR] PDF Processing returned failure or aborted.")
                return

            fin_file_path, book_folder_path, book_row_index_in_table = result
            self.progress_bar.set(0.4)
            if self.is_canceling: return

            gemini_msg = f"1. A new Gemini chat has been opened\n2. Drag the file: {folder_name}_con.pdf\n3. Paste the prompt\n4. Save CSV as 'toc.csv'"
            self.async_blocking_checkpoint("Gemini Transcription", gemini_msg)
            if self.is_canceling: return

            self.log("Back to add_toc_to_pdf\nTOC applied successfully.")
            self.progress_bar.set(0.55)

            self.async_blocking_checkpoint("MANUAL INSPECTION REQUIRED", "Verify bookmarks tab in Adobe Acrobat.")
            if self.is_canceling: return

            self.async_blocking_checkpoint("MANUAL ACTION REQUIRED", "Add front and back covers as needed.")
            if self.is_canceling: return

            self.async_blocking_checkpoint("MANUAL ACTION REQUIRED", "Remove 'Blank Page' bookmarks.")
            if self.is_canceling: return

            # 3. File System Checks
            check_file_size(fin_file_path)
            self.progress_bar.set(0.7)
            if self.is_canceling: return

            # 4. Folder Asset Cleanups
            self.async_blocking_checkpoint("Close Applications", "Close Adobe Acrobat and Excel.")
            if self.is_canceling: return

            folder_in_amazon = clean_up_folder_after_processing(str(book_folder_path))
            self.progress_bar.set(0.8)
            if self.is_canceling: return

            # 5. Selenium Browser Automation Layer
            is_hebrew = self.async_ask_yes_no("Language Profiling", "Is the book in Hebrew?")
            if self.is_canceling: return

            fliphtml5_automation(folder_in_amazon, display_title, book_row_index_in_table)
            self.progress_bar.set(0.95)
            if self.is_canceling: return

            # 6. Database Serialization
            run_excel_update_workflow(book_row_index_in_table, folder_name)

            self.progress_bar.set(1.0)
            self.log("[SUCCESS] Entire automation chain executed cleanly.")

            self.overlay_title.configure(text="Pipeline Complete", text_color="#2ECC71")
            self.overlay_msg.configure(text="All automation cycles have wrapped successfully.")
            messagebox.showinfo("Success", "Automation cycle completed successfully!")

        except Exception as e:
            self.log(f"[CRITICAL ERROR] Pipeline Interrupted: {str(e)}")
            self.overlay_title.configure(text="Pipeline Failed", text_color="#E74C3C")
            self.overlay_msg.configure(text=f"Interrupted with error: {str(e)}")

        finally:
            # Revert UI state safely
            if self.is_canceling:
                self.log("[ABORTED] Process safely torn down by user cancellation.")
                self.overlay_title.configure(text="Pipeline Aborted", text_color="#E74C3C")
                self.overlay_msg.configure(text="The user triggered an abort request. State changes reverted.")

            for widget in self.overlay_btn_frame.winfo_children():
                widget.pack_forget()

            self.run_btn.configure(
                text="Start Automation Process",
                fg_color=ctk.ThemeManager.theme["CTkButton"]["fg_color"],
                hover_color=ctk.ThemeManager.theme["CTkButton"]["hover_color"],
                state="normal",
                command=self.start_process
            )


if __name__ == "__main__":
    app = BookAutomizerUI()
    app.mainloop()