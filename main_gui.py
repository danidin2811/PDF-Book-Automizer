import re
import os
import sys
import shutil
import threading
from pathlib import Path
import pandas as pd
import customtkinter as ctk
from tkinter import messagebox

# --- BACKEND MODULE IMPORTS ---
from src.logic.file_operations import check_file_size, validate_pdf_path
from src.constants import READY_TO_UPLOAD_TO_AMAZON_FOLDER, BOOK_TRACKER_EXCEL_FILE_PATH
from utils.norm_book_title import normalize_book_title, get_book_metadata
from src.logic.pdf_processor import process_pdf
from src.logic.file_operations import check_file_size
from src.logic.system_tools import clean_up_folder_after_processing
from src.fliphtml5.flip_html_automation import fliphtml5_automation
from src.logic.excel_tools import run_excel_update_workflow

ctk.set_appearance_mode("System")

class CustomStdoutStream:
    """Intercepts terminal print data signals and safely routes them to the interactive UI console box."""
    def __init__(self, ui_instance):
        self.ui = ui_instance
        self.terminal = sys.stdout

    def write(self, message):
        self.terminal.write(message) # Keep printing to standard terminal output
        if message.strip():
            # Thread-safe scheduling to append text messages to your logging box
            self.ui.after(0, lambda: self.ui.log(message.strip()))

    def flush(self):
        self.terminal.flush()

class BookAutomizerUI(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window Configuration
        self.title("YBZ Institute - PDF Book Automizer")
        self.geometry("780x640")
        self.resizable(False, False)

        # State Variables
        self.source_file = ctk.StringVar()
        self.target_dir = ctk.StringVar()
        self.excel_path = ctk.StringVar()

        # Asynchronous Flow Context Controls
        self.is_canceling = False
        self.checkpoint_event = threading.Event()
        self.current_dialog_response = None

        self.setup_ui()
        self.force_english_keyboard_layout()

        sys.stdout = CustomStdoutStream(self)

    def force_english_keyboard_layout(self):
        """Forces the current active window thread to switch to the English keyboard layout."""
        if sys.platform == "win32":
            import ctypes
            # 0x04090409 is the standard load identifier for English (US)
            # 1 indicates that the layout should be activated immediately for the current thread
            try:
                ctypes.windll.user32.ActivateKeyboardLayout(0x04090409, 1)
                print("[SYSTEM] Keyboard layout successfully forced to English.")
            except Exception as e:
                print(f"[SYSTEM ERROR] Could not automatically shift keyboard layout: {e}")

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

        # --- LIVE INTERACTIVE OVERLAY PANEL ---
        self.overlay_frame = ctk.CTkFrame(self, fg_color="#1E1E1E", border_width=1, border_color="#333333")
        self.overlay_frame.pack(pady=10, padx=30, fill="x")

        self.overlay_title = ctk.CTkLabel(self.overlay_frame, text="Pipeline Status: Idle",
                                          font=ctk.CTkFont(weight="bold", size=13))
        self.overlay_title.pack(pady=(8, 4), padx=15, anchor="w")

        self.overlay_msg_box = ctk.CTkTextbox(self.overlay_frame, height=75, fg_color="#1E1E1E", text_color="#FFFFFF",
                                              activate_scrollbars=False, wrap="word", font=ctk.CTkFont(size=12))
        self.overlay_msg_box.pack(pady=(0, 5), padx=15, fill="x")
        self.overlay_msg_box.insert("1.0", "Start the operation pipeline to stream live checkpoints.")
        self.overlay_msg_box.configure(state="disabled")

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
        self.log_box = ctk.CTkTextbox(self, height=120, width=720, font=ctk.CTkFont(family="Consolas", size=11))
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
        if path:
            self.source_file.set(path.strip('"'))

    def browse_target(self):
        path = ctk.filedialog.askdirectory(title="Select Target Archive Folder")
        if path: self.target_dir.set(path)

    def browse_excel(self):
        path = ctk.filedialog.askopenfilename(title="Select Excel Tracker File", filetypes=[("Excel Files", "*.xlsx")])
        if path: self.excel_path.set(path)

    def update_overlay_text(self, text_string):
        self.overlay_msg_box.configure(state="normal")
        self.overlay_msg_box.delete("1.0", "end")
        self.overlay_msg_box.insert("1.0", text_string)
        self.overlay_msg_box.configure(state="disabled")

    def log(self, message):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"{message}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        self.update_idletasks()

    # --- ASYNCHRONOUS UI SAFE POPUP WRAPPERS ---
    def async_ask_yes_no(self, title, message):
        self.checkpoint_event.clear()
        self.current_dialog_response = None  # FIX: Reset stale state data

        self.overlay_title.configure(text=f"❓ Setup Decision: {title}", text_color="#3498DB")
        self.update_overlay_text(message)

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
        self.current_dialog_response = None

        self.overlay_title.configure(text=f"✏️ Input Required: {title}", text_color="#F1C40F")
        self.update_overlay_text(prompt)

        for widget in self.overlay_btn_frame.winfo_children():
            widget.pack_forget()

        entry_val = ctk.CTkEntry(self.overlay_btn_frame, width=300)
        entry_val.pack(side="left", padx=(0, 10))
        entry_val.focus()

        # Local element binding (This is perfectly clean and stays safe)
        entry_val.bind("<Return>", lambda event: self._resolve_async_dialog(entry_val.get()))

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
            if str(val).strip().lstrip('-').isdigit():
                return int(str(val).strip())
            self.log("[WARN] Non-integer detected. Re-prompting input variables...")

    def async_ask_ranges(self, include_english=False):
        """Displays an integrated matrix inside the overlay window to handle all numbers concurrently."""
        while True:
            self.checkpoint_event.clear()
            self.current_dialog_response = None  # Reset stale state data

            self.overlay_title.configure(text="📊 Section Target Configuration Grid", text_color="#F1C40F")
            self.update_overlay_text("Input start and end page numbers for each requested subsection split.")

            for widget in self.overlay_btn_frame.winfo_children():
                widget.pack_forget()

            grid_frame = ctk.CTkFrame(self.overlay_btn_frame, fg_color="transparent")
            grid_frame.pack(fill="x", expand=True, pady=10, padx=5)

            grid_frame.columnconfigure(0, weight=2, minsize=140)
            grid_frame.columnconfigure(1, weight=1, minsize=100)
            grid_frame.columnconfigure(2, weight=1, minsize=100)

            headers = ["Section Prefix", "Start Page", "End Page"]
            for col_idx, text in enumerate(headers):
                anchor_dir = "w" if col_idx == 0 else "center"
                lbl = ctk.CTkLabel(
                    grid_frame,
                    text=text,
                    font=ctk.CTkFont(weight="bold", size=12),
                    text_color="#AAAAAA",
                    anchor=anchor_dir
                )
                if col_idx == 0:
                    lbl.grid(row=0, column=col_idx, padx=10, pady=(0, 8), sticky="w")
                else:
                    lbl.grid(row=0, column=col_idx, padx=10, pady=(0, 8))

            sections = ["CON", "PRE", "CHAP"]
            if include_english:
                sections.append("ENG")

            entries = {}
            for row_idx, sec in enumerate(sections, start=1):
                lbl = ctk.CTkLabel(
                    grid_frame,
                    text=f"{sec} Section:",
                    anchor="w",
                    font=ctk.CTkFont(size=13, weight="bold"),
                    text_color="#FFFFFF"
                )
                lbl.grid(row=row_idx, column=0, padx=10, pady=6, sticky="w")

                ent_start = ctk.CTkEntry(grid_frame, width=90, placeholder_text="0", justify="center")
                ent_start.grid(row=row_idx, column=1, padx=10, pady=6)

                ent_end = ctk.CTkEntry(grid_frame, width=90, placeholder_text="0", justify="center")
                ent_end.grid(row=row_idx, column=2, padx=10, pady=6)

                entries[sec.lower()] = (ent_start, ent_end)

            # --- CRITICAL FIX: The validation code and submit controls must be OUTSIDE the section loop ---
            def submit_range_validation(event=None):
                extracted_output = {}
                try:
                    for sec_key, (start_w, end_w) in entries.items():
                        s_val, e_val = start_w.get().strip(), end_w.get().strip()
                        if not s_val.isdigit() or not e_val.isdigit():
                            raise ValueError(f"All values for {sec_key.upper()} must be valid positive integers.")
                        extracted_output[f"{sec_key}_start"] = int(s_val)
                        extracted_output[f"{sec_key}_end"] = int(e_val)

                    self._resolve_async_dialog(extracted_output)
                except ValueError as err:
                    self.log(f"[WARN] Input Verification Error: {str(err)}")

            # Bind the Enter key to all input fields in the range matrix grid
            for ent_start, ent_end in entries.values():
                ent_start.bind("<Return>", submit_range_validation)
                ent_end.bind("<Return>", submit_range_validation)

            # Draw the button once at the bottom after the grid has completely built
            btn_submit = ctk.CTkButton(
                self.overlay_btn_frame,
                text="Submit Section Ranges",
                width=240,
                height=32,
                font=ctk.CTkFont(weight="bold"),
                fg_color="#2ECC71",
                hover_color="#27AE60",
                command=submit_range_validation
            )
            btn_submit.pack(pady=(12, 5))

            # Block worker thread execution only AFTER all sections have successfully loaded inside the frame matrix
            self.checkpoint_event.wait()
            if self.is_canceling or self.current_dialog_response is not None:
                return self.current_dialog_response

    def async_blocking_checkpoint(self, title, action_message):
        self.checkpoint_event.clear()
        self.current_dialog_response = None

        self.overlay_title.configure(text=f"⚠️ Action Required: {title}", text_color="#E67E22")
        self.update_overlay_text(action_message)

        for widget in self.overlay_btn_frame.winfo_children():
            widget.pack_forget()

        btn_confirm = ctk.CTkButton(self.overlay_btn_frame, text="I Have Completed This Step", width=220,
                                    fg_color="#2980B9", hover_color="#2471A3", command=self._release_checkpoint)
        btn_confirm.pack(side="left")

        # FIX: Focus the button elements directly so standard OS event-loops allow
        # pressing spacebar or Enter to instantly proceed WITHOUT breaking application-wide keymaps
        btn_confirm.focus_set()

        self.log(f"[ACTION REQUIRED] {title} - Complete step in background.")
        self.checkpoint_event.wait()

    def _resolve_async_dialog(self, response_value):
        self.current_dialog_response = response_value
        self._release_checkpoint()

    def _release_checkpoint(self):
        self.checkpoint_event.set()

    def stop_process(self):
        self.is_canceling = True
        self.log("[SHUTDOWN] Stop execution request received. Closing automation manager program...")
        self._release_checkpoint()
        self.destroy()
        sys.exit(0)

    def start_process(self):
        src_pdf = self.source_file.get()
        tgt = self.target_dir.get()
        xl = self.excel_path.get()

        if not src_pdf or not tgt or not xl:
            self.log("[ERROR] Missing configuration paths. Fill out all source/target parameters.")
            messagebox.showerror("Configuration Error", "Please fill out all file and folder fields before proceeding.")
            return

        is_valid, error_msg = validate_pdf_path(src_pdf)
        if not is_valid:
            self.log(f"[CRITICAL PATH ERROR] {error_msg}")
            messagebox.showerror(
                "Invalid PDF Source Path",
                f"Validation failed:\n{error_msg}\n\nPlease check your input file path or try re-browsing for the file."
            )
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
            while True:
                eng_title = self.async_ask_string("Book Naming", "Enter book title in English:")

                # If user closes or hits cancel, abort the pipeline
                if self.is_canceling or eng_title is None:
                    self.log("[ABORTED] Pipeline execution stopped.")
                    return

                # FIX: Regex ensuring the string contains ONLY English characters (a-z, A-Z), numbers (0-9), and spaces
                # it also checks that it isn't just an empty string of pure spaces
                if eng_title.strip() != "" and re.match(r"^[a-zA-Z0-9 ]+$", eng_title):
                    break  # Valid English input received, break the validation loop

                # Warn the user and re-loop if invalid characters are passed
                self.log(
                    "[WARN] Invalid title syntax. Title must contain ONLY English characters (A-Z, a-z) or numbers (0-9).")
                messagebox.showwarning(
                    "Validation Error",
                    "The book title must be written in English containing valid alphanumeric characters, and cannot be left blank."
                )

                # Continue with your existing case conversion selection logic...
            convert_case = self.async_ask_yes_no("Title Case Conversion", "Do you want to CONVERT this to Title Case?")
            if self.is_canceling: return

            if convert_case:
                display_title = eng_title.title()
                folder_name = eng_title.lower()
            else:
                display_title = eng_title
                folder_name = eng_title.lower()

            self.clipboard_clear()
            self.clipboard_append(folder_name)

            self.log("-" * 30)
            self.log(f"Display Title: {display_title}")
            self.log(f"Folder Name:   {folder_name} [COPIED TO CLIPBOARD]")
            self.log("-" * 30)

            if self.is_canceling: return

            self.async_blocking_checkpoint(
                "Rename Action Required",
                f"Rename the book folder to: '{folder_name}'\n(This text has been automatically copied to your system clipboard!)"
            )
            if self.is_canceling: return

            extract_sections = self.async_ask_yes_no("Section Extraction", "Do you want to extract section PDFs?")
            if self.is_canceling: return

            ranges = {}
            if extract_sections:
                # Check for an English section first to see if we should include ENG fields in our single window layout
                has_english = self.async_ask_yes_no("English Section Check", "Does the book have an English section?")
                if self.is_canceling: return

                # Query all ranges concurrently inside one unified matrix layout frame
                ranges = self.async_ask_ranges(include_english=has_english)
                if self.is_canceling or not ranges: return

            offset_val: int | None = self.async_ask_int("Page Offset Management", "Please enter the amount of offset pages:")
            if self.is_canceling: return

            # Backend worker processing
            result = process_pdf(ui=self)
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
            self.update_overlay_text("All automation cycles have wrapped successfully.")
            messagebox.showinfo("Success", "Automation cycle completed successfully!")

        except Exception as e:
            self.log(f"[CRITICAL ERROR] Pipeline Interrupted: {str(e)}")
            self.overlay_title.configure(text="Pipeline Failed", text_color="#E74C3C")
            self.update_overlay_text(f"Interrupted with error: {str(e)}")

        finally:
            if not self.is_canceling:
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