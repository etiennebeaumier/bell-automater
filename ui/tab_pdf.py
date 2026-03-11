"""PDF upload and processing tab."""

import os
from tkinter import filedialog
import customtkinter as ctk
from config import get_default_pdf_dir


class PdfTab(ctk.CTkFrame):
    def __init__(self, master, app_window, **kwargs):
        super().__init__(master, **kwargs)
        self.app_window = app_window
        self.selected_files = []
        self._is_processing = False

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # -- Header --
        header = ctk.CTkLabel(self, text="Local PDF Intake", font=ctk.CTkFont(size=15, weight="bold"))
        header.grid(row=0, column=0, padx=16, pady=(12, 2), sticky="w")

        desc = ctk.CTkLabel(self, text="Select BCECN PDFs individually or pick a folder.", font=ctk.CTkFont(size=12), text_color="gray")
        desc.grid(row=1, column=0, padx=16, pady=(0, 8), sticky="w")

        # -- File list --
        self.file_list = ctk.CTkTextbox(self, height=120, state="disabled", font=ctk.CTkFont(size=12))
        self.file_list.grid(row=2, column=0, padx=16, pady=(0, 8), sticky="nsew")

        # -- Buttons row --
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.grid(row=3, column=0, padx=16, pady=(0, 8), sticky="ew")
        btn_frame.grid_columnconfigure(1, weight=1)

        self.select_btn = ctk.CTkButton(btn_frame, text="Select Files", command=self._select_files)
        self.select_btn.grid(row=0, column=0, padx=(0, 8))

        self.select_folder_btn = ctk.CTkButton(btn_frame, text="Select Folder", command=self._select_folder)
        self.select_folder_btn.grid(row=0, column=1, padx=(0, 8))

        self.clear_btn = ctk.CTkButton(btn_frame, text="Clear", width=70, fg_color="gray", command=self._clear_files)
        self.clear_btn.grid(row=0, column=2, sticky="w")

        self.file_count_label = ctk.CTkLabel(btn_frame, text="No files selected", font=ctk.CTkFont(size=12), text_color="gray")
        self.file_count_label.grid(row=0, column=3, padx=(16, 0))

        # -- Progress bar --
        self.progress = ctk.CTkProgressBar(self)
        self.progress.grid(row=4, column=0, padx=16, pady=(0, 4), sticky="ew")
        self.progress.set(0)
        self.progress.grid_remove()

        # -- Process button --
        self.process_btn = ctk.CTkButton(
            self, text="Process Files", height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._process_files, state="disabled",
        )
        self.process_btn.grid(row=5, column=0, padx=16, pady=(4, 16), sticky="ew")

    def _select_files(self):
        initial_dir = self._get_default_pdf_dir()
        paths = filedialog.askopenfilenames(
            title="Select BCECN PDF Files",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialdir=initial_dir,
        )
        if paths:
            self.selected_files = list(paths)
            self._update_file_list()
            self._maybe_autoprocess()

    def _select_folder(self):
        folder = filedialog.askdirectory(
            title="Select Folder Containing BCECN PDFs",
            initialdir=self._get_default_pdf_dir(),
        )
        if folder:
            pdfs = sorted(
                os.path.join(folder, f)
                for f in os.listdir(folder)
                if f.lower().endswith(".pdf")
            )
            if pdfs:
                self.selected_files = pdfs
                self._update_file_list()
                self._maybe_autoprocess()

    def _clear_files(self):
        self.selected_files = []
        self._update_file_list()

    def _update_file_list(self):
        self.file_list.configure(state="normal")
        self.file_list.delete("1.0", "end")
        if self.selected_files:
            for f in self.selected_files:
                name = os.path.basename(f)
                size_mb = os.path.getsize(f) / (1024 * 1024)
                self.file_list.insert("end", f"{name}  ({size_mb:.2f} MB)\n")
            self.file_count_label.configure(text=f"{len(self.selected_files)} file(s) ready")
            self.process_btn.configure(state="normal")
        else:
            self.file_count_label.configure(text="No files selected")
            self.process_btn.configure(state="disabled")
        self.file_list.configure(state="disabled")

    def _get_default_pdf_dir(self) -> str:
        preferred = get_default_pdf_dir()
        if os.path.isdir(preferred):
            return preferred
        automation_dir = os.path.dirname(preferred)
        if os.path.isdir(automation_dir):
            return automation_dir
        return os.path.expanduser("~")

    def _process_files(self):
        if self._is_processing:
            return

        settings = self.app_window.settings
        if not self._has_valid_selected_files():
            from tkinter import messagebox
            messagebox.showerror("File Error", "Selected files must all be PDFs.")
            return

        if not settings.dry_run and not settings.is_workbook_ready():
            from tkinter import messagebox
            messagebox.showerror("Workbook Error", "Workbook is not ready. Fix the path in settings or enable dry-run mode.")
            return

        self._set_processing(True)
        self._is_processing = True
        self.progress.grid()
        self.progress.set(0)

        from ui.workers import PdfProcessWorker
        worker = PdfProcessWorker(
            app=self.winfo_toplevel(),
            pdf_paths=self.selected_files,
            master_file=settings.master_file,
            dry_run=settings.dry_run,
            avg_start_year=settings.avg_start_year,
            avg_end_year=settings.avg_end_year,
            on_progress=self._on_progress,
            on_result=self._on_result,
            on_complete=self._on_complete,
        )
        worker.start()

    def _on_progress(self, fraction):
        self.progress.set(fraction)

    def _on_result(self, message, success):
        self.app_window.results.log(message)

    def _on_complete(self, summary):
        self.app_window.results.log(summary)
        self.app_window.results.log("All selected documents were processed.")
        self._is_processing = False
        self._set_processing(False)
        self.progress.grid_remove()

    def _set_processing(self, active):
        state = "disabled" if active else "normal"
        self.process_btn.configure(state=state)
        self.select_btn.configure(state=state)
        self.select_folder_btn.configure(state=state)
        self.clear_btn.configure(state=state)

    def _has_valid_selected_files(self) -> bool:
        if not self.selected_files:
            return False
        return all(os.path.isfile(path) and path.lower().endswith(".pdf") for path in self.selected_files)

    def _maybe_autoprocess(self):
        if self._is_processing or not self._has_valid_selected_files():
            return

        settings = self.app_window.settings
        if settings.is_workbook_ready():
            self._process_files()
