"""Outlook email fetch tab."""

import customtkinter as ctk


class OutlookTab(ctk.CTkFrame):
    def __init__(self, master, app_window, config, **kwargs):
        super().__init__(master, **kwargs)
        self.app_window = app_window
        self.config = config

        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # -- Header --
        header = ctk.CTkLabel(self, text="Outlook Fetch", font=ctk.CTkFont(size=15, weight="bold"))
        header.grid(row=0, column=0, columnspan=2, padx=16, pady=(12, 2), sticky="w")

        desc = ctk.CTkLabel(self, text="Pull BCECN attachments from Outlook and process them.", font=ctk.CTkFont(size=12), text_color="gray")
        desc.grid(row=1, column=0, columnspan=2, padx=16, pady=(0, 12), sticky="w")

        # -- Email --
        ctk.CTkLabel(self, text="Email", font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=2, column=0, padx=16, pady=(4, 2), sticky="w"
        )
        self.email_var = ctk.StringVar(value=config.get("outlook_email", ""))
        ctk.CTkEntry(self, textvariable=self.email_var, placeholder_text="your.email@bell.ca").grid(
            row=3, column=0, padx=16, pady=(0, 8), sticky="ew"
        )

        # -- Server --
        ctk.CTkLabel(self, text="Server", font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=2, column=1, padx=16, pady=(4, 2), sticky="w"
        )
        self.server_var = ctk.StringVar(value=config.get("outlook_server", "outlook.office365.com"))
        ctk.CTkEntry(self, textvariable=self.server_var).grid(
            row=3, column=1, padx=16, pady=(0, 8), sticky="ew"
        )

        # -- Days back --
        ctk.CTkLabel(self, text="Days Back", font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=4, column=0, padx=16, pady=(4, 2), sticky="w"
        )
        self.days_var = ctk.StringVar(value=str(config.get("outlook_days", 7)))
        ctk.CTkEntry(self, textvariable=self.days_var, width=80).grid(
            row=5, column=0, padx=16, pady=(0, 8), sticky="w"
        )

        # -- Sender filter --
        ctk.CTkLabel(self, text="Sender Filter (optional)", font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=4, column=1, padx=16, pady=(4, 2), sticky="w"
        )
        self.sender_var = ctk.StringVar(value=config.get("bcecn_sender", ""))
        ctk.CTkEntry(self, textvariable=self.sender_var, placeholder_text="sender@bank.com").grid(
            row=5, column=1, padx=16, pady=(0, 8), sticky="ew"
        )

        # -- Auth status display --
        self.auth_frame = ctk.CTkFrame(self, fg_color=("gray90", "gray17"), corner_radius=10)
        self.auth_frame.grid(row=6, column=0, columnspan=2, padx=16, pady=(8, 4), sticky="ew")
        self.auth_frame.grid_columnconfigure(0, weight=1)

        self.auth_label = ctk.CTkLabel(
            self.auth_frame, text="Ready to connect",
            font=ctk.CTkFont(size=12), text_color="gray", wraplength=500, justify="left",
        )
        self.auth_label.grid(row=0, column=0, padx=12, pady=8, sticky="w")

        self.copy_btn = ctk.CTkButton(self.auth_frame, text="Copy Code", width=90, command=self._copy_code, state="disabled")
        self.copy_btn.grid(row=0, column=1, padx=(0, 12), pady=8)

        self._device_code = ""

        # -- Progress bar --
        self.progress = ctk.CTkProgressBar(self)
        self.progress.grid(row=7, column=0, columnspan=2, padx=16, pady=(4, 4), sticky="ew")
        self.progress.set(0)
        self.progress.grid_remove()

        # -- Fetch button --
        self.fetch_btn = ctk.CTkButton(
            self, text="Fetch and Process", height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._fetch,
        )
        self.fetch_btn.grid(row=8, column=0, columnspan=2, padx=16, pady=(4, 16), sticky="ew")

    def _copy_code(self):
        if self._device_code:
            self.winfo_toplevel().clipboard_clear()
            self.winfo_toplevel().clipboard_append(self._device_code)
            self.copy_btn.configure(text="Copied!")
            self.after(1500, lambda: self.copy_btn.configure(text="Copy Code"))

    def _fetch(self):
        email = self.email_var.get().strip()
        if not email:
            from tkinter import messagebox
            messagebox.showerror("Missing Email", "Email address is required.")
            return

        settings = self.app_window.settings
        if not settings.dry_run and not settings.is_workbook_ready():
            from tkinter import messagebox
            messagebox.showerror("Workbook Error", "Workbook is not ready. Fix the path in settings or enable dry-run mode.")
            return

        try:
            days = int(self.days_var.get())
            if days < 1:
                raise ValueError
        except ValueError:
            from tkinter import messagebox
            messagebox.showerror("Invalid Input", "Days back must be a positive integer.")
            return

        self._set_processing(True)
        self.progress.grid()
        self.progress.set(0)
        self.auth_label.configure(text="Authenticating via Microsoft...", text_color=("gray30", "gray70"))

        from ui.workers import OutlookFetchWorker
        worker = OutlookFetchWorker(
            app=self.winfo_toplevel(),
            email=email,
            server=self.server_var.get().strip(),
            days=days,
            sender=self.sender_var.get().strip(),
            master_file=settings.master_file,
            dry_run=settings.dry_run,
            avg_start_year=settings.avg_start_year,
            avg_end_year=settings.avg_end_year,
            on_auth_status=self._on_auth_status,
            on_progress=self._on_progress,
            on_result=self._on_result,
            on_complete=self._on_complete,
            on_error=self._on_error,
        )
        worker.start()

    def _on_auth_status(self, message):
        self.auth_label.configure(text=message, text_color=("gray20", "gray80"))
        # Extract device code from the message if present
        import re
        match = re.search(r"enter the code\s+(\S+)", message, re.IGNORECASE)
        if match:
            self._device_code = match.group(1)
            self.copy_btn.configure(state="normal")
        elif "code" not in message.lower():
            self.copy_btn.configure(state="disabled")

    def _on_progress(self, fraction):
        self.progress.set(fraction)

    def _on_result(self, message, success):
        self.app_window.results.log(message)

    def _on_complete(self, summary):
        self.app_window.results.log(summary)
        self.auth_label.configure(text="Ready to connect", text_color="gray")
        self.copy_btn.configure(state="disabled")
        self._device_code = ""
        self._set_processing(False)
        self.progress.grid_remove()

    def _on_error(self, error_msg):
        self.app_window.results.log(f"Outlook error: {error_msg}")
        self.auth_label.configure(text=f"Error: {error_msg}", text_color="#f87171")
        self._set_processing(False)
        self.progress.grid_remove()

    def _set_processing(self, active):
        state = "disabled" if active else "normal"
        self.fetch_btn.configure(state=state)
