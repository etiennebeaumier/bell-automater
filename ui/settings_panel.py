"""Settings sidebar panel."""

import os
from tkinter import filedialog
import customtkinter as ctk


class SettingsPanel(ctk.CTkFrame):
    def __init__(self, master, config, on_settings_changed=None, **kwargs):
        super().__init__(master, width=280, **kwargs)
        self.config = config
        self.on_settings_changed = on_settings_changed

        self.grid_columnconfigure(0, weight=1)

        # -- Title --
        title = ctk.CTkLabel(self, text="Control Panel", font=ctk.CTkFont(size=16, weight="bold"))
        title.grid(row=0, column=0, padx=16, pady=(16, 4), sticky="w")

        subtitle = ctk.CTkLabel(self, text="Workbook path and execution mode", font=ctk.CTkFont(size=12), text_color="gray")
        subtitle.grid(row=1, column=0, padx=16, pady=(0, 12), sticky="w")

        # -- Workbook path --
        wb_label = ctk.CTkLabel(self, text="Master Workbook", font=ctk.CTkFont(size=13, weight="bold"))
        wb_label.grid(row=2, column=0, padx=16, pady=(8, 2), sticky="w")

        path_frame = ctk.CTkFrame(self, fg_color="transparent")
        path_frame.grid(row=3, column=0, padx=16, pady=(0, 4), sticky="ew")
        path_frame.grid_columnconfigure(0, weight=1)

        self.workbook_var = ctk.StringVar(value=config.get("master_file", ""))
        self.workbook_entry = ctk.CTkEntry(path_frame, textvariable=self.workbook_var, placeholder_text="Select workbook...")
        self.workbook_entry.grid(row=0, column=0, sticky="ew", padx=(0, 6))

        browse_btn = ctk.CTkButton(path_frame, text="Browse", width=70, command=self._browse_workbook)
        browse_btn.grid(row=0, column=1)

        # -- Status indicator --
        self.status_label = ctk.CTkLabel(self, text="", font=ctk.CTkFont(size=12))
        self.status_label.grid(row=4, column=0, padx=16, pady=(4, 8), sticky="w")
        self._update_status()

        # -- Dry run checkbox --
        self.dry_run_var = ctk.BooleanVar(value=config.get("dry_run", False))
        self.dry_run_check = ctk.CTkCheckBox(
            self, text="Dry run (preview only)", variable=self.dry_run_var,
            command=self._on_change,
        )
        self.dry_run_check.grid(row=5, column=0, padx=16, pady=(8, 8), sticky="w")

        # -- Appearance mode --
        theme_label = ctk.CTkLabel(self, text="Appearance", font=ctk.CTkFont(size=13, weight="bold"))
        theme_label.grid(row=6, column=0, padx=16, pady=(16, 2), sticky="w")

        self.theme_menu = ctk.CTkOptionMenu(
            self, values=["Dark", "Light", "System"],
            command=self._change_theme,
        )
        self.theme_menu.set(config.get("appearance_mode", "dark").capitalize())
        self.theme_menu.grid(row=7, column=0, padx=16, pady=(0, 8), sticky="ew")

        # -- Save settings button --
        save_btn = ctk.CTkButton(self, text="Save Settings", command=self._save_settings)
        save_btn.grid(row=8, column=0, padx=16, pady=(16, 16), sticky="ew")

        # Bind entry change
        self.workbook_var.trace_add("write", lambda *_: self._on_workbook_change())

    def _browse_workbook(self):
        path = filedialog.askopenfilename(
            title="Select Master Workbook",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.workbook_var.set(path)

    def _on_workbook_change(self):
        self._update_status()
        self._on_change()

    def _update_status(self):
        path = self.workbook_var.get()
        if not path:
            self.status_label.configure(text="  No workbook selected", text_color="#f87171")
            return
        if not os.path.isfile(path):
            self.status_label.configure(text="  Workbook not found", text_color="#f87171")
            return
        try:
            from openpyxl import load_workbook
            wb = load_workbook(path, read_only=True)
            sheets = set(wb.sheetnames)
            wb.close()
            required = {"Pricing", "Summary Charts"}
            missing = required - sheets
            if missing:
                self.status_label.configure(text=f"  Missing: {', '.join(missing)}", text_color="#f87171")
            else:
                self.status_label.configure(text="  Workbook ready", text_color="#4ade80")
        except Exception as e:
            self.status_label.configure(text=f"  Error: {e}", text_color="#f87171")

    def _change_theme(self, choice):
        ctk.set_appearance_mode(choice.lower())
        self._on_change()

    def _on_change(self):
        if self.on_settings_changed:
            self.on_settings_changed()

    def _save_settings(self):
        self.config["master_file"] = self.workbook_var.get()
        self.config["dry_run"] = self.dry_run_var.get()
        self.config["appearance_mode"] = self.theme_menu.get().lower()
        self.config.save()
        self._update_status()

    @property
    def master_file(self) -> str:
        return self.workbook_var.get()

    @property
    def dry_run(self) -> bool:
        return self.dry_run_var.get()

    def is_workbook_ready(self) -> bool:
        path = self.workbook_var.get()
        if not path or not os.path.isfile(path):
            return False
        try:
            from openpyxl import load_workbook
            wb = load_workbook(path, read_only=True)
            sheets = set(wb.sheetnames)
            wb.close()
            return {"Pricing", "Summary Charts"}.issubset(sheets)
        except Exception:
            return False
