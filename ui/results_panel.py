"""Scrollable results log widget."""

from datetime import datetime
import customtkinter as ctk


class ResultsPanel(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        header = ctk.CTkLabel(self, text="Results Log", font=ctk.CTkFont(size=14, weight="bold"))
        header.grid(row=0, column=0, sticky="w", padx=10, pady=(8, 2))

        self.textbox = ctk.CTkTextbox(self, state="disabled", font=ctk.CTkFont(family="Courier", size=13))
        self.textbox.grid(row=1, column=0, sticky="nsew", padx=8, pady=(2, 8))

    def log(self, message: str, tag: str = "info"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        line = f"[{timestamp}] {message}\n"
        self.textbox.configure(state="normal")
        self.textbox.insert("end", line)
        self.textbox.see("end")
        self.textbox.configure(state="disabled")

    def clear(self):
        self.textbox.configure(state="normal")
        self.textbox.delete("1.0", "end")
        self.textbox.configure(state="disabled")
