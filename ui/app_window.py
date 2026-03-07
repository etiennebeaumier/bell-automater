"""Main application window."""

import customtkinter as ctk
from ui.settings_panel import SettingsPanel
from ui.tab_pdf import PdfTab
from ui.tab_outlook import OutlookTab
from ui.results_panel import ResultsPanel


class AppWindow(ctk.CTk):
    def __init__(self, config):
        super().__init__()
        self.config = config

        self.title("BCECN Pricing Tool")
        self.geometry("1050x680")
        self.minsize(900, 550)

        ctk.set_appearance_mode(config.get("appearance_mode", "dark"))
        ctk.set_default_color_theme("blue")

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # -- Settings panel (left sidebar) --
        self.settings = SettingsPanel(self, config, on_settings_changed=self._on_settings_changed)
        self.settings.grid(row=0, column=0, sticky="nsew", padx=(8, 0), pady=8)

        # -- Right content area --
        right_frame = ctk.CTkFrame(self, fg_color="transparent")
        right_frame.grid(row=0, column=1, sticky="nsew", padx=8, pady=8)
        right_frame.grid_columnconfigure(0, weight=1)
        right_frame.grid_rowconfigure(0, weight=0)
        right_frame.grid_rowconfigure(1, weight=1)
        right_frame.grid_rowconfigure(2, weight=1)

        # -- Header card --
        self.header = ctk.CTkFrame(right_frame, corner_radius=12)
        self.header.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self.header.grid_columnconfigure(0, weight=1)

        title_label = ctk.CTkLabel(
            self.header, text="BCECN Pricing Tool",
            font=ctk.CTkFont(size=22, weight="bold"),
        )
        title_label.grid(row=0, column=0, padx=16, pady=(12, 0), sticky="w")

        subtitle_label = ctk.CTkLabel(
            self.header, text="Bell Canada new issue spread and yield automation",
            font=ctk.CTkFont(size=13), text_color="gray",
        )
        subtitle_label.grid(row=1, column=0, padx=16, pady=(0, 4), sticky="w")

        self.mode_label = ctk.CTkLabel(
            self.header, text="", font=ctk.CTkFont(size=12), text_color="gray",
        )
        self.mode_label.grid(row=2, column=0, padx=16, pady=(0, 10), sticky="w")
        self._update_mode_label()

        # -- Tabview --
        self.tabview = ctk.CTkTabview(right_frame)
        self.tabview.grid(row=1, column=0, sticky="nsew", pady=(0, 4))

        self.tabview.add("Upload PDFs")
        self.tabview.add("Fetch from Outlook")

        self.pdf_tab = PdfTab(self.tabview.tab("Upload PDFs"), app_window=self)
        self.pdf_tab.pack(fill="both", expand=True)

        self.outlook_tab = OutlookTab(self.tabview.tab("Fetch from Outlook"), app_window=self, config=config)
        self.outlook_tab.pack(fill="both", expand=True)

        # -- Results panel --
        self.results = ResultsPanel(right_frame)
        self.results.grid(row=2, column=0, sticky="nsew", pady=(4, 0))

    def _on_settings_changed(self):
        self._update_mode_label()

    def _update_mode_label(self):
        mode = "Dry-run preview" if self.settings.dry_run else "Live workbook"
        wb_status = "ready" if self.settings.is_workbook_ready() else "not set"
        self.mode_label.configure(text=f"Mode: {mode}  |  Workbook: {wb_status}")
