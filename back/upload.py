import os
import sys
import shutil
from tkinter import filedialog
from pathlib import Path
from PIL import Image
import customtkinter as ctk
from pptx import Presentation

class Upload:
    def __init__(self, master, settings_manager):
        self.master = master
        self.settings_manager = settings_manager


    def excel(self):
        file_path = filedialog.askopenfilename(
            title="Выберите файл Excel",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if not file_path:
            print("Файл не выбран")
            return

        main_screen = self.master.winfo_toplevel()
        if hasattr(main_screen, "frames") and "SelectAnalysisScreen" in main_screen.frames:
            analysis_screen = main_screen.frames["SelectAnalysisScreen"]
            analysis_screen.set_excel_file(file_path)
            main_screen.show_screen("SelectAnalysisScreen")