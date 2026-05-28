import customtkinter as ctk
from back.upload import Upload

class UploadExcelScreen(ctk.CTkFrame):
    def __init__(self, master, settings_manager):
        super().__init__(master, fg_color="transparent")
        self.settings_manager = settings_manager
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)
        
        self.uploader = Upload(master=self, settings_manager=self.settings_manager)

        self.title_label = ctk.CTkLabel(
            self, 
            text="Анализ Excel", 
            font=ctk.CTkFont(family="Arial", size=22, weight="bold")
        )
        self.title_label.grid(row=0, column=0, pady=(10, 20), sticky="w", padx=10)

        self.info_card = ctk.CTkFrame(self, corner_radius=12)
        self.info_card.grid(row=1, column=0, sticky="ew", pady=10, padx=10)
        self.info_card.grid_columnconfigure(0, weight=1)


        self.info_label = ctk.CTkLabel(
            self.info_card, 
            text="Загрузите Excel-файл выгрузки ответов.\nПриложение автоматически обработает структуру данных\nи подготовит аналитическую презентацию.",
            font=ctk.CTkFont(size=14),
            justify="center"
        )
        self.info_label.grid(row=1, column=0, padx=30, pady=(5, 20), sticky="ew")

        self.btn_action = ctk.CTkButton(
            self, 
            text="Прикрепить файл Excel", 
            height=48, 
            width=300,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self.start_analysis
        )
        self.btn_action.grid(row=2, column=0, pady=25, sticky="n")
        
        self.status_label = ctk.CTkLabel(
            self, 
            text="", 
            font=ctk.CTkFont(size=13, weight="bold")
        )
        self.status_label.grid(row=3, column=0, pady=10, sticky="n")

    def start_analysis(self):
        self.status_label.configure(text="") 
        
        try:
            self.uploader.excel()
        except Exception as e:
            self.show_status_message(f"Ошибка: {str(e)}")
    
    def show_status_message(self, message):
        color = "#E42E25"
        
        self.status_label.configure(
            text=f"{message}",
            text_color=color
        )
        self.after(4000, self.clear_status_message)

    def clear_status_message(self):
        if self.status_label.winfo_exists():
            self.status_label.configure(text="")