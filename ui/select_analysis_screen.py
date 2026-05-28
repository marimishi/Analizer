import customtkinter as ctk

class SelectAnalysisScreen(ctk.CTkFrame):
    def __init__(self, master, settings_manager):
        super().__init__(master, fg_color="transparent")
        self.settings_manager = settings_manager
        self.excel_file_path = None
        
        self.grid_columnconfigure(0, weight=1)
        
        self.title_label = ctk.CTkLabel(
            self, 
            text="Выбор источника данных", 
            font=ctk.CTkFont(family="Arial", size=20, weight="bold")
        )
        self.title_label.grid(row=0, column=0, sticky="w", pady=(0, 15), padx=5)
        
        self.yandex_card = ctk.CTkFrame(self, corner_radius=12)
        self.yandex_card.grid(row=1, column=0, sticky="ew", pady=10, padx=5)
        self.yandex_card.grid_columnconfigure(1, weight=1)
        
        self.yandex_label = ctk.CTkLabel(
            self.yandex_card, 
            text="Яндекс.Формы", 
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.yandex_label.grid(row=0, column=0, padx=20, pady=20, sticky="w")
        
        
        self.btn_yandex = ctk.CTkButton(
            self.yandex_card,
            text="Выбрать",
            width=120,
            height=35,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=self.choose_yandex
        )
        self.btn_yandex.grid(row=0, column=2, padx=20, pady=20, sticky="e")
        
        self.survey_card = ctk.CTkFrame(self, corner_radius=12)
        self.survey_card.grid(row=2, column=0, sticky="ew", pady=10, padx=5)
        self.survey_card.grid_columnconfigure(1, weight=1)
        
        self.survey_label = ctk.CTkLabel(
            self.survey_card, 
            text="Анкетирование", 
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.survey_label.grid(row=0, column=0, padx=20, pady=20, sticky="w")
        
        
        self.btn_survey = ctk.CTkButton(
            self.survey_card,
            text="Выбрать",
            width=120,
            height=35,
            fg_color="#555555",
            hover_color="#444444",
            font=ctk.CTkFont(size=13, weight="bold"),
            command=self.choose_survey
        )
        self.btn_survey.grid(row=0, column=2, padx=20, pady=20, sticky="e")
        

    def set_excel_file(self, path):
        self.excel_file_path = path

    def choose_yandex(self):
        main_screen = self.master
        while hasattr(main_screen, "master") and main_screen.master is not None:
            main_screen = main_screen.master
            
        if hasattr(main_screen, "frames") and "YandexConfigScreen" in main_screen.frames:
            config_screen = main_screen.frames["YandexConfigScreen"]
            config_screen.load_file_structure(self.excel_file_path)
            main_screen.show_screen("YandexConfigScreen")

    def choose_survey(self):
        print("Выбран Survey. На данный момент функция находится в разработке.")