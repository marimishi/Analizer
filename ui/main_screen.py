import customtkinter as ctk
from ui.upload_excel_screen import UploadExcelScreen
from ui.select_analysis_screen import SelectAnalysisScreen
from ui.settings_screen import SettingsScreen
from ui.yandex_config_screen import YandexConfigScreen


class MainScreen(ctk.CTk):
    def __init__(self, settings_manager, backend_handler):
        super().__init__()
        self.settings_manager = settings_manager
        self.backend_handler = backend_handler

        self.title("Анализатор отчетов")
        self.geometry("1000x650")
        self.minsize(900, 550)

        # Конфигурация главной сетки окна (столбец 0 — сайдбар, столбец 1 — контент)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Создаем боковую панель меню (Sidebar) в column=0
        self._create_sidebar()

        # Контейнер для экранов справа в column=1
        self.container = ctk.CTkFrame(self, fg_color="transparent")
        self.container.grid(row=0, column=1, sticky="nsew", padx=25, pady=25)
        self.container.grid_columnconfigure(0, weight=1)
        self.container.grid_rowconfigure(0, weight=1)

        self.frames = {}
        
        for ScreenClass in [UploadExcelScreen, SelectAnalysisScreen, SettingsScreen, YandexConfigScreen]:
            screen_name = ScreenClass.__name__
            frame = ScreenClass(master=self.container, settings_manager=self.settings_manager)
            self.frames[screen_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_screen("UploadExcelScreen")

    def _create_sidebar(self):
        """Создает левую панель навигации"""
        self.sidebar_frame = ctk.CTkFrame(self, width=220, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(3, weight=1) # Растягиваем пустое место внизу

        self.logo_label = ctk.CTkLabel(
            self.sidebar_frame, 
            text="Анализатор", 
            font=ctk.CTkFont(family="Arial", size=22, weight="bold")
        )
        self.logo_label.grid(row=0, column=0, padx=20, pady=25, sticky="w")

        self.btn_excel = ctk.CTkButton(
            self.sidebar_frame, 
            text="Анализ Excel", 
            font=ctk.CTkFont(size=13, weight="bold"),
            height=40, corner_radius=8,
            command=self.click_excel_nav
        )
        self.btn_excel.grid(row=1, column=0, padx=15, pady=8, sticky="ew")

        self.btn_settings = ctk.CTkButton(
            self.sidebar_frame, 
            text="Настройки", 
            font=ctk.CTkFont(size=13, weight="bold"),
            height=40, corner_radius=8,
            command=self.click_settings_nav
        )
        self.btn_settings.grid(row=2, column=0, padx=15, pady=8, sticky="ew")

    def click_excel_nav(self):
        """Безопасный обработчик клика по меню 'Анализ Excel'"""
        self.show_screen("UploadExcelScreen")

    def click_settings_nav(self):
        """Безопасный обработчик клика по меню 'Настройки'"""
        self.show_screen("SettingsScreen")

    def show_screen(self, screen_name: str):
        """Показывает экран по его имени, скрывая остальные"""
        if screen_name in self.frames:
            frame = self.frames[screen_name]
            frame.tkraise()  # Поднимаем нужный слой наверх внутри контейнера
        else:
            print(f"Ошибка: Экран '{screen_name}' не найден в зарегистрированных фреймах.")