import customtkinter as ctk

class SettingsScreen(ctk.CTkFrame):
    def __init__(self, master, settings_manager):
        super().__init__(master, fg_color="transparent")
        self.settings_manager = settings_manager
        
        self.grid_columnconfigure(0, weight=1)
        self.current_settings = self.settings_manager.load()

        self.ru_to_en_mode = {
            "Системная": "System",
            "Темная": "Dark",
            "Светлая": "Light"
        }
        self.en_to_ru_mode = {v: k for k, v in self.ru_to_en_mode.items()}

        self.ru_to_en_color = {
            "Синий": "blue",
            "Зеленый": "green",
            "Темно-синий": "dark-blue"
        }
        self.en_to_ru_color = {v: k for k, v in self.ru_to_en_color.items()}

        self.title_label = ctk.CTkLabel(
            self, 
            text="Настройки приложения", 
            font=ctk.CTkFont(family="Arial", size=20, weight="bold")
        )
        self.title_label.grid(row=0, column=0, sticky="w", pady=(0, 15))

        self.mode_card = ctk.CTkFrame(self, corner_radius=12)
        self.mode_card.grid(row=1, column=0, sticky="ew", pady=10, padx=5)
        self.mode_card.grid_columnconfigure(1, weight=1)

        self.mode_label = ctk.CTkLabel(self.mode_card, text="Режим оформления:", font=ctk.CTkFont(size=14))
        self.mode_label.grid(row=0, column=0, padx=20, pady=20, sticky="w")

        self.mode_switch = ctk.CTkOptionMenu(
            self.mode_card, 
            values=list(self.ru_to_en_mode.keys()),
            command=self._change_mode
        )
        self.mode_switch.grid(row=0, column=1, padx=20, pady=20, sticky="e")
        
        # Инициализация текущего режима
        current_mode_en = self.current_settings.get("mode", "System")
        if isinstance(current_mode_en, dict): current_mode_en = "System"
        self.mode_switch.set(self.en_to_ru_mode.get(current_mode_en, "Системная"))

        self.color_card = ctk.CTkFrame(self, corner_radius=12)
        self.color_card.grid(row=2, column=0, sticky="ew", pady=10, padx=5)
        self.color_card.grid_columnconfigure(1, weight=1)

        self.color_label = ctk.CTkLabel(self.color_card, text="Цветовая схема:", font=ctk.CTkFont(size=14))
        self.color_label.grid(row=0, column=0, padx=20, pady=20, sticky="w")

        self.color_switch = ctk.CTkOptionMenu(
            self.color_card,
            values=list(self.ru_to_en_color.keys()),
            command=self._change_color
        )
        self.color_switch.grid(row=0, column=1, padx=20, pady=20, sticky="e")

        current_color_en = self.current_settings.get("color_theme", "blue")
        if isinstance(current_color_en, dict): current_color_en = "blue"
        self.color_switch.set(self.en_to_ru_color.get(current_color_en, "Синий"))
        
        self.info_label = ctk.CTkLabel(
            self, 
            text="", 
            font=ctk.CTkFont(size=12, slant="italic"),
            text_color="gray"
        )
        self.info_label.grid(row=3, column=0, pady=10, padx=5, sticky="w")

    def _change_mode(self, new_mode_ru):
        new_mode_en = self.ru_to_en_mode.get(new_mode_ru, "System")
        ctk.set_appearance_mode(new_mode_en)
        self.settings_manager.save("mode", new_mode_en)

    def _change_color(self, new_color_ru):
        new_color_en = self.ru_to_en_color.get(new_color_ru, "blue")
        ctk.set_default_color_theme(new_color_en)
        self.settings_manager.save("color_theme", new_color_en)
        
        self.info_label.configure(
            text="* Цветовая схема применится полностью после перезапуска приложения."
        )