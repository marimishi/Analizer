import json
import os

class Settings:
    def __init__(self):
        appdata_path = os.environ.get("APPDATA", os.getcwd())
        
        self.app_folder = os.path.join(appdata_path, "Analizer")
        
        self.file_path = os.path.join(self.app_folder, "settings.json")
        
        self.default_settings = {
            "color_theme": "dark-blue",
            "mode": "System"
        }
        
        self.memory_settings = self.load()

    def load(self):
        if os.path.exists(self.file_path):
            try:
                with open(self.file_path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                return self.default_settings.copy()
        else:
            return self.default_settings.copy()

    def save(self, choice, value):
        self.memory_settings[str(choice)] = value
        
        try:
            if not os.path.exists(self.app_folder):
                os.makedirs(self.app_folder)
            
            with open(self.file_path, "w", encoding="utf-8") as f:
                json.dump(self.memory_settings, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"Ошибка при сохранении настроек: {e}")