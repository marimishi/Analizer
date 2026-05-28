import json
from pathlib import Path
import os
import json


class Settings:
    SETTINGS_FILE = Path(__file__).resolve().parent.parent / "settings.json"

    def load(self):
        if not self.SETTINGS_FILE.exists():
            return {}
            
        try:
            with open(self.SETTINGS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except json.JSONDecodeError:
            return {}

    def save(self, choice, value):
        data = self.load()
        data[f"{choice}"] = value

        with open(self.SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
        print(f"Сохранено в settings.json: {choice} = {value}")
