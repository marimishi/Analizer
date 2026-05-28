import customtkinter as ctk
import ui.main_screen as main
from ui.settings import Settings

if __name__ == "__main__":
    settings_manager = Settings()

    current_settings = settings_manager.load()
    saved_mode = current_settings.get("mode", "System")
    saved_color = current_settings.get("color_theme", "blue")

    ctk.set_appearance_mode(saved_mode)
    ctk.set_default_color_theme(saved_color)

    app = main.MainScreen(settings_manager=settings_manager, backend_handler=None)
    
    app.mainloop()