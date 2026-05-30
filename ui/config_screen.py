import customtkinter as ctk
import threading
import os
from tkinter import messagebox, filedialog


class BaseConfigScreen(ctk.CTkFrame):
    """Базовый класс для экранов настройки анализа"""
    
    def __init__(self, master, settings_manager):
        super().__init__(master, fg_color="transparent")
        self.settings_manager = settings_manager
        self.file_path = None
        self.df = None
        
        self.grid_columnconfigure(0, weight=1)
        
        self.title_label = None
        
        self.title_card = None
        self.title_entry = None
        
        self.btn_generate = None
        
        self._create_title_ui()
        
    def _create_title_ui(self):
        """Создает общий UI для названия презентации"""
        self.title_card = ctk.CTkFrame(self, corner_radius=12)
        self.title_card.grid(row=1, column=0, sticky="ew", pady=10, padx=5)
        self.title_card.grid_columnconfigure(1, weight=1)
        
        self.title_label_card = ctk.CTkLabel(
            self.title_card, 
            text="Название презентации:", 
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.title_label_card.grid(row=0, column=0, padx=20, pady=20, sticky="w")
        
        self.title_entry = ctk.CTkEntry(
            self.title_card, 
            width=300,
            placeholder_text="Итоговый отчет"
        )
        self.title_entry.grid(row=0, column=1, padx=20, pady=20, sticky="ew")
        
    def create_generate_button(self, text="Сгенерировать презентацию", row=11):
        """Создает кнопку генерации"""
        self.btn_generate = ctk.CTkButton(
            self, 
            text=text, 
            height=45, 
            width=400,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.run_generation
        )
        self.btn_generate.grid(row=row, column=0, pady=20)
        
    def set_title_from_filename(self, file_path):
        """Устанавливает название презентации из имени файла"""
        filename_clean = os.path.splitext(os.path.basename(file_path))[0]
        self.title_entry.delete(0, "end")
        self.title_entry.insert(0, filename_clean)
        
    def get_title(self) -> str:
        """Возвращает название презентации"""
        title = self.title_entry.get().strip()
        return title if title else "Отчет"
    
    def save_file_dialog(self, default_filename: str) -> str:
        """Открывает диалог сохранения файла"""
        return filedialog.asksaveasfilename(
            title="Выберите место для сохранения презентации",
            initialfile=default_filename,
            defaultextension=".pptx",
            filetypes=[("PowerPoint Presentations", "*.pptx"), ("All Files", "*.*")]
        )
    
    def show_success_message(self, file_path: str):
        """Показывает сообщение об успехе"""
        messagebox.showinfo(
            "Успех", 
            f"Презентация успешно сохранена:\n{os.path.abspath(file_path)}"
        )
    
    def show_error_message(self, error_msg: str):
        """Показывает сообщение об ошибке"""
        messagebox.showerror("Ошибка генерации", f"Произошла ошибка:\n{error_msg}")
    
    def show_warning_message(self, warning_msg: str):
        """Показывает предупреждение"""
        messagebox.showwarning("Внимание", warning_msg)
    
    def set_generating_state(self, is_generating: bool):
        """Блокирует/разблокирует кнопку генерации"""
        if self.btn_generate:
            if is_generating:
                self.btn_generate.configure(state="disabled")
            else:
                self.btn_generate.configure(state="normal")
    
    def run_in_thread(self, target_func, on_success=None, on_error=None):
        """Запускает функцию в отдельном потоке"""
        def wrapper():
            try:
                result = target_func()
                if on_success and result:
                    self.after(0, lambda: on_success(result))
            except Exception as ex:
                import traceback
                traceback.print_exc()
                if on_error:
                    self.after(0, lambda err=ex: on_error(str(err)))
                else:
                    self.after(0, lambda err=ex: self.show_error_message(str(err)))
            finally:
                self.after(0, lambda: self.set_generating_state(False))
        
        self.set_generating_state(True)
        threading.Thread(target=wrapper, daemon=True).start()
    
    def run_generation(self):
        """Метод для переопределения в наследниках"""
        raise NotImplementedError("Метод run_generation должен быть переопределен")