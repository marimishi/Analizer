import customtkinter as ctk
import threading
import os
from tkinter import messagebox, filedialog
from back.yandex_forms.excel_to_pandas import ExcelToPandasProcessor
from back.handler import YandexFormsHandler

class YandexConfigScreen(ctk.CTkFrame):
    def __init__(self, master, settings_manager):
        super().__init__(master, fg_color="transparent")
        self.settings_manager = settings_manager
        self.file_path = None
        self.df = None
        self.date_col = None
        
        self.selected_dates = []
        self.checkboxes = []

        self.grid_columnconfigure(0, weight=1)
        
        self.title_label = ctk.CTkLabel(
            self, 
            text="Настройки анализа Яндекс.Форм", 
            font=ctk.CTkFont(family="Arial", size=20, weight="bold")
        )
        self.title_label.grid(row=0, column=0, sticky="w", pady=(0, 15), padx=5)
        
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
            placeholder_text="Итоговый отчет по тестированию"
        )
        self.title_entry.grid(row=0, column=1, padx=20, pady=20, sticky="ew")
        
        self.date_card = None
        self.scroll_frame = None
        
        self.info_label = ctk.CTkLabel(
            self,
            text="",
            font=ctk.CTkFont(size=12, slant="italic"),
            text_color="gray"
        )
        self.info_label.grid(row=10, column=0, pady=(20, 10), padx=5, sticky="w")
        
        self.btn_generate = ctk.CTkButton(
            self, 
            text="Сгенерировать отчет презентации", 
            height=45, 
            width=400,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.run_generation
        )
        self.btn_generate.grid(row=11, column=0, pady=20)

    def load_file_structure(self, file_path):
        self.file_path = file_path
        
        default_title = YandexFormsHandler.clean_filename_from_date(file_path)
        self.title_entry.delete(0, "end")
        self.title_entry.insert(0, default_title)
        
        if self.date_card and self.date_card.winfo_exists():
            self.date_card.destroy()
        self.checkboxes.clear()

        try:
            self.df = ExcelToPandasProcessor.excel_to_pandas(file_path)
            columns = self.df.columns.tolist()
            
            self.date_col = ExcelToPandasProcessor.find_date_column(columns)
            
            if self.date_col:
                self.create_date_selection_ui()
            else:
                self.info_label.configure(
                    text="* Столбец с датой не найден. Будут обработаны все данные."
                )
                    
        except Exception as e:
            messagebox.showerror("Ошибка чтения файла", f"Не удалось прочитать структуру Excel:\n{e}")
    
    def create_date_selection_ui(self):
        """Создает UI для выбора дат"""
        self.date_card = ctk.CTkFrame(self, corner_radius=12)
        self.date_card.grid(row=2, column=0, sticky="ew", pady=10, padx=5)
        self.date_card.grid_columnconfigure(1, weight=1)
        
        date_label = ctk.CTkLabel(
            self.date_card, 
            text=f"Выберите даты ({self.date_col}):", 
            font=ctk.CTkFont(size=14, weight="bold")
        )
        date_label.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="nw")
        
        self.scroll_frame = ctk.CTkScrollableFrame(self.date_card, height=150)
        self.scroll_frame.grid(row=1, column=0, columnspan=2, padx=20, pady=(0, 20), sticky="ew")
        
        unique_dates = ExcelToPandasProcessor.get_unique_dates(self.df, self.date_col)
        
        buttons_frame = ctk.CTkFrame(self.date_card, fg_color="transparent")
        buttons_frame.grid(row=0, column=1, padx=20, pady=(20, 10), sticky="ne")
        
        btn_select_all = ctk.CTkButton(
            buttons_frame, 
            text="Выбрать все", 
            width=100, 
            height=30,
            font=ctk.CTkFont(size=12),
            command=self.select_all_dates
        )
        btn_select_all.pack(side="left", padx=5)
        
        btn_deselect_all = ctk.CTkButton(
            buttons_frame, 
            text="Снять все", 
            width=100, 
            height=30,
            font=ctk.CTkFont(size=12),
            fg_color="#555555",
            command=self.deselect_all_dates
        )
        btn_deselect_all.pack(side="left", padx=5)
        
        for date_str in unique_dates:
            cb = ctk.CTkCheckBox(self.scroll_frame, text=date_str)
            cb.select()  
            cb.pack(anchor="w", padx=10, pady=4)
            self.checkboxes.append(cb)
    
    def select_all_dates(self):
        for cb in self.checkboxes:
            cb.select()
    
    def deselect_all_dates(self):
        for cb in self.checkboxes:
            cb.deselect()

    def run_generation(self):
        if not self.file_path:
            messagebox.showwarning("Внимание", "Сначала выберите исходный файл Excel.")
            return

        chosen_dates = [cb.cget("text") for cb in self.checkboxes if cb.get() == 1] if self.checkboxes else []
        
        title_text = self.title_entry.get().strip()
        if not title_text:
            title_text = YandexFormsHandler.clean_filename_from_date(self.file_path)

        default_filename = f"{title_text}.pptx"
        
        save_path = filedialog.asksaveasfilename(
            title="Выберите место для сохранения презентации",
            initialfile=default_filename,
            defaultextension=".pptx",
            filetypes=[("PowerPoint Presentations", "*.pptx"), ("All Files", "*.*")]
        )
        
        if not save_path:
            return

        self.btn_generate.configure(state="disabled", text="Генерация слайдов...")
        
        def worker():
            try:
                filtered_df = self.df
                if chosen_dates and self.date_col:
                    filtered_df = ExcelToPandasProcessor.filter_by_dates(self.df, self.date_col, chosen_dates)
                
                filtered_df = ExcelToPandasProcessor.drop(filtered_df)
                
                handler = YandexFormsHandler(
                    file_path=self.file_path,
                    df=filtered_df,  
                    title=title_text,
                    settings_manager=self.settings_manager,
                    output_path=save_path
                )

                output_file = handler.process()
                
                self.after(0, lambda f=output_file: messagebox.showinfo(
                    "Успех", f"Презентация успешно создана и сохранена:\n{os.path.abspath(f)}"
                ))
            except Exception as ex:
                error_msg = str(ex)
                self.after(0, lambda msg=error_msg: messagebox.showerror(
                    "Ошибка генерации", f"Произошла ошибка:\n{msg}"
                ))
            finally:
                self.after(0, lambda: self.btn_generate.configure(state="normal", text="Сгенерировать отчет презентации"))

        threading.Thread(target=worker, daemon=True).start()