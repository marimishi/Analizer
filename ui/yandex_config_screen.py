#TODO: 
#1. Если есть столбец дата создания ответа, то можно выбирать по каким датам
#Связано с back/yandex_forms/excel_to_pandas.py

import customtkinter as ctk
import threading
from tkinter import messagebox
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
        
        # Заголовок
        self.title_label = ctk.CTkLabel(
            self, 
            text="Настройки анализа Яндекс.Форм", 
            font=ctk.CTkFont(family="Arial", size=20, weight="bold")
        )
        self.title_label.grid(row=0, column=0, sticky="w", pady=(0, 15), padx=5)
        
        # Карточка названия презентации
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
        
        # Карточка выбора дат (появится, если найден столбец с датой)
        self.date_card = None
        self.scroll_frame = None
        
        # Информационная метка
        self.info_label = ctk.CTkLabel(
            self,
            text="",
            font=ctk.CTkFont(size=12, slant="italic"),
            text_color="gray"
        )
        self.info_label.grid(row=10, column=0, pady=(20, 10), padx=5, sticky="w")
        
        # Кнопка генерации
        self.btn_generate = ctk.CTkButton(
            self, 
            text="Сгенерировать отчет презентации", 
            height=45, 
            width=400,
            fg_color="#2E7D32", 
            hover_color="#1B5E20",
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.run_generation
        )
        self.btn_generate.grid(row=11, column=0, pady=20)

    def load_file_structure(self, file_path):
        """Загружает структуру файла и создает UI для выбора дат"""
        self.file_path = file_path
        
        # Очищаем предыдущие элементы
        if self.date_card and self.date_card.winfo_exists():
            self.date_card.destroy()
        self.checkboxes.clear()

        try:
            # Загружаем данные через ExcelToPandasProcessor
            self.df = ExcelToPandasProcessor.excel_to_pandas(file_path)
            columns = self.df.columns.tolist()
            
            # Поиск столбца с датой через ExcelToPandasProcessor
            self.date_col = ExcelToPandasProcessor.find_date_column(columns)
            
            # Карточка выбора дат
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
        
        # Создаем скроллируемый фрейм для чекбоксов
        self.scroll_frame = ctk.CTkScrollableFrame(self.date_card, height=150)
        self.scroll_frame.grid(row=1, column=0, columnspan=2, padx=20, pady=(0, 20), sticky="ew")
        
        # Получаем уникальные даты через ExcelToPandasProcessor
        unique_dates = ExcelToPandasProcessor.get_unique_dates(self.df, self.date_col)
        
        # Кнопки "Выбрать все" и "Снять все"
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
        
        # Создаем чекбоксы для каждой даты
        for date_str in unique_dates:
            cb = ctk.CTkCheckBox(self.scroll_frame, text=date_str)
            cb.select()  # По дефолту все выбраны
            cb.pack(anchor="w", padx=10, pady=4)
            self.checkboxes.append(cb)
    
    def select_all_dates(self):
        """Выбирает все даты"""
        for cb in self.checkboxes:
            cb.select()
    
    def deselect_all_dates(self):
        """Снимает выбор со всех дат"""
        for cb in self.checkboxes:
            cb.deselect()

    def run_generation(self):
        """Запускает генерацию отчета"""
        chosen_dates = [cb.cget("text") for cb in self.checkboxes if cb.get() == 1] if self.checkboxes else []
        
        title_text = self.title_entry.get().strip()
        if not title_text:
            title_text = "Отчет по результатам тестирования"

        self.btn_generate.configure(state="disabled", text="Генерация слайдов...")
        
        def worker():
            try:
                # Фильтруем данные по датам через ExcelToPandasProcessor
                filtered_df = self.df
                if chosen_dates and self.date_col:
                    filtered_df = ExcelToPandasProcessor.filter_by_dates(self.df, self.date_col, chosen_dates)
                
                # Применяем фильтр столбцов через ExcelToPandasProcessor
                filtered_df = ExcelToPandasProcessor.drop(filtered_df)
                
                handler = YandexFormsHandler(
                    file_path=self.file_path,
                    df=filtered_df,  # Передаем уже обработанный DataFrame
                    title=title_text,
                    settings_manager=self.settings_manager,
                )

                output_file = handler.process()
                
                self.after(0, lambda f=output_file: messagebox.showinfo(
                    "Успех", f"Презентация успешно создана и сохранена в корень проекта как:\n{f}"
                ))
            except Exception as ex:
                error_msg = str(ex)
                self.after(0, lambda msg=error_msg: messagebox.showerror(
                    "Ошибка генерации", f"Произошла ошибка:\n{msg}"
                ))
            finally:
                self.after(0, lambda: self.btn_generate.configure(state="normal", text="Сгенерировать отчет презентации"))

        threading.Thread(target=worker, daemon=True).start()