import customtkinter as ctk
import threading
import os
import pandas as pd
from tkinter import messagebox, filedialog
from pptx import Presentation
from pptx.util import Inches
from back.handler import Survey

from back.excel_to_pptx.survey.survey_diagrams import survey_base_diagram
from back.excel_to_pptx.survey.conclusion import create_survey_conclusions

class SurveyConfigScreen(ctk.CTkFrame):
    def __init__(self, master, settings_manager):
        super().__init__(master, fg_color="transparent")
        self.settings_manager = settings_manager
        
        self.file_path = None
        self.df = None
        self.prs = None
        self.blank_layout = None
        
        self.charts_config = []
        self.current_index = 0
        self.user_choices = {}

        self.grid_columnconfigure(0, weight=1)
        
        self.title_label = ctk.CTkLabel(
            self, 
            text="Настройки анализа Опросов (Survey)", 
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
            placeholder_text="Итоговый отчет по анкетированию"
        )
        self.title_entry.grid(row=0, column=1, padx=20, pady=20, sticky="ew")

        self.card_frame = ctk.CTkFrame(self, corner_radius=12)
        self.card_frame.grid(row=2, column=0, sticky="ew", pady=10, padx=5)
        self.card_frame.grid_columnconfigure(0, weight=1)

        self.question_title = ctk.CTkLabel(
            self.card_frame, text="Загрузка вопросов...", 
            font=ctk.CTkFont(size=15, weight="bold"), wraplength=600, justify="left"
        )
        self.question_title.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")
        
        self.preview_data_label = ctk.CTkLabel(
            self.card_frame, text="", 
            font=ctk.CTkFont(size=13), justify="left", text_color="gray"
        )
        self.preview_data_label.grid(row=1, column=0, padx=20, pady=10, sticky="w")
        
        self.radio_var = ctk.StringVar(value="")
        self.radio_frame = ctk.CTkFrame(self.card_frame, fg_color="transparent")
        self.radio_frame.grid(row=2, column=0, padx=20, pady=15, sticky="w")
        
        self.r1 = ctk.CTkRadioButton(self.radio_frame, text="Позитивный аспект", variable=self.radio_var, value="Позитивный аспект", command=self.on_choice_changed)
        self.r1.grid(row=0, column=0, padx=(0, 20))
        
        self.r2 = ctk.CTkRadioButton(self.radio_frame, text="Проблемная область", variable=self.radio_var, value="Проблемная область", command=self.on_choice_changed)
        self.r2.grid(row=0, column=1, padx=20)
        
        self.r3 = ctk.CTkRadioButton(self.radio_frame, text="Не включать в выводы", variable=self.radio_var, value="Не включать в выводы", command=self.on_choice_changed)
        self.r3.grid(row=0, column=2, padx=20)

        self.nav_frame = ctk.CTkFrame(self.card_frame, fg_color="transparent")
        self.nav_frame.grid(row=3, column=0, padx=20, pady=(10, 20), sticky="ew")
        
        self.btn_prev = ctk.CTkButton(self.nav_frame, text="Назад", width=100, command=self.prev_item)
        self.btn_prev.pack(side="left", padx=(0, 20))
        
        self.counter_label = ctk.CTkLabel(self.nav_frame, text="0 / 0", font=ctk.CTkFont(size=13, weight="bold"))
        self.counter_label.pack(side="left", padx=10)
        
        self.btn_next = ctk.CTkButton(self.nav_frame, text="Вперед", width=100, command=self.next_item)
        self.btn_next.pack(side="left", padx=20)

        self.btn_generate = ctk.CTkButton(
            self, 
            text="Сформировать презентацию с выводами", 
            height=45, 
            width=400,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.run_generation
        )
        self.btn_generate.grid(row=11, column=0, pady=20)

    def load_file_structure(self, file_path):
        self.file_path = file_path
        
        filename_clean = os.path.splitext(os.path.basename(file_path))[0]
        self.title_entry.delete(0, "end")
        self.title_entry.insert(0, filename_clean)
        
        try:
            if file_path.endswith('.csv'):
                self.df = pd.read_csv(file_path)
            else:
                self.df = pd.read_excel(file_path)
            
            self.prs = Presentation()
            self.prs.slide_width = Inches(13.333)
            self.prs.slide_height = Inches(7.5)
            self.blank_layout = self.prs.slide_layouts[6]
            
            self.survey_manager = Survey(self.df, self.prs, self.blank_layout)
            
            self.charts_config = self.survey_manager.generate_base_diagrams(
                file_path=self.file_path, 
                title_text=filename_clean
            )
            
            self.current_index = 0
            self.user_choices.clear()
            for cfg in self.charts_config:
                self.user_choices[cfg["column_name"]] = cfg["predicted_aspect"]
                
            self.update_current_item_view()
            
        except Exception as e:
            messagebox.showerror("Ошибка чтения файла", f"Не удалось обработать опрос:\n{e}")
    def update_current_item_view(self):
        if not self.charts_config:
            self.question_title.configure(text="Подходящие вопросы для диаграмм не найдены.")
            self.counter_label.configure(text="0 / 0")
            return
            
        current_cfg = self.charts_config[self.current_index]
        col_name = current_cfg["column_name"]
        
        self.question_title.configure(text=f"Вопрос ({self.current_index + 1}/{len(self.charts_config)}):\n{col_name}")
        self.counter_label.configure(text=f"{self.current_index + 1} / {len(self.charts_config)}")
        
        valid_series = self.df[col_name].dropna().astype(str).str.strip()
        total = len(valid_series)
        val_counts = valid_series.value_counts()
        
        preview_text = "Текущее распределение ответов в файле:\n"
        
        for ans, count in val_counts.items():
            percentage = (count / total) * 100
            bar_chars = int(percentage / 10)
            bar = "█" * bar_chars + "░" * (10 - bar_chars)
            
            short_ans = ans if len(ans) <= 40 else ans[:37] + "..."
            preview_text += f"  {bar}  {percentage:5.1f}%  —  {short_ans} ({count} шт.)\n"
            
        self.preview_data_label.configure(text=preview_text)
        
        current_choice = self.user_choices.get(col_name, "Не включать в выводы")
        self.radio_var.set(current_choice)
        
        if current_choice == "Позитивный аспект":
            self.card_frame.configure(border_color="#3d9e76", border_width=2)
        elif current_choice == "Проблемная область":
            self.card_frame.configure(border_color="#c7522a", border_width=2)
        else:
            self.card_frame.configure(border_color="", border_width=0)

    def on_choice_changed(self):
        if self.charts_config:
            col_name = self.charts_config[self.current_index]["column_name"]
            self.user_choices[col_name] = self.radio_var.get()

    def next_item(self):
        if self.current_index < len(self.charts_config) - 1:
            self.current_index += 1
            self.update_current_item_view()

    def prev_item(self):
        if self.current_index > 0:
            self.current_index -= 1
            self.update_current_item_view()

    def run_generation(self):
        if not self.file_path or not self.charts_config:
            messagebox.showwarning("Внимание", "Нет данных для генерации отчета.")
            return

        title_text = self.title_entry.get().strip()
        if not title_text:
            title_text = "Отчет по опросу"

        save_path = filedialog.asksaveasfilename(
            title="Выберите место для сохранения отчета опроса",
            initialfile=f"{title_text}.pptx",
            defaultextension=".pptx",
            filetypes=[("PowerPoint Presentations", "*.pptx"), ("All Files", "*.*")]
        )
        
        if not save_path:
            return

        self.btn_generate.configure(state="disabled", text="Добавление выводов и сохранение...")

        def worker():
            try:
                structured_conclusions = {
                    "Позитивный аспект": [],
                    "Проблемная область": [],
                    "Не включать в выводы": []
                }
                for q_name, status in self.user_choices.items():
                    structured_conclusions[status].append(q_name)
                
                create_survey_conclusions(self.prs, structured_conclusions, self.blank_layout)
                
                self.prs.save(save_path)
                
                self.after(0, lambda: messagebox.showinfo(
                    "Успех", f"Презентация опроса успешно сохранена:\n{os.path.abspath(save_path)}"
                ))
            except Exception as ex:
                self.after(0, lambda msg=str(ex): messagebox.showerror(
                    "Ошибка генерации", f"Произошла ошибка при сохранении файла:\n{msg}"
                ))
            finally:
                self.after(0, lambda: self.btn_generate.configure(state="normal", text="Сгенерировать отчет презентации"))

        threading.Thread(target=worker, daemon=True).start()