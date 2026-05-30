import customtkinter as ctk
import pandas as pd
from back.excel_to_pptx.survey.analyzer import SurveyAnalyzer
from back.handler import SurveyHandler
from ui.config_screen import BaseConfigScreen


class SurveyConfigScreen(BaseConfigScreen):
    def __init__(self, master, settings_manager):
        super().__init__(master, settings_manager)
        
        self.questions_data = []
        self.current_index = 0
        self.user_choices = {}
        
        # Заголовок
        self.title_label = ctk.CTkLabel(
            self, 
            text="Настройки анализа Опросов", 
            font=ctk.CTkFont(family="Arial", size=20, weight="bold")
        )
        self.title_label.grid(row=0, column=0, sticky="w", pady=(0, 15), padx=5)
        
        # Карточка с вопросами
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

        self.create_generate_button("Сгенерировать презентацию", row=11)

    def load_file_structure(self, file_path):
        self.file_path = file_path
        
        # Очищаем название от дат
        from back.handler import clean_filename_from_date
        clean_title = clean_filename_from_date(file_path)
        self.title_entry.delete(0, "end")
        self.title_entry.insert(0, clean_title)
        
        try:
            if file_path.endswith('.csv'):
                self.df = pd.read_csv(file_path)
            else:
                self.df = pd.read_excel(file_path)
            
            analyzer = SurveyAnalyzer(self.df)
            self.questions_data = analyzer.analyze()
            
            self.current_index = 0
            self.user_choices.clear()
            for q in self.questions_data:
                self.user_choices[q["column_name"]] = q["predicted_aspect"]
                
            self.update_current_item_view()
            
        except Exception as e:
            self.show_error_message(f"Не удалось обработать опрос:\n{e}")
    
    def update_current_item_view(self):
        if not self.questions_data:
            self.question_title.configure(text="Подходящие вопросы для диаграмм не найдены.")
            self.counter_label.configure(text="0 / 0")
            return
            
        current_q = self.questions_data[self.current_index]
        col_name = current_q["column_name"]
        
        self.question_title.configure(text=f"Вопрос ({self.current_index + 1}/{len(self.questions_data)}):\n{col_name}")
        self.counter_label.configure(text=f"{self.current_index + 1} / {len(self.questions_data)}")
        
        answers_preview = current_q.get("answers_preview", {})
        
        if answers_preview:
            total = sum(answers_preview.values())
            preview_text = "Текущее распределение ответов в файле:\n"
            
            for ans, count in list(answers_preview.items())[:5]:
                percentage = (count / total) * 100
                bar_chars = int(percentage / 10)
                bar = "█" * bar_chars + "░" * (10 - bar_chars)
                
                short_ans = str(ans) if len(str(ans)) <= 40 else str(ans)[:37] + "..."
                preview_text += f"  {bar}  {percentage:5.1f}%  —  {short_ans} ({count} шт.)\n"
        else:
            preview_text = "Нет данных для预览"
            
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
        if self.questions_data:
            col_name = self.questions_data[self.current_index]["column_name"]
            self.user_choices[col_name] = self.radio_var.get()

    def next_item(self):
        if self.current_index < len(self.questions_data) - 1:
            self.current_index += 1
            self.update_current_item_view()

    def prev_item(self):
        if self.current_index > 0:
            self.current_index -= 1
            self.update_current_item_view()

    def run_generation(self):
        if not self.file_path or not self.questions_data:
            self.show_warning_message("Нет данных для генерации отчета.")
            return

        title_text = self.get_title()
        default_filename = f"{title_text}.pptx"
        
        save_path = self.save_file_dialog(default_filename)
        if not save_path:
            return

        def generate():
            handler = SurveyHandler(
                df=self.df,
                title=title_text,
                user_choices=self.user_choices,
                file_path=self.file_path  # ← передаем file_path
            )
            return handler.generate(save_path)
        
        self.run_in_thread(generate, on_success=self.show_success_message)