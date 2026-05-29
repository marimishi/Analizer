import os
import re
import pandas as pd
from back.yandex_forms.excel_to_pandas import ExcelToPandasProcessor
from back.excel_to_pptx.save_pptx import PptxSaver

class YandexFormsHandler:
    def __init__(
        self,
        file_path,
        date_column=None,
        selected_dates=None,
        job_column=None,
        title=None,
        settings_manager=None,
        job_groups=None,
        df=None,
        output_path=None,
    ):
        self.file_path = file_path
        self.date_column = date_column
        self.selected_dates = selected_dates
        self.job_column = job_column
        self.title = title
        self.settings_manager = settings_manager
        self.job_groups = job_groups or {}
        self.df = df
        self.output_path = output_path

    @staticmethod
    def clean_filename_from_date(file_path: str) -> str:
        """
        Извлекает имя файла без расширения и удаляет из него даты 
        в форматах YYYY-MM-DD, DD.MM.YYYY и т.д.
        """
        if not file_path:
            return ""
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        cleaned = re.sub(r'\b\d{4}[-./]\d{2}[-./]\d{2}\b|\b\d{2}[-./]\d{2}[-./]\d{4}\b', '', base_name)
        cleaned = re.sub(r'\s*[-_]\s*$', '', cleaned)
        cleaned = ' '.join(cleaned.split())
        return cleaned if cleaned else "Отчет по результатам тестирования"

    def process(self):
        if self.df is not None:
            df = self.df
        else:
            df = ExcelToPandasProcessor.excel_to_pandas(self.file_path)
            df = ExcelToPandasProcessor.drop(df)

            if self.date_column and self.selected_dates:
                df = ExcelToPandasProcessor.filter_by_dates(df, self.date_column, self.selected_dates)

        if df.empty:
            raise ValueError("После фильтрации по датам не осталось данных для анализа.")

        if self.job_column and self.job_groups and self.job_column in df.columns:
            df[self.job_column] = df[self.job_column].map(self.job_groups).fillna(df[self.job_column])

        final_title = self.title.strip() if self.title else ""
        if not final_title:
            final_title = self.clean_filename_from_date(self.file_path)

        from back.excel_to_pptx.yandex_forms import Generator
        
        generator = Generator(df, final_title)
        prs = generator.generate_all_slides()
        
        if not self.output_path:
            self.output_path = final_title + ".pptx"
            
        PptxSaver.save(prs, self.output_path)
        return self.output_path


from pptx import Presentation
from back.excel_to_pptx.survey.survey_diagrams import survey_base_diagram
from back.excel_to_pptx.survey.conclusion import create_survey_conclusions

class Survey:
    def __init__(self, df: pd.DataFrame, prs: Presentation, blank_layout):
        self.df = df
        self.prs = prs
        self.blank_layout = blank_layout
        self.charts_config = []

    def generate_base_diagrams(self, file_path: str, title_text: str = None) -> list:
        self.charts_config = survey_base_diagram(
            df=self.df, 
            prs=self.prs, 
            file_path=file_path,
            title_text=title_text
        )
        return self.charts_config