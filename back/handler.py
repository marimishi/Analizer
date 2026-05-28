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
    ):
        self.file_path = file_path
        self.date_column = date_column
        self.selected_dates = selected_dates
        self.job_column = job_column
        self.title = title
        self.settings_manager = settings_manager
        self.job_groups = job_groups or {}
        self.df = df

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


        from back.excel_to_pptx.yandex_forms import Generator
        
        generator = Generator(df, self.title)
        prs = generator.generate_all_slides()
        
        output_path = "готовая_презентация.pptx"
        PptxSaver.save(prs, output_path)
        return output_path


class Survey:
    """Заглушка для будущего функционала Опросов"""
    def __init__(self):
        pass