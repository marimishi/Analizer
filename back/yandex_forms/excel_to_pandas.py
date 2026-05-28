import os
import re
import pandas as pd
from datetime import datetime

class ExcelToPandasProcessor:
    @staticmethod
    def excel_to_pandas(file_path: str) -> pd.DataFrame:
        """Читает Excel-файл и возвращает DataFrame."""
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Файл не найден: {file_path}")
        df = pd.read_excel(file_path)
        df.columns = [str(col).strip() for col in df.columns]
        return df
    
    @staticmethod
    def drop(df: pd.DataFrame) -> pd.DataFrame:
        """
        Удаляет все столбцы, кроме:
        - Время создания
        - Ваша должность
        - Укажите подразделение/отделение
        - Заканчивающихся на '/ Баллы'
        """
        columns_to_keep = []
        for col in df.columns:
            if col in ["Время создания", "Ваша должность", "Укажите подразделение/отделение"]:
                columns_to_keep.append(col)
            elif col.endswith("/ Баллы") or col.endswith("/Баллы"):
                columns_to_keep.append(col)
                
        return df[columns_to_keep].copy()
    
    @staticmethod
    def find_date_column(df_columns: list, possible_date_cols: list = None) -> str:
        """Поиск столбца с датой среди возможных вариантов"""
        if possible_date_cols is None:
            possible_date_cols = ["Время создания", "Дата создания", "Время задания", "Дата"]
        
        for col in possible_date_cols:
            if col in df_columns:
                return col
        return None
    
    @staticmethod
    def get_unique_dates(df: pd.DataFrame, date_column: str) -> list:
        """Получает уникальные даты из столбца"""
        if date_column not in df.columns:
            return []
        
        # Преобразуем в datetime и форматируем
        dates = pd.to_datetime(df[date_column], errors='coerce')
        unique_dates = dates.dropna().dt.strftime('%Y-%m-%d').unique()
        return sorted(unique_dates)
    
    @staticmethod
    def filter_by_dates(df: pd.DataFrame, date_column: str, selected_dates: list) -> pd.DataFrame:
        """Фильтрует DataFrame по выбранным датам"""
        if not selected_dates or date_column not in df.columns:
            return df
        
        df_copy = df.copy()
        df_copy['_date_filter'] = pd.to_datetime(df_copy[date_column], errors='coerce').dt.strftime('%Y-%m-%d')
        filtered_df = df_copy[df_copy['_date_filter'].isin(selected_dates)].copy()
        filtered_df.drop('_date_filter', axis=1, inplace=True)
        
        return filtered_df