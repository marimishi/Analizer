import pandas as pd
import re
import logging

logger = logging.getLogger(__name__)


class SurveyAnalyzer:
    def __init__(self, df: pd.DataFrame):
        self.df = df
        
    def is_technical_or_excluded_col(self, col_name: str) -> bool:
        col_clean = col_name.strip().lower()
        exclude_patterns = [
            r'\bid\b', r'идентификатор', r'фио', r'имя', r'фамилия',
            r'дата', r'время', r'ваши предложения'
        ]
        for pattern in exclude_patterns:
            if re.search(pattern, col_clean):
                return True
        return False

    def is_demographic_col(self, col_name: str) -> bool:
        col_clean = col_name.strip().lower()
        return 'пол' in col_clean or 'возраст' in col_clean

    def calculate_column_score(self, series: pd.Series) -> float:
        """Возвращает процент позитивных ответов"""
        valid_data = series.dropna().astype(str).str.lower().str.strip()
        if valid_data.empty:
            return 0.0

        total_count = len(valid_data)
        pos_patterns = [r'\bда\b', r'полностью', r'отлично', r'хорошо', 
                       r'удовлетворен', r'легко', r'рекомендую']

        pos_clicks = 0
        for val in valid_data:
            if any(p in val for p in pos_patterns) and "не " not in val and "плохо" not in val:
                pos_clicks += 1

        return pos_clicks / total_count

    def group_age_column(self):
        """Группирует возрастные значения"""
        import re
        
        age_col = None
        for col in self.df.columns:
            if 'возраст' in col.lower() or 'age' in col.lower():
                age_col = col
                break
        
        if not age_col:
            return
        
        def categorize_age(value):
            try:
                if isinstance(value, (int, float)):
                    age = int(value)
                else:
                    numbers = re.findall(r'\d+', str(value))
                    if not numbers:
                        return "Не указан"
                    age = int(numbers[0])

                if age < 18:
                    return "Младше 18"
                elif 18 <= age <= 29:
                    return "18-29"
                elif 30 <= age <= 49:
                    return "30-49"
                elif 50 <= age <= 60:
                    return "50-60"
                else:
                    return "Старше 60"
            except (ValueError, TypeError):
                return "Не указан"
        
        self.df[age_col] = self.df[age_col].apply(categorize_age)
        logger.info(f"Возраст сгруппирован в колонке: {age_col}")

    def analyze(self) -> list:
        """
        Анализирует DataFrame и возвращает список вопросов с предсказаниями
        """
        # Группируем возраст перед анализом
        self.group_age_column()
        
        raw_computed_cols = []
        
        for col in self.df.columns:
            if self.is_technical_or_excluded_col(col):
                continue

            valid_series = self.df[col].dropna().astype(str).str.strip()
            unique_answers = valid_series.unique()

            if len(unique_answers) > 8 or len(unique_answers) <= 1:
                continue

            is_demographic = self.is_demographic_col(col)
            score = 0.0 if is_demographic else self.calculate_column_score(self.df[col])

            raw_computed_cols.append({
                "column_name": col,
                "score": score,
                "is_demographic": is_demographic,
                "answers_preview": dict(valid_series.value_counts())
            })

        if not raw_computed_cols:
            return []

        scoring_questions = [q for q in raw_computed_cols if not q["is_demographic"]]
        scoring_questions.sort(key=lambda x: x["score"])

        total_valid_questions = len(scoring_questions)

        for index, item in enumerate(scoring_questions):
            if index < 3:
                item["predicted_aspect"] = "Проблемная область"
            elif index >= (total_valid_questions - 3):
                item["predicted_aspect"] = "Позитивный аспект"
            else:
                item["predicted_aspect"] = "Не включать в выводы"

        for item in raw_computed_cols:
            if item["is_demographic"]:
                item["predicted_aspect"] = "Не включать в выводы"

        return raw_computed_cols