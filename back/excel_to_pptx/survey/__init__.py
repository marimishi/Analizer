from pptx import Presentation
from pptx.util import Inches

from ..title import create_title_slide
from .analyzer import SurveyAnalyzer
from .charts import create_chart_slides
from .conclusions import create_conclusion_slides
from .common_results import create_common_results_slide


class Generator:
    def __init__(self, df, title):
        self.df = df
        self.title = title
        self.prs = Presentation()
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)
        self.blank_layout = self.prs.slide_layouts[6]
        self.title_layout = self.prs.slide_layouts[0]
        
    def generate_all_slides(self, user_choices: dict = None):
        create_title_slide(self.prs, self.title, self.title_layout)
        create_common_results_slide(self.prs, self.df, self.blank_layout)
        create_chart_slides(self.prs, self.df, self.blank_layout)
        
        if user_choices:
            create_conclusion_slides(self.prs, user_choices, self.blank_layout)
        
        return self.prs
    
    def get_questions_for_ui(self) -> list:
        analyzer = SurveyAnalyzer(self.df)
        return analyzer.analyze()