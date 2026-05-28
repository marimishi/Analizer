from pptx import Presentation

from ..title import create_title_slide
from .common_results import create_common_results_slide
from .range_score import create_range_score_slide
from .questions_result import create_questions_slide
from .department import create_department_slides
from .top_10 import create_top_10_slide, create_top_10_success_slide
from .conclusion import create_conclusion_slide
from pptx.util import Inches

class Generator:
    def __init__(self, df, title):
        self.df = df
        self.title = title
        self.prs = Presentation() 
        
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)

    def generate_all_slides(self):
        title_layout = self.prs.slide_layouts[0]
        blank_layout = self.prs.slide_layouts[6]

        create_title_slide(self.prs, self.title, title_layout)
        create_common_results_slide(self.prs, self.df, blank_layout)
        create_range_score_slide(self.prs, self.df, blank_layout)
        create_questions_slide(self.prs, self.df, blank_layout)
        create_department_slides(self.prs, self.df, blank_layout)
        create_top_10_slide(self.prs, self.df, blank_layout)
        create_top_10_success_slide(self.prs, self.df, blank_layout)
        create_conclusion_slide(self.prs, self.df, blank_layout)
        
        return self.prs