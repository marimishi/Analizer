import os
from pptx import Presentation

class PptxSaver:
    @staticmethod
    def save(prs: Presentation, output_path: str):
        if not output_path.endswith(".pptx"):
            output_path += ".pptx"
            
        prs.save(output_path)
        print(f"Презентация успешно сохранена как: {output_path}")