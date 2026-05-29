from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

class Constants:
# --- БАЗОВАЯ ТИПОГРАФИКА И ЦВЕТА ---
    FONT_NAME = "Arial"
    COLOR_PRIMARY = RGBColor(17, 24, 39)       # rgb(17, 24, 39)
    COLOR_SECONDARY = RGBColor(107, 114, 128) # rgb(107, 114, 128)
    COLOR_ACCENT = RGBColor(59, 130, 246)     # rgb(59, 130, 246)
    COLOR_MUTED_TAG = RGBColor(0x9C, 0xA3, 0xAF) 
        
    COLOR_THIRDLY = RGBColor(201, 255, 116)   # rgb(201, 255, 116)
    COLOR_WARNING = RGBColor(255, 255, 179) #rgb(255, 255, 179)
    COLOR_ERROR = RGBColor(255, 167, 124) #rgb(255, 167, 124)
    COLOR_TILE_BG = RGBColor(253, 254, 255)

    # --- ТИПОВОЙ СТАНДАРТНЫЙ ЗАГОЛОВОК СЛАЙДОВ ---
    HEADER_LINE_GEOMETRY = (Inches(0.6), Inches(0.5), Inches(0.04), Inches(0.6))
    HEADER_TEXT_BOX_GEOMETRY = (Inches(0.8), Inches(0.45), Inches(11.5), Inches(0.8))
    HEADER_FONT_SIZE = Pt(28)

    # --- ГЕОМЕТРИЯ ДИАГРАММЫ (Правые 50% экрана) ---
    CHART_GEOMETRY = (Inches(6.7), Inches(1.5), Inches(6.0), Inches(4.8))
    CHART_LEGEND_SIZE = Pt(12)
    CHART_LABEL_SIZE = Pt(14)
    CHART_NUMBER_FORMAT = '0"%"'

    # --- НОВЫЙ СТИЛЬ: ЧИСТЫЕ МЕТРИКИ (БЕЗ ПЛИТОК) ---
    METRIC_START_X = Inches(0.6)
    METRIC_START_Y = Inches(1.7)
    METRIC_WIDTH = Inches(5.5)
    METRIC_HEIGHT = Inches(1.1)
    METRIC_GAP = Inches(0.3)  # Отступ между блоками данных

    # Шрифты для метрик
    METRIC_VALUE_SIZE = Pt(44)  # Огромные, акцентирующие на себе внимание цифры
    METRIC_LABEL_SIZE = Pt(20)  # Аккуратная подпись под ними

    # --- ПЛИТКИ (ЗАКЛЮЧЕНИЕ) ---
    TILE_W = Inches(5.8)
    TILE_H = Inches(2.7)
    TILE_X_LEFT = Inches(0.6)
    TILE_X_RIGHT = Inches(6.9)
    TILE_Y_TOP = Inches(1.2)
    TILE_Y_BOTTOM = Inches(4.1)
    TILE_LINE_WIDTH = Pt(1.5)
    
    FONT_SIZE_TAG = Pt(16)
    FONT_SIZE_MAIN = Pt(36)
    FONT_SIZE_SUB = Pt(20)
    
    SURVEY_COLORS = [
    RGBColor(23, 87, 87),    # '#175757'
    RGBColor(61, 158, 118),  # '#3d9e76'
    RGBColor(184, 205, 171), # '#b8cdab'
    RGBColor(251, 242, 196), # '#fbf2c4'
    RGBColor(229, 193, 133), # '#e5c185'
    RGBColor(199, 82, 42)    # '#c7522a'
]