import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from back.excel_to_pptx.constants import Constants

def create_top_10_slide(prs: Presentation, df: pd.DataFrame, layout):
    dep_col = "Укажите подразделение/отделение"
    if dep_col not in df.columns:
        return
        
    slide = prs.slides.add_slide(layout)
    
    for shape in list(slide.shapes):
        if shape.is_placeholder:
            sp = shape._element
            sp.getparent().remove(sp)

    header_line = slide.shapes.add_shape(1, *Constants.HEADER_LINE_GEOMETRY)
    header_line.fill.solid()
    header_line.fill.fore_color.rgb = Constants.COLOR_ACCENT
    header_line.line.fill.background()
    header_line.shadow.inherit = False

    tx_title = slide.shapes.add_textbox(*Constants.HEADER_TEXT_BOX_GEOMETRY)
    p_t = tx_title.text_frame.paragraphs[0]
    p_t.text = "Топ-10 подразделений по количеству участников"
    p_t.font.name = Constants.FONT_NAME
    p_t.font.size = Constants.HEADER_FONT_SIZE
    p_t.font.bold = True
    p_t.font.color.rgb = Constants.COLOR_PRIMARY

    top10 = df[dep_col].value_counts().head(10)
    
    chart_data = CategoryChartData()
    chart_data.categories = [str(name)[:20] + "..." if len(str(name)) > 20 else str(name) for name in top10.index]
    chart_data.add_series('Сотрудников', tuple(int(v) for v in top10.values))
    
    chart_x = Inches(0.8)
    chart_y = Inches(1.8)
    chart_width = Inches(11.7)
    chart_height = Inches(4.8)
    
    chart_shape = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, 
        chart_x, chart_y, chart_width, chart_height, 
        chart_data
    )
    chart = chart_shape.chart
    chart.has_legend = False
    
    series = chart.plots[0].series[0]
    series.format.fill.solid()
    series.format.fill.fore_color.rgb = Constants.COLOR_ACCENT
    
    category_axis = chart.category_axis
    category_axis.tick_labels.font.name = Constants.FONT_NAME
    category_axis.tick_labels.font.size = Constants.CHART_LABEL_SIZE
    category_axis.tick_labels.font.color.rgb = Constants.COLOR_PRIMARY

    value_axis = chart.value_axis
    value_axis.tick_labels.font.name = Constants.FONT_NAME
    value_axis.tick_labels.font.size = Constants.CHART_LABEL_SIZE
    value_axis.tick_labels.font.color.rgb = Constants.COLOR_SECONDARY
    
    plot = chart.plots[0]
    plot.has_data_labels = True
    data_labels = plot.data_labels
    data_labels.font.name = Constants.FONT_NAME
    data_labels.font.size = Constants.CHART_LABEL_SIZE
    data_labels.font.bold = True
    data_labels.font.color.rgb = Constants.COLOR_PRIMARY
    data_labels.number_format = '0'
    
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from back.excel_to_pptx.constants import Constants

def create_top_10_success_slide(prs: Presentation, df: pd.DataFrame, layout):
    dep_col = "Укажите подразделение/отделение"
    if dep_col not in df.columns:
        return
        
    slide = prs.slides.add_slide(layout)
    
    for shape in list(slide.shapes):
        if shape.is_placeholder:
            sp = shape._element
            sp.getparent().remove(sp)

    header_line = slide.shapes.add_shape(1, *Constants.HEADER_LINE_GEOMETRY)
    header_line.fill.solid()
    header_line.fill.fore_color.rgb = Constants.COLOR_ACCENT
    header_line.line.fill.background()
    header_line.shadow.inherit = False

    tx_title = slide.shapes.add_textbox(*Constants.HEADER_TEXT_BOX_GEOMETRY)
    p_t = tx_title.text_frame.paragraphs[0]
    p_t.text = "Топ-10 подразделений по успешности тестирования"
    p_t.font.name = Constants.FONT_NAME
    p_t.font.size = Constants.HEADER_FONT_SIZE
    p_t.font.bold = True
    p_t.font.color.rgb = Constants.COLOR_PRIMARY

    score_cols = [c for c in df.columns if "/ Баллы" in c]
    df['Total_Score'] = df[score_cols].apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1)
    max_score = len(score_cols)
    
    summary = df.groupby(dep_col).agg(
        Средний_балл=('Total_Score', 'mean')
    ).reset_index()
    
    summary['Успешность'] = (summary['Средний_балл'] / max_score * 100).round(1) if max_score > 0 else 0
    top10_success = summary.sort_values(by='Успешность', ascending=False).head(10)
    
    chart_data = CategoryChartData()
    chart_data.categories = [str(name)[:20] + "..." if len(str(name)) > 20 else str(name) for name in top10_success[dep_col]]
    chart_data.add_series('Успешность', tuple(float(v) for v in top10_success['Успешность']))
    
    chart_x = Inches(0.8)
    chart_y = Inches(1.8)
    chart_width = Inches(11.7)
    chart_height = Inches(4.8)
    
    chart_shape = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, 
        chart_x, chart_y, chart_width, chart_height, 
        chart_data
    )
    chart = chart_shape.chart
    chart.has_legend = False
    
    series = chart.plots[0].series[0]
    series.format.fill.solid()
    series.format.fill.fore_color.rgb = Constants.COLOR_ACCENT
    
    category_axis = chart.category_axis
    category_axis.tick_labels.font.name = Constants.FONT_NAME
    category_axis.tick_labels.font.size = Constants.CHART_LABEL_SIZE
    category_axis.tick_labels.font.color.rgb = Constants.COLOR_PRIMARY

    value_axis = chart.value_axis
    value_axis.maximum_scale = 100
    value_axis.tick_labels.font.name = Constants.FONT_NAME
    value_axis.tick_labels.font.size = Constants.CHART_LABEL_SIZE
    value_axis.tick_labels.font.color.rgb = Constants.COLOR_SECONDARY
    value_axis.number_format = '0'
    
    plot = chart.plots[0]
    plot.has_data_labels = True
    data_labels = plot.data_labels
    data_labels.font.name = Constants.FONT_NAME
    data_labels.font.size = Constants.CHART_LABEL_SIZE
    data_labels.font.bold = True
    data_labels.font.color.rgb = Constants.COLOR_PRIMARY
    
    data_labels.show_percentage = False
    data_labels.show_value = True
    data_labels.number_format = '0"%"'