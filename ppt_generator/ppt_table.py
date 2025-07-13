from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

class ppt:
    def __init__(self, filename):
        self.filename = filename
        self.prs = Presentation()
        self.current_slide = None

    def add_slide(self):
        self.current_slide = self.prs.slides.add_slide(self.prs.slide_layouts[5])  # blank layout

    def add_title(self, text):
        title = self.current_slide.shapes.title
        title.text = text
        paragraph = title.text_frame.paragraphs[0]
        run = paragraph.runs[0]
        run.font.size = Pt(28)

    def add_table(self, data):
        rows = len(data) + 1
        cols = len(data[0])
        left = Inches(0.5)
        top = Inches(1.5)
        width = Inches(9)
        height = Inches(0.8 + 0.3 * len(data))

        table = self.current_slide.shapes.add_table(rows, cols, left, top, width, height).table

        # Header
        headers = list(data[0].keys())
        for col_idx, header in enumerate(headers):
            cell = table.cell(0, col_idx)
            cell.text = header
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            p.font.size = Pt(18)

        # Rows
        for row_idx, row_data in enumerate(data, start=1):
            for col_idx, key in enumerate(headers):
                cell = table.cell(row_idx, col_idx)
                cell.text = str(row_data[key])
                p = cell.text_frame.paragraphs[0]
                # Center only second column (index 1)
                if col_idx == 1:
                    p.alignment = PP_ALIGN.CENTER
                else:
                    p.alignment = PP_ALIGN.LEFT  

    def add_graph(self, data, chart_type=XL_CHART_TYPE.COLUMN_CLUSTERED):
        if not data:
            return

        if not self.current_slide:
            self.add_slide()

        headers = list(data[0].keys())
        if len(headers) < 2:
            raise ValueError("Data must contain at least two keys per dictionary.")

        label_key, value_key = headers[0], headers[1]
        try:
            categories = [d[label_key] for d in data]
            values = [d[value_key] for d in data]
        except KeyError as e:
            raise ValueError(f"Key '{e.args[0]}' not found in data items.")

        chart_data = CategoryChartData()
        chart_data.categories = categories
        chart_data.add_series(value_key, values)

        x, y, cx, cy = Inches(2), Inches(2), Inches(6), Inches(4)
        self.current_slide.shapes.add_chart(chart_type, x, y, cx, cy, chart_data)
    
    def save(self):
        self.prs.save(self.filename)
