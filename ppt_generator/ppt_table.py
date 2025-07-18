from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR,  MSO_VERTICAL_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

class ppt:
    MAX_CONTENT_HEIGHT = Inches(7.0)      # usable vertical space
    DEFAULT_TOP_OFFSET = Inches(0.8)      # starting top margin
    ELEMENT_SPACING = Inches(0.2)         # space between elements

    def __init__(self, filename):
        self.filename = filename
        self.prs = Presentation()
        self.current_slide = None
        self.chart_type = XL_CHART_TYPE.BAR_CLUSTERED
        self.slide_top_offset = self.DEFAULT_TOP_OFFSET

    def add_slide(self):
    # Use a fully blank slide layout (no title box)
        self.current_slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.slide_top_offset = self.DEFAULT_TOP_OFFSET

    def ensure_space(self, element_height):
        if self.current_slide is None or (self.slide_top_offset + element_height > self.MAX_CONTENT_HEIGHT):
            self.add_slide()

    def add_title(self, text):
        height = Inches(0.5)
        self.ensure_space(height)

        left = Inches(0.5)
        width = Inches(9)
        top = self.slide_top_offset

        textbox = self.current_slide.shapes.add_textbox(left, top, width, height)
        tf = textbox.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE  # vertical centering

        p = tf.paragraphs[0]
        p.text = text
        p.alignment = PP_ALIGN.CENTER          # horizontal centering
        p.font.size = Pt(28)

        self.slide_top_offset += height + self.ELEMENT_SPACING

    def add_table(self, data):
        if not data:
            return

        headers = list(data[0].keys())
        rows = len(data) + 1
        cols = len(headers)
        row_height = 0.3
        table_height = Inches(0.5 + row_height * len(data))

        self.ensure_space(table_height)

        left = Inches(0.5)
        width = Inches(9)
        top = self.slide_top_offset

        table = self.current_slide.shapes.add_table(rows, cols, left, top, width, table_height).table

        for col_idx, header in enumerate(headers):
            cell = table.cell(0, col_idx)
            cell.text = header
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            p.font.size = Pt(18)

        for row_idx, row_data in enumerate(data, start=1):
            for col_idx, key in enumerate(headers):
                cell = table.cell(row_idx, col_idx)
                cell.text = str(row_data[key])
                p = cell.text_frame.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER if col_idx == 1 else PP_ALIGN.LEFT

        self.slide_top_offset += table_height + Inches(1)  

    def add_graph(self, data):
        if not data:
            return

        headers = list(data[0].keys())
        label_key, value_key = headers[0], headers[1]

        categories = [d[label_key] for d in data]
        values = [d[value_key] for d in data]

        chart_data = CategoryChartData()
        chart_data.categories = categories
        chart_data.add_series(value_key, values)

        chart_height = Inches(4)
        chart_width = Inches(6)

        self.ensure_space(chart_height)

        x = Inches(1.5)
        y = self.slide_top_offset
        cx = chart_width
        cy = chart_height

        self.current_slide.shapes.add_chart(self.chart_type, x, y, cx, cy, chart_data)
        self.slide_top_offset += chart_height + self.ELEMENT_SPACING

    def jobs_cancelled_add_graph(self,data, chart_type=XL_CHART_TYPE.COLUMN_CLUSTERED):
        if not data:
            return

        if not self.current_slide:
            self.add_slide()

        headers = list(data[0].keys())
        if len(headers) < 2:
            raise ValueError("Data must contain at least two keys per dictionary.")

        date_key, job_key, cancelled_key = headers[0], headers[1], headers[2]
        print(headers)
        try:
            categories = [d[date_key] for d in data]
            job_total = [d[job_key] for d in data]
            job_cancelled = [d[cancelled_key] for d in data]

        except KeyError as e:
            raise ValueError(f"Key '{e.args[0]}' not found in data items.")
        print(job_total)

        chart_data = CategoryChartData()
        chart_data.categories = categories
        cancelled_total = sum(job_cancelled)
        job_total_sum = sum(job_total)

        #summary textbox dimension/parameters
        left = Inches(6.5)
        top = Inches(6.2)
        width = Inches(3)
        height = Inches(1.2)

        textbox = self.current_slide.shapes.add_textbox(left, top, width, height)
        text_frame = textbox.text_frame
        text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
        if text_frame.paragraphs and not text_frame.paragraphs[0].text.strip():
         text_frame.paragraphs[0]._element.getparent().remove(text_frame.paragraphs[0]._element)
        p = text_frame.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        p.text = "Total Job Received Count"
        p.font.bold = True
        p.font.size = Pt(16)
        p = text_frame.add_paragraph()
        p.text = f"Total Jobs: {job_total_sum:,}\nCancelled Jobs: {cancelled_total:,}"
        p.font.size = Pt(14)
        p = text_frame.add_paragraph()
        p.text = "Note: \n - Data shown includes possible duplicate submissions"
        p.font.size = Pt(8)
        fill = textbox.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(204, 255, 204)
        line = textbox.line
        line.color.rgb = RGBColor(0, 100, 0)

        chart_data.add_series(cancelled_key, job_cancelled)
        chart_data.add_series(job_key, job_total)

        x, y, cx, cy = Inches(0.5), Inches(1), Inches(9), Inches(6)
        chart_frame = self.current_slide.shapes.add_chart(chart_type, x, y, cx, cy, chart_data)
        chart = chart_frame.chart

        colors = [RGBColor (245,105,36),
                RGBColor(0,66,99)]   

        category_axis = chart.category_axis
        category_labels_font = category_axis.tick_labels.font
        category_labels_font.size = Pt(10)
        category_labels_font.name = 'Arial'


        #adds value at the top of graph
        for series in chart.series:
            series.has_data_labels = True
            data_labels = series.data_labels
            data_labels.show_value = True
            data_labels.number_format = '#,##0'
            data_labels.font.size = Pt(12)
            data_labels.font.name = 'Arial'
        
        for idx, series in enumerate(chart.series):
            fill = series.format.fill
            fill.solid()
            fill.fore_color.rgb = colors[idx]
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
                
    def add_SLA_graph(self, data, chart_type=XL_CHART_TYPE.COLUMN_CLUSTERED):
        if not data:
            return

        if not self.current_slide:
            self.add_slide()

        headers = list(data[0][0].keys())
        jobs = list(data[1][0].values())

        if len(headers) < 2:
            raise ValueError("Data must contain at least two keys per dictionary.")

        label_key, value_key = headers[0], headers[2]
        try:
            categories = [d[label_key] for d in data[0]]
            values = [d[value_key] for d in data[0]]
        except KeyError as e:
            raise ValueError(f"Key '{e.args[0]}' not found in data items.")

        chart_data = CategoryChartData()
        chart_data.categories = categories
        chart_data.add_series(value_key, values)

        x, y, cx, cy = Inches(1), Inches(1), Inches(8), Inches(6)
        chart_frame = self.current_slide.shapes.add_chart(chart_type, x, y, cx, cy, chart_data)
        chart = chart_frame.chart
        #adds value at the top of graph
        for series in chart.series:
            series.has_data_labels = True
            data_labels = series.data_labels
            data_labels.show_value = True
            data_labels.number_format = '#,##0'
            data_labels.font.size = Pt(12)
            data_labels.font.name = 'Arial'

    def add_SLA_table(self, data):
        if not data:
            print("SLA table data is empty")
            return
        
        rows = len(data[0]) + 1
        cols = len(data[0][0])
        left = Inches(0.5)
        top = Inches(1.5)
        width = Inches(9)
        height = Inches(0.8 + 0.3 * len(data[0]))

        table = self.current_slide.shapes.add_table(rows, cols, left, top, width, height).table

        # Header
        headers = list(data[0][0].keys())
        for col_idx, header in enumerate(headers):
            cell = table.cell(0, col_idx)
            cell.text = header
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            p.font.size = Pt(18)

        # Rows
        for row_idx, row_data in enumerate(data[0], start=1):
            for col_idx, key in enumerate(headers):
                cell = table.cell(row_idx, col_idx)

                if isinstance(row_data[key], (int, float)):
                    cell.text = f"{row_data[key]:,.0f}"  # No decimal places
                else:
                    cell.text = str(row_data[key])

                p = cell.text_frame.paragraphs[0]
                # Center only second column (index 1)
                if col_idx == 1:
                    p.alignment = PP_ALIGN.CENTER
                else:
                    p.alignment = PP_ALIGN.LEFT  
                    
        summary_top = height + top + Inches(0.3)
        summary_left = left
        summary_width = Inches(4)
        summary_height = Inches(2)
        jobs = list(data[1][0].values())

        #Summary Text Box
        summary_text = (
            f"Job SLA Status\nTotal Done Jobs: {jobs[0]}\n(includes duplicate hash jobs)\nJobs Done Within SLA: {jobs[1]}\nJobs Done Outside SLA: {jobs[0]-jobs[1]}\nOverall SLA Compliance: {((jobs[1]/jobs[0])*100):.2f}%"
        )
        textbox = self.current_slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
            summary_left, summary_top, summary_width, summary_height)
        
        text_frame = textbox.text_frame
        text_frame.clear()
        text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
        text_frame.margin_top = 0

        p = text_frame.add_paragraph()
        p.text = summary_text
        if text_frame.paragraphs and not text_frame.paragraphs[0].text.strip():
            text_frame.paragraphs[0]._element.getparent().remove(text_frame.paragraphs[0]._element)
        p.font.size = Pt(14)
        p.alignment = PP_ALIGN.LEFT
        p.font.name = 'Arial'
        p.font.color.rgb = RGBColor(0,0,0)
        p.number_format = '#,##0'
        
        fill = textbox.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(204, 255, 204)
        line = textbox.line
        line.color.rgb = RGBColor(0, 100, 0)

    def save(self):
        self.prs.save(self.filename)
