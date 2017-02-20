Generate excel files.
Based on "openpyxl" - A Python library to read/write Excel 2010 xlsx/xlsm files

Install:

1) pip install openpyxl
2) Copy folder "excel" to you project or create GIT submodules


USING EXAMPLE:


from excel.ExcelFile import ExcelFile
from excel.ExcelRowTemplate import ExcelRowTemplate
from excel.ExcelCellOptions import ExcelCellOptions
from excel.ExcelRowOptions import ExcelRowOptions
from openpyxl.styles import Side


class DownloadReportView(View):
    def get(self, request, report_id):
        report = get_object_or_404(Report, pk=report_id)

        excel_file = ExcelFile('test1')

        border = Side(border_style='thin', color='000000')

        counter_option = ExcelCellOptions()
        counter_option\
            .set_rowspan(2)\
            .set_alignment(vertical='center', horizontal='center')\
            .set_width('5')\
            .set_border(perimeter=border)

        bordered_option = ExcelCellOptions()
        bordered_option\
            .set_font(size=8) \
            .set_border(perimeter=border)\
            .set_width(30)\
            .set_alignment(wrap_text=True, shrink_to_fit=True, vertical='top')

        body_cell_options = ExcelCellOptions()
        body_cell_options.set_font(size=4)\
            .set_colspan(3) \
            .set_alignment(horizontal='center', wrap_text=True, shrink_to_fit=True)\
            .set_border(perimeter=border)

        body_row_options = ExcelRowOptions()
        body_row_options.set_height(50)

        counter = 1
        for post in report.report_posts.all():
            template = ExcelRowTemplate()
            template.add_row() \
                .add_column(counter, counter_option)\
                .add_column(post.title, bordered_option) \
                .add_column(post.get_notices_string(), bordered_option) \
                .add_column(post.get_attributes_string(), bordered_option) \
                .add_row(body_row_options)\
                .add_column(value='', is_empty=True)\
                .add_column(post.body.strip(), body_cell_options)

            excel_file.render_row_by_template(template)
            counter += 1

        return HttpResponse(excel_file.get_virtual_workbook(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

