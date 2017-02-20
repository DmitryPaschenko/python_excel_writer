from dp_excel.ExcelRowOptions import ExcelRowOptions
from dp_excel.ExcelCell import ExcelCell


class ExcelRowTemplate:
    def __init__(self):
        self.rows = []

    def __get_current_row(self):
        return self.rows[-1]

    def add_column(self, value, options=None, is_empty=False):
        cell = ExcelCell(value, options, is_empty)
        self.__get_current_row()['columns'].append(cell)

        return self

    def add_row(self, options=None):
        if options:
            if not isinstance(options, ExcelRowOptions):
                raise ValueError('options must be is instance of ExcelRowOptions')
        else:
            options = ExcelRowOptions()

        self.rows.append({'columns': [], 'options': options})
        return self

    def get_rows(self):
        for row in self.rows:
            yield row

    def get_columns(self, row):
        for column in row['columns']:
            yield column

    def get_row_options(self, row):
        return row['options']

    def get_cell_options(self, column):
        return column.get('options', None)
