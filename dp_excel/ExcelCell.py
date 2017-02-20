from dp_excel.ExcelCellOptions import ExcelCellOptions


class ExcelCell:
    def __init__(self, value, options=None, is_empty=False):
        self.value = value
        self.is_empty = is_empty

        if options:
            if isinstance(options, ExcelCellOptions):
                self.options = options
            else:
                raise ValueError('options must be ExcelCellOptions instance')
        else:
            self.options = ExcelCellOptions()
