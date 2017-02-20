from openpyxl.worksheet.dimensions import RowDimension


class ExcelRowOptions:
    OPTION_ROW_DIMENSION = 'row_dimension'
    OPTION_ROW_HEIGHT = 'height'

    def __init__(self):
        self.options = []

    def _add_option(self, option, value):
        self.options.append((option, value))
        return self

    def get_options(self):
        for option in self.options:
            yield option

    def set_row_dimension(self, row_dimension):
        if isinstance(row_dimension, RowDimension):
            self._add_option(self.OPTION_ROW_DIMENSION, row_dimension)
        else:
            raise ValueError('row_dimension option must be is instance of openpyxl.worksheet.dimensions.RowDimension')

        return self

    def set_height(self, height):
        self._add_option(self.OPTION_ROW_HEIGHT, height)
        return self

    def apply_row_dimension(self, worksheet, value, row_number):
        worksheet.row_dimensions[row_number] = value
        return self

    def apply_height(self, worksheet, value, row_number):
        worksheet.row_dimensions[row_number].height = value
        return self

    def apply_row_options(self, worksheet, row_number):
        for option, value in self.get_options():
            if option == self.OPTION_ROW_DIMENSION:
                self.apply_row_dimension(worksheet, value, row_number)
            elif option == self.OPTION_ROW_HEIGHT:
                self.apply_height(worksheet, value, row_number)
