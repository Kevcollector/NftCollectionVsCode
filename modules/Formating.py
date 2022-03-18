from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


class excel:
    def writeToExcel(worksheet, data):
        for r in dataframe_to_rows(data, index=False):
            worksheet.append(r)
        dims = {}
        for row in worksheet.rows:
            for cell in row:
                if cell.value:
                    dims[cell.column_letter] = max(
                        (dims.get(cell.column_letter, 0), len(str(cell.value))))
        for col, value in dims.items():
            worksheet.column_dimensions[col].width = value
