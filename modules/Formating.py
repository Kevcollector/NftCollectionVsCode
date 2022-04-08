from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


class excel:
    def writeToExcel_dataframe(worksheet, data):
        """ method to write excel, takes the workheet and the data frame

        Args:
            worksheet (worksheet): Worksheet openpyxl object
            data (dataframe): DataFrame pandas object
        """

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

    def writeToExcel_dataframe(worksheet, data, add):
        """ method to write excel, takes the workheet and the data frame

        Args:
            worksheet (object): Worksheet openpyxl object
            data (dataframe): DataFrame pandas object
            add(int): adds a space onto the end of the cells
        """
        dims = {}
        for row in worksheet.rows:
            for cell in row:
                if cell.value:
                    temp = cell.value
                    dims[cell.column_letter] = max(
                        (dims.get(cell.column_letter, 0), len(str(cell.value))))
        for col, temp in dims.items():
            dataInSheet = temp
            worksheet.column_dimensions[col].width = dataInSheet + add
