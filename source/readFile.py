import pandas as pd

class ReadExcelFile:
    def __init__(self, file_location):
        self.file_location = file_location


    def loopThroughSheet(self, excel_sheet, sheet_name):
        return pd.read_excel(excel_sheet, sheet_name).fillna(' ').to_numpy()

    def readDataFromSheet(self, sheet_name):
        sheet_data = [[]]
        pandas_xls = pd.ExcelFile(self.file_location)
        sheet_data = self.loopThroughSheet(pandas_xls, sheet_name)
        return sheet_data

    # test print
    def __str__(self):
        return f'this is the excel file location {self.file_location}'