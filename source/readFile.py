import pandas as pd

class ReadExcelFile:
    def __init__(self, file_location):
        self.file_location = file_location


    def loopThroughSheet(self, excel_sheet, sheet_name):
        return pd.read_excel(excel_sheet, sheet_name).fillna(' ').to_numpy()

    def readDataFromSheet(self, sheet_name):
        sheet_data = []
        pandas_xls = pd.ExcelFile(self.file_location)
        sheet_data = self.loopThroughSheet(pandas_xls, sheet_name)
        return sheet_data

    def rowAndCol(self, sheet_name):
        data = self.readDataFromSheet(sheet_name)
        rows, cols = len(data), len(data[0])
        return rows, cols
    
    def sheetNames(self):
        return pd.ExcelFile(self.file_location).sheet_names
    
    def readMultipleSheetData(self):
        excel_file = pd.ExcelFile(self.file_location)
        all_data = dict({})
        for sheet in excel_file.sheet_names:
            df = pd.read_excel(excel_file, sheet)
            df = df.fillna(' ')
            sheet_data = df.to_numpy()
            all_data.setdefault(f'{sheet}', sheet_data)
        return all_data



    # test print
    def __str__(self):
        return f'this is the excel file location {self.file_location}'