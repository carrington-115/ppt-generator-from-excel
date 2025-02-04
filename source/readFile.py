import pandas as pd

"""
    - how use pandas to read the data of each sheet
    - store the all the sheets data of the excel file in a structured manner

"""

def print_file_data(file):
    print(file)

class readFile:
    def __init__(self, filename, sheets, sheet_data):
        self.filename = filename

    def __str__(self):
        return self.filename
    
    def readSheetData(self, file):
        data = ""
        return data