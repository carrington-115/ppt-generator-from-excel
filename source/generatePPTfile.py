from pptx import Presentation
from pptx.util import Inches

prs = Presentation()

class GeneratePPTFromFile:
    def __init__(self, file_name, excel_file_data,):
        self.file_name = file_name
        self.excel_file_data = excel_file_data


    def addSlideToPresentation(self, slide_position):
        slide_layout = prs.slide_layouts[slide_position] # blank slide
        slide_info = prs.slide.add_slide(slide_layout)
        return slide_info
    
    def enterTableData(self, rows, cols, table, sheet_data):
        for i in range(rows):
            for j in range(cols):
                table.cell(i, j).text = str(sheet_data[i][j])

    def generatePPTForSingleSheet(self, ppt_location, row_num, col_num):
        # defined the sheet sizes
        height = Inches(6.49)
        width = Inches(12.48)
        top = Inches(0.5)
        left = Inches(0.5)
        slide_info = self.addSlideToPresentation(5)
        rows, cols = row_num, col_num
        # create the table
        table_data = slide_info.shapes.add_table(rows, cols, left, top, width, height).table
        
        # get the sheet data
        sheet_data =[[], []]
        self.enterTableData(rows, cols, table_data, sheet_data)
        # enter slide_data
        prs.save(ppt_location)

    def sayHello(self): 
        print('hello world')

    # this method is used for a print action in the class
    def __str__(self):
        return self.filename