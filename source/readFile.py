import pandas as pd

class readFile:
    def __init__(self, filename):
        self.filename = filename

    def __str__(self):
        return self.filename
    
    def readTextFile(file):
        pd.read_csv(
            file,
            sep='\t',
            lineterminator='\n',
            header=None
        )
    

