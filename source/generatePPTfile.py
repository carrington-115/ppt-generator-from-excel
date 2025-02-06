import pptx as Presentation

class generatePPTFromFile:
    def __init__(self, filename):
        self.filename = filename

    def __str__(self):
        return self.filename