import openpyxl
from openpyxl_image_loader import SheetImageLoader
from pyzbar.pyzbar import decode
from PIL import Image as PILImage

class ExelManager:
    def __init__(self, filepath, sheetName):
        # OpenExel
        self.workbook = openpyxl.load_workbook(filepath)
        # Sheet 지정
        self.sheet = self.workbook[sheetName]

    def close(self):
        self.workbook.close()

    def numtocell(self):
        pass

    def write(self, coordinate, data):
        self.sheet[coordinate] = data