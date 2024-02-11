import openpyxl
from openpyxl_image_loader import SheetImageLoader
import re

class ExcelManager:
    def __init__(self, filepath):
        # OpenExel
        self.workbook = openpyxl.load_workbook(filepath)
        self.filepath = filepath
        # Sheet 지정
        self.QRsheet = self.workbook["QR인쇄용"]
        self.dataSheet = self.workbook["수정본"]

    def setSheet(self, sheetName):
        self.sheet = self.workbook[sheetName]

    def close(self):
        self.workbook.close()

    def numtocell(self):
        pass

    def write(self, coordinate, data, sheet):
        self.sheet[coordinate] = data

    def findCoordinate(self, target_row, target_value, sheet):
        for row in sheet.iter_rows(min_row=target_row, max_row=target_row):
            for cell in row:
                if cell.value == target_value:
                    return cell 
            else:
                continue
            break  # 값 찾았으므로 더 이상 반복할 필요 없음
        else:
            print(f"찾을 값 '{target_value}'은(는) 행 {target_row}에서 발견되지 않았습니다.")

    def findCoordinate_in_column(self, target_column, target_value, sheet, max_rows=None):
        i = 0
        for column in sheet.iter_cols(min_col=target_column, max_col=target_column, max_row=max_rows):
            for cell in column:
                i += 1

                if cell.value == target_value:
                    return cell 
            else:
                continue
            break  # 값 찾았으므로 더 이상 반복할 필요 없음
        else:
            print(f"찾을 값 '{target_value}'은(는) 행 {target_column}에서 발견되지 않았습니다.")

    def nth_row_coordinate(self, cell_coordinate, n):
        # 정규표현식을 사용하여 열과 행을 분리합니다
        match = re.match(r'([A-Z]+)(\d+)', cell_coordinate)
        if match:
            row = int(match.group(2)) + n
            if row > 0:
                return match.group(1) + str(row)
            else:
                raise ValueError("Row index must be positive")
        else:
            raise ValueError("Invalid cell coordinate format")
    
    def nth_column_coordinate(self, cell_coordinate, n):
        # 정규표현식을 사용하여 열과 행을 분리합니다
        match = re.match(r'([A-Z]+)(\d+)', cell_coordinate)
        if match:
            column = match.group(1)
            # 열 좌표에서 n을 더합니다
            new_column = self.excel_column_to_number(column) + n
            if new_column > 0:
                # 새 열 좌표를 다시 문자열로 변환합니다
                new_column_str = self.number_to_excel_column(new_column)
                return new_column_str + match.group(2)
            else:
                raise ValueError("Column index must be positive")
        else:
            raise ValueError("Invalid cell coordinate format")
    
    def excel_column_to_number(self, column_str):
        result = 0
        for char in column_str:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result

    def number_to_excel_column(self, n):
        result = ''
        while n > 0:
            remainder = (n - 1) % 26
            result = chr(ord('A') + remainder) + result
            n = (n - 1) // 26
        return result
    
    def extract_numeric_part(self, cell_coordinate):
        numeric_part = ""
        for char in cell_coordinate:
            if char.isdigit():
                numeric_part += char
        return numeric_part


