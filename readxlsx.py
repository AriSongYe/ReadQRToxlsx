import openpyxl
from openpyxl_image_loader import SheetImageLoader
from pyzbar.pyzbar import decode
from PIL import Image as PILImage

# 엑셀 파일 열기
workbook = openpyxl.load_workbook('QRTest.xlsx')
sheet = workbook['엑셀 구동']

# SheetImageLoader를 사용하여 이미지 읽어오기
image_loader = SheetImageLoader(sheet)

# C2 셀부터 시작하여 이미지 배열에 저장
start_row = 4
end_row = 9
image_array = []

for row in range(start_row, end_row + 1):
    cell_address = f'C{row}'
    
    # SheetImageLoader로 이미지 로드
    img_cell = image_loader.get(cell_address)

    if img_cell is not None:
        # 이미지 배열에 추가
        image_array.append(img_cell)

# 이미지를 파일로 저장하고 외부 이미지 뷰어로 열기
for idx, img in enumerate(image_array):
    img_path = f'image_{idx+1}.png'
    img.save(img_path)
    print(f"이미지 저장: {img_path}")

# 엑셀 파일 닫기
workbook.close()
