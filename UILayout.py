import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook

class MyApp:
    def __init__(self, root):
        self.root = root
        self.root.title("서울반석교회_선교헌금_QRtoxlsxReader")
        self.file_path =  ""
        self.root.geometry("400x300")

        # 버튼 1: 데이터 전송
        btn_send_data = tk.Button(root, text="데이터 전송", command=self.send_data)
        btn_send_data.pack(pady=10)
        if self.file_path ==  "":
            print("file_path is Empty! Please Connect file_path")

        # 버튼 2: 엑셀 파일 경로 수정
        btn_modify_path = tk.Button(root, text="엑셀 파일 경로 수정", command=self.modify_excel_path)
        btn_modify_path.pack(pady=10)

        # 버튼 3: 나가기
        btn_exit = tk.Button(root, text="나가기", command=root.destroy)
        btn_exit.pack(pady=10)

    def send_data(self):
        # 데이터 전송 로직을 작성
        data_window = DataWindow(self.root)
        data_window.show()

    def modify_excel_path(self):
        # 엑셀 파일 경로 수정 로직을 작성
        self.file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx;*.xls")])
        if self.file_path:
            print(f"선택된 엑셀 파일 경로: {self.file_path}")

class DataWindow:
    def __init__(self, master):
        self.top = tk.Toplevel(master)
        self.top.title("Data Window")
        self.top.geometry("300x200")

        # Entry 위젯 추가
        self.entry_qr = tk.Entry(self.top)
        self.entry_qr.pack(pady=10)

        self.entry_qr.bind('<Return>', self.read_qr)
        # 버튼: Read QR
        btn_read_qr = tk.Button(self.top, text="Read QR", command=self.read_qr)
        btn_read_qr.pack(pady=10)

        # 라벨 초기화
        self.result_label = tk.Label(self.top, text="")
        self.result_label.pack(pady=10)

        # 버튼: 뒤로가기
        btn_back = tk.Button(self.top, text="뒤로가기", command=self.close_window)
        btn_back.pack(pady=10)

    def close_window(self):
        self.top.destroy()

    def read_qr(self):
        # Entry에서 QR 코드를 가져오기
        qr_data = self.entry_qr.get()

        # 이전에 생성한 라벨 삭제
        self.result_label.destroy()

        # 가져온 QR 코드 출력
        self.result_label = tk.Label(self.top, text=f"Read QR: {qr_data}")
        self.result_label.pack(pady=10)