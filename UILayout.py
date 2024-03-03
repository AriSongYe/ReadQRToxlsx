import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from openpyxl import load_workbook
from readxlsx import ExcelManager
from DateFucntion import get_current_sunday

class MyApp:
    ExcelManager = 0

    def __init__(self, root):
        self.root = root
        self.root.title("서울반석교회 선교헌금 Writer")
        self.file_path =  ""
        self.root.geometry("500x400")


        # UI
        btn_send_data = tk.Button(root, text="데이터 전송", command=self.send_data)
        btn_send_data.pack(pady=10)


        btn_modify_path = tk.Button(root, text="엑셀 파일 경로 수정", command=self.modify_excel_path)
        btn_modify_path.pack(pady=10)

        btn_modify_path = tk.Button(root, text="나가기", command=self.close_app)
        btn_modify_path.pack(pady=10)

        self.Menual_label = tk.Label(self.root, text="메뉴얼\n\n1. Excel 연결 시 Open 시간이 조금 걸립니다! \n\n2. Save는 창을 닫을 시 자동으로 됩니다! \n\n3. 참고로 Save는 시간이 조금 걸립니다!\n\n4. 해당 주일날만 체크 됩니다!\n\n5. 특이사항은 수기로 엑셀을 수정해야 합니다.\n\n6. 엑셀파일이 열린 상태로 save는 안됩니다 \n\n7. Save오류로 File이 손상이 될경우가 있으니 항상 백업하시길 바랍니다")
        self.Menual_label.pack(pady=10)

    def send_data(self):
        if  MyApp.ExcelManager == 0:
            messagebox.showinfo("File Connect Error", "xlsx파일 path를 연결해주세요!")
        else:    
            data_window = DataWindow(self.root, MyApp.ExcelManager)


    def modify_excel_path(self):
        self.file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx;*.xls")])
        if self.file_path:
            print(f"선택된 엑셀 파일 경로: {self.file_path}")
            MyApp.ExcelManager = ExcelManager(self.file_path)
            messagebox.showinfo("Success connecting a excel file!", "성공적으로 엑셀 파일을 연결했습니다! ")
            
    def close_app(self):
        if MyApp.ExcelManager:
            MyApp.ExcelManager.workbook.save(MyApp.ExcelManager.filepath)
            MyApp.ExcelManager.workbook.close()
            messagebox.showinfo("Success Save and Close", "성공적으로 데이터를 저장했습니다!")
            print("Success Save and Close!")
        self.root.destroy()

        

class DataWindow():
    categoryArr = ["일반QR","신학생QR","선교사QR","개척교회QR","북한QR"]
    casheCoordi = ""
    name = ""

    def __init__(self, master, ExcelManager):

        self.top = tk.Toplevel(master)
        self.top.title("Data Window")
        self.top.geometry("300x400")
        self.top.iconbitmap("icon.ico")
        self.ExcelManager = ExcelManager

        self.entry_qr = tk.Entry(self.top)
        self.entry_qr.pack(pady=10)
        self.entry_qr.bind('<Return>', self.read_qr)
        
        btn_read_qr = tk.Button(self.top, text="Read QR", command=self.read_qr)
        btn_read_qr.pack(pady=10)


        self.entry_fix = tk.Entry(self.top)
        self.entry_fix.pack(pady=10)
        self.entry_fix.bind('<Return>', self.fixMoney)
        
        self.result_label = tk.Label(self.top, text="")
        self.result_label.pack(pady=10)

        btn_exit = tk.Button(self.top, text="Save Data", command=self.saveData)
        btn_exit.pack(pady=10)

    def close_window(self):
        self.top.destroy()

    def fixMoney(self, event=None):
        if self.casheCoordi != "":
            fixData = self.entry_fix.get()
            self.ExcelManager.dataSheet[self.casheCoordi] = int(fixData)
            self.result_label.destroy()
            self.result_label = tk.Label(self.top, text=f"이름: {self.name}\n 수정된 작정금액:  {self.ExcelManager.dataSheet[self.casheCoordi].value}")
            self.result_label.pack(pady=10)
        else:
            messagebox.showinfo("Error: There is no Cashe 해당하는 QR Data가 존재하지 않습니다!", "해당하는 Cache Data가 존재하지 않습니다!")
        self.entry_qr.focus()

        # 엔트리 Reset
        self.entry_fix.delete(0, 'end')

    def read_qr(self,event=None):
        # Entry에서 QR 코드를 가져오기
        paddingValue = 2
        qrData = self.entry_qr.get()
        Data = self.parsingQR(qrData)
        category = int(Data[0])
        checkEmpty = self.ExcelManager.findCoordinate_in_column(category, qrData, self.ExcelManager.dataSheet, 2300)
        if checkEmpty is None: 
            messagebox.showinfo("Error: There is no data", "{qrData} 해당하는 24년도 작정 Data가 존재하지 않습니다!")
            return
        row = checkEmpty.row        
        self.name = self.ExcelManager.dataSheet["G"+str(row)].value
        # QR Data 예외처리


        decidedCell = self.ExcelManager.findCoordinate(1, "2024년 작정금액", self.ExcelManager.dataSheet)
        categorizedCell = self.ExcelManager.nth_column_coordinate(decidedCell.coordinate, category - 1)
        cell = self.ExcelManager.nth_row_coordinate(categorizedCell, row - 1)
        decidedMoney = self.ExcelManager.dataSheet[cell].value
        print(decidedMoney)
        
        # Write decided Money
        if (decidedMoney != None) :
            # Get date
            date = get_current_sunday()
            dateCoordi = self.ExcelManager.findCoordinate(1, date, self.ExcelManager.dataSheet)
            if dateCoordi == None:
                messagebox.showinfo("Error: Date Cell does not exist", "해당하는 날짜 셀이 존재하지 않습니다!")
                return 
            dateCategorizedCoordi = self.ExcelManager.nth_column_coordinate(dateCoordi.coordinate, category - 1)
            dateCell = self.ExcelManager.nth_row_coordinate(dateCategorizedCoordi, row - 1)
            tempCheck = self.ExcelManager.dataSheet[dateCell] = decidedMoney
            
            # Make a casheCoordi
            self.casheCoordi = dateCell
            print("Check value cell data is " + str(tempCheck))
        else:
            messagebox.showinfo("Error: Empty Cell", "작정금액이 없습니다")
            return

        self.result_label.destroy()
        self.result_label = tk.Label(self.top, text=f"이름: {self.name}\n작정항목: {DataWindow.categoryArr[category - 1]}\n작정금액:  {decidedMoney} \n날짜: {date}")
        self.result_label.pack(pady=10)

        # 엔트리 Reset
        self.entry_qr.delete(0, 'end')

    def saveData(self):
        MyApp.ExcelManager.workbook.save(MyApp.ExcelManager.filepath)

    def parsingQR(self, input):
        parsed_values = input.split('-')

        return parsed_values