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

        self.result_label = tk.Label(self.top, text="")
        self.result_label.pack(pady=10)

        btn_exit = tk.Button(self.top, text="Save Data", command=self.saveData)
        btn_exit.pack(pady=10)

    def close_window(self):
        self.top.destroy()

    def read_qr(self,event=None):
        # Entry에서 QR 코드를 가져오기
        paddingValue = 2
        qrData = self.entry_qr.get()

        # QR Data -> Category
        parsedData = self.parsingQR(qrData)
        category = int(parsedData[0])
        
        # QR Data -> Find person id
        findPersonidCoordi = self.ExcelManager.findCoordinate_in_column(category + 1, qrData, self.ExcelManager.QRsheet, 2300)
        
        # QR Data 예외처리
        if findPersonidCoordi is None: 
            messagebox.showinfo("Error: There is no data", "{qrData} 해당하는 QR Data가 존재하지 않습니다!")
            return
        
        personIdCoordi =  self.ExcelManager.nth_column_coordinate(findPersonidCoordi.coordinate, 6 - category)
        personId = self.ExcelManager.QRsheet[personIdCoordi].value

        print(personId)
        # find Name
        nameCoordi =  self.ExcelManager.nth_column_coordinate(findPersonidCoordi.coordinate, 7 - category)
        name = self.ExcelManager.QRsheet[nameCoordi].value
        print(name)
        # Categorized 작정헌금 금액
        categorizedCoordi = self.ExcelManager.nth_column_coordinate(findPersonidCoordi.coordinate, 16)
        decidedMoney = self.ExcelManager.QRsheet[categorizedCoordi].value

        date = None
        print(decidedMoney)
        
        # Write decided Money
        if (decidedMoney != None) :
            # Get date
            date = get_current_sunday()
            dateCoordi = self.ExcelManager.findCoordinate(1, date, self.ExcelManager.dataSheet)
            if dateCoordi == None:
                messagebox.showinfo("Error: Date Cell does not exist", "해당하는 날짜 셀이 존재하지 않습니다!")
                return 
            print("date coordi" + dateCoordi.coordinate)
            dateCategorizedCoordi = self.ExcelManager.nth_column_coordinate(dateCoordi.coordinate, category - 1)
            print("categorized " + dateCategorizedCoordi)

            preDataIdCoordi = self.ExcelManager.findCoordinate_in_column(7, personId, self.ExcelManager.dataSheet, 2300)
            dataRowCoordi = int(self.ExcelManager.extract_numeric_part(preDataIdCoordi.coordinate))        
            dateIdCoordi = self.ExcelManager.nth_row_coordinate(dateCategorizedCoordi, dataRowCoordi  - 1)
            print("rowcategorized " + dateIdCoordi)
            tempCheck = self.ExcelManager.dataSheet[dateIdCoordi] = decidedMoney
            print("Check value cell data is " + str(tempCheck))
        else:
            messagebox.showinfo("Error: Empty Cell", "작정금액이 없습니다")

        self.result_label.destroy()
        self.result_label = tk.Label(self.top, text=f"이름: {name}\n작정항목: {DataWindow.categoryArr[category - 1]}\n작정금액:  {decidedMoney} \n날짜: {date}")
        self.result_label.pack(pady=10)

        # 엔트리 Reset
        self.entry_qr.delete(0, 'end')

    def saveData(self):
        MyApp.ExcelManager.workbook.save(MyApp.ExcelManager.filepath)

    def parsingQR(self, input):
        parsed_values = input.split('-')

        return parsed_values