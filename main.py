from UILayout import MyApp
import tkinter as tk
from readxlsx import ExcelManager
from tkinter import messagebox
import asyncio

async def save_and_close():
    # save 작업 수행
    MyApp.ExcelManager.workbook.save(MyApp.ExcelManager.filepath)
    messagebox.showinfo("Success Save and Close", "성공적으로 데이터를 저장했습니다!")
    print("Success Sva and Close!")
    # close 작업 수행
    MyApp.ExcelManager.close()
    app.root.destroy()



def on_closing():
    if app.ExcelManager == 0:
        app.root.destroy()
    else:
        asyncio.run(save_and_close())


# 애플리케이션 실행
root = tk.Tk()
app = MyApp(root)
root.iconbitmap("icon.ico")

# Auto Save
root.protocol("WM_DELETE_WINDOW", on_closing)

root.mainloop()