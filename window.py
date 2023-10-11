import os
import tkinter as tk
from tkinter import filedialog
import PyPDF2
import openpyxl as px
import tk_payments
import datetime

# 创建主窗口
root = tk.Tk()
root.title("PDF内容计算和导出到Excel")
root.geometry("600x300+10+10")
# 设置列宽度
root.grid_columnconfigure(0, minsize=100)
root.grid_columnconfigure(1, minsize=500)

# 设置默认目录为[下载]目录
default_folder_path = os.path.expanduser("~/Downloads")

# 更新日志信息的函数
def update_log(message):
    text_widget.config(state=tk.NORMAL)
    text_widget.insert(tk.END, message + "\n")
    text_widget.config(state=tk.DISABLED)
    text_widget.see(tk.END)

# 选择目录
def select_directory():
    global folder_path
    # folder_path = filedialog.askdirectory(initialdir=default_folder_path)
    folder_path = filedialog.askdirectory(initialdir=default_folder_path)
    print(folder_path)
    folder_path_label.config(text=folder_path)


# 计算PDF文件内容并导出到Excel
def process_pdf_to_excel():
    if not folder_path:
        return

    # 1. 获取所有pdf文件并重命名 "VIS eStatement 20230907 - 88634961.pdf ==> 20230907.pdf"
    filenames = tk_payments.getFileNames(folder_path)
    if not filenames:
        update_log("=========datas目录下无pdf文件==========")
        print("=========datas目录下无pdf文件==========")
        return
    update_log("=========文件重命名完成==========")
    print("=========文件重命名完成==========")
    update_log(str(filenames))
    print(filenames)
    # 2. 获取所有文件表格
    tables = tk_payments.getPdfTables(filenames)
    update_log("=========数据提取完成==========")
    print("=========数据提取完成==========")
    print(tables)
    # 3. 格式化表格数据
    tables = tk_payments.formatTable(tables)
    update_log("=========格式化数据完成==========")
    print("=========格式化数据完成==========")
    print(tables)
    # 4. 解析数据
    dicts = tk_payments.parseDatas(tables)
    update_log("=========解析数据完成==========")
    print("=========解析数据完成==========")
    print(dicts)
    sum_datas = tk_payments.parseSumData(dicts)
    update_log(str(sum_datas))
    # 3. 保存到文件
    # 获取当前日期和时间
    current_datetime = datetime.datetime.now()
    # 格式化日期和时间为所需的字符串格式
    formatted_datetime = current_datetime.strftime("%Y%m%d%H%M%S")
    # 创建文件名
    xlsx_name = f"payments_{formatted_datetime}"
    tk_payments.save2Excel(dicts, xlsx_name)
    update_log("=========Excel保存到桌面完成==========")
    update_log(xlsx_name + ".xlsx")
    # 保存Excel文件
    # excel_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel文件", "*.xlsx")])
    # if excel_path:
    #     excel_file.save(excel_path)
    print("=========Excel保存到桌面完成==========")


# 创建选择目录按钮
select_button = tk.Button(root, text="选择PDF文件目录", command=select_directory)
select_button.grid(row=0, column=0, padx=10, pady=10, sticky='w')

# 显示选择的目录路径
folder_path_label = tk.Label(root, text=":")
folder_path_label.grid(row=0, column=1, padx=10, pady=10, sticky='w')

# 创建处理按钮
process_button = tk.Button(root, text="分析计算并导出Excel", command=process_pdf_to_excel)
process_button.grid(row=1, column=0, padx=10, pady=10, sticky='w')

# 创建一个Text控件
text_widget = tk.Text(root, wrap="word", height=10)
text_widget.config(state=tk.DISABLED)
text_widget.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky='w')


root.mainloop()
