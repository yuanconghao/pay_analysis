import os
import re
import pdfplumber
import pandas as pd

def get_desktop_path():
    if os.name == 'posix':
        # 如果操作系统是类Unix（如Linux或macOS）
        return os.path.expanduser("~/Desktop")
    elif os.name == 'nt':
        # 如果操作系统是Windows
        return os.path.join(os.path.expanduser("~"), "Desktop")
    else:
        # 对于其他操作系统，您可以在这里添加适当的处理逻辑
        return "Unsupported OS"

desktop_path = get_desktop_path()

def getFileNames(dir_path):
    """
    get all legal pdf files
    :return:
    """
    filenames = []
    if not dir_path:
        dir_path = "./datas"
    files = os.listdir(dir_path)
    for file in files:
        match = re.search(r'\b\d{8}\b', file)
        if match:
            extracted_date = match.group()
            new_file_name = f'{extracted_date}.pdf'

            filenames.append(extracted_date)
        else:
            print('date info not found')
            new_file_name = None

        if new_file_name:
            # construct new path
            new_file_path = os.path.join(dir_path, new_file_name)
            file_path = dir_path + "/" + file
            # rename
            os.rename(file_path, new_file_path)

    return filenames


def getPdfTable(file):
    """
    parse single pdf file, and get table data
    :param file: 20230831
    :return:
    """
    # read pdf file
    #pdf = pdfplumber.open("./datas/" + file + ".pdf")
    pdf = pdfplumber.open(file)
    # get page
    first_page = pdf.pages[0]
    # auto read table info, return list
    table = first_page.extract_table()
    return table


def formatTable(tables):
    """
    format all tables data, output dict can be used
    :param tables:
    :return:
    """
    tables_dict = {}
    for file, table in tables.items():
        table = table[5:11]
        table_dict = {}
        instalment_period = [item[1] for item in table]
        txs = [item[3] for item in table]
        amt = [item[5] for item in table]
        rate = [item[10] for item in table]
        fixed = [item[12] for item in table]
        net_cost = [item[14] for item in table]
        table_dict['instalment_period'] = instalment_period
        table_dict['txs'] = [int(e) for e in txs]
        table_dict['amt'] = [float(e.replace(',', '')) for e in amt]
        table_dict['amt_s'] = amt
        table_dict['rate'] = [round(float(e.rstrip('%')) / 100, 4) for e in rate]
        table_dict['rate_s'] = rate
        table_dict['fixed'] = [float(e.rstrip('@')) for e in fixed]
        table_dict['fixed_s'] = fixed
        table_dict['net_cost'] = [float(e.replace(',', '')) for e in net_cost]
        table_dict['net_cost_s'] = net_cost
        tables_dict[file] = table_dict
    return tables_dict


def getPdfTables(path, files):
    """
    get all pdf tables, loop gain
    :param files:
    :return:
    """
    tables_dict = {}
    for file in files:
        file_path = path + "/" + file + '.pdf'
        tables_dict[file] = getPdfTable(file_path)
    return tables_dict


def save2Excel(dicts, name):
    """
    save parse dicts data to excel
    :param dicts:
    :param name:
    :return:
    """
    # create outputs dir if not exist
    # if not os.path.exists('outputs'):
    #     os.makedirs('outputs')
    #path = "~/Desktop/"
    path = desktop_path
    if os.name == 'nt':
        path += "\\"
    else:
        path += "/"
    # save to excel
    df = pd.DataFrame(dicts)
    # df.to_excel("./outputs/" + name + '.xlsx', index=False, auto_column_width=True)
    # create Pandas ExcelWriter Object
    writer = pd.ExcelWriter(path + name + '.xlsx', engine='xlsxwriter')

    # put DataFrame write to Excel file
    df.to_excel(writer, sheet_name='Sheet1', index=False)

    # get ExcelWriter Object attribute workbook and worksheet
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    # create Pandas Styler Object
    style = df.style
    # Align cell content to the right
    style.applymap(lambda x: 'text-align: right', subset=pd.IndexSlice[:, :])

    # write the formatted styler Object to excel file
    style.to_excel(writer, sheet_name='Sheet1', index=False, engine='openpyxl')

    # auto adjust column width
    for i, col in enumerate(df.columns):
        column_len = max(df[col].astype(str).str.len().max(), len(col))
        col_width = (column_len + 3) * 1.2  # 增加一些额外空间
        worksheet.set_column(i, i, col_width)

    # save writer
    writer.save()


def parseDatas(tables):
    first_key = next(iter(tables))
    first_e = tables[first_key]
    # 分期方案
    instalment_period = first_e['instalment_period']
    # 分期费率
    rate = first_e['rate']
    rate_s = first_e['rate_s']
    # 固定charge
    fixed = first_e['fixed']
    fixed_s = first_e['fixed']
    first_e_txs_len = len(first_e['txs'])
    order_num = [0] * first_e_txs_len

    c_amt_w = {}
    c_net_cost_w = {}
    for week, datas in tables.items():
        order_num = [x + y for x, y in zip(order_num, datas['txs'])]
        # amt_key = "w_" + week
        c_amt_w[week] = datas['amt']
        c_net_cost_w[week] = datas['net_cost']

    c_amt_w = {k: c_amt_w[k] for k in sorted(c_amt_w)}
    c_net_cost_w = {k: c_net_cost_w[k] for k in sorted(c_net_cost_w)}
    c_fixed_charge = [x * y for x, y in zip(order_num, fixed)]

    l = len(instalment_period)
    a_balance = [None] * l
    b_income = [None] * l
    d_balance = [None] * l

    dicts = {
        "A.月初余额": a_balance,
        "B.收入": b_income,
        "分期方案": instalment_period,
        "订单数": order_num,
        "分期费率": rate_s,
        "固定charge": fixed,
    }

    # 追加提现订单总额部分，几个pdf文件则dict长度大小为多少
    dicts.update(c_amt_w)

    # 创建一个字典用于存放相同位置值的和
    c_amt_w_sum = []
    c_net_cost_w_sum = []

    # 循环迭代每个键
    for key in c_amt_w:
        values = c_amt_w[key]

        # 如果sum_dict为空，直接将values赋值给sum_dict
        if not c_amt_w_sum:
            c_amt_w_sum = [v for k, v in zip(range(len(values)), values)]
        else:
            # 对相同位置的值进行累加
            for i, value in enumerate(values):
                c_amt_w_sum[i] += value

    # 使用循环迭代对应位置的元素相乘
    c_instalment_period_fee = []
    for i in range(len(c_amt_w_sum)):
        result = c_amt_w_sum[i] * rate[i]
        c_instalment_period_fee.append(round(result, 4))

    # net_cost 对应位置元素相加
    for key in c_net_cost_w:
        values = c_net_cost_w[key]

        # 如果c_net_cost_w_sum为空，直接赋值
        if not c_net_cost_w_sum:
            c_net_cost_w_sum = [v for k, v in zip(range(len(values)), values)]
        else:
            # 对应位置值相加
            for i, value in enumerate(values):
                c_net_cost_w_sum[i] += value

    print("\033[31m=========分期手续费求和===========\033[0m")
    print(round(sum(c_instalment_period_fee), 4))
    print("\033[31m=========Fixed Charge求和===========\033[0m")
    print(round(sum(c_fixed_charge), 4))
    print("\033[31m=========提现到账金额求和===========\033[0m")
    print(round(sum(c_net_cost_w_sum), 4))
    dicts_c = {
        "C.1:其中分期手续费": c_instalment_period_fee,
        "C.2:Fixed Charge": c_fixed_charge,
        "C.3:提现到账金额": c_net_cost_w_sum,
        "D.余额": d_balance
    }
    # 追加C列部分，该部分主要是计算获得
    dicts.update(dicts_c)

    return dicts


def parseSumData(dicts):
    c_instalment_period_fee = dicts["C.1:其中分期手续费"]
    c_fixed_charge = dicts["C.2:Fixed Charge"]
    c_net_cost_w_sum = dicts["C.3:提现到账金额"]
    return {
        "分期手续费求和": round(sum(c_instalment_period_fee), 4),
        "FixedCharge求和": round(sum(c_fixed_charge), 4),
        "提现到账金额求和": round(sum(c_net_cost_w_sum), 4)
    }


import os
import tkinter as tk
from tkinter import filedialog
import datetime

# 创建主窗口
root = tk.Tk()
root.title("PDF内容计算和导出到Excel")
root.geometry("600x400+10+10")
# 设置列宽度
root.grid_columnconfigure(0, minsize=100)
root.grid_columnconfigure(1, minsize=500)

# 设置默认目录为[下载]目录
#default_folder_path = os.path.expanduser("~/Downloads")
default_folder_path = desktop_path


# 更新日志信息的函数
def update_log(message, color = None):
    text_widget.config(state=tk.NORMAL)
    if color:
        text_widget.insert(tk.END, message + "\n", color)
    else:
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
    filenames = getFileNames(folder_path)
    if not filenames:
        update_log("=========datas目录下无pdf文件==========")
        print("=========datas目录下无pdf文件==========")
        return
    update_log("=========文件重命名完成==========")
    print("=========文件重命名完成==========")
    update_log(str(filenames))
    print(filenames)
    # 2. 获取所有文件表格
    tables = getPdfTables(folder_path, filenames)
    update_log("=========数据提取完成==========")
    print("=========数据提取完成==========")
    print(tables)
    # 3. 格式化表格数据
    tables = formatTable(tables)
    update_log("=========格式化数据完成==========")
    print("=========格式化数据完成==========")
    print(tables)
    update_log(str(tables))
    # 4. 解析数据
    dicts = parseDatas(tables)
    update_log("=========解析数据完成==========")
    print("=========解析数据完成==========")
    print(dicts)
    update_log(str(dicts))
    sum_datas = parseSumData(dicts)
    print("=========数据计算完成==========")
    update_log("=========数据计算完成==========", 'red')
    print(sum_datas)
    update_log(str(sum_datas), 'red')
    # 3. 保存到文件
    # 获取当前日期和时间
    current_datetime = datetime.datetime.now()
    # 格式化日期和时间为所需的字符串格式
    formatted_datetime = current_datetime.strftime("%Y%m%d%H%M%S")
    # 创建文件名
    xlsx_name = f"payments_{formatted_datetime}"
    save2Excel(dicts, xlsx_name)
    update_log("=========Excel保存到桌面完成==========", 'green')
    update_log(xlsx_name + ".xlsx", 'green')
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
text_widget = tk.Text(root, wrap="word", height=20)
text_widget.config(state=tk.DISABLED)
text_widget.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky='w')
text_widget.tag_configure('red', foreground='red')
text_widget.tag_configure('green', foreground='green')
text_widget.tag_configure('blue', foreground='blue')

# developer_label = tk.Label(root, text="©conghao")
# developer_label.grid(row=3, columnspan=1, column=0, padx=1, pady=1, sticky='we')
root.mainloop()
