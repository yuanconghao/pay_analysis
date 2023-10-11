import os
import re
import pdfplumber
import pandas as pd
from styleframe import StyleFrame


# columns = ["A.月初余额", "B.收入", "分期方案", "订单数", "分期费率", "固定charge",
#            "C.提现订单总额w1", "C.提现订单总额w2", "C.提现订单总额w3",
#            "C.提现订单总额w4", "C.提现订单总额w5", "C.1:其中分期手续费",
#            "C.2:Fixed Charge", "C.3:提现到账金额", "D.余额"]

def getFileNames():
    """
    get all legal pdf files
    :return:
    """
    filenames = []
    dir_path = "./datas/"
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
            file_path = dir_path + file
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
    pdf = pdfplumber.open("./datas/" + file + ".pdf")
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
        table_dict['instalment_period'] = instalment_period
        table_dict['txs'] = [int(e) for e in txs]
        table_dict['amt'] = [float(e.replace(',', '')) for e in amt]
        table_dict['amt_s'] = amt
        table_dict['rate'] = [round(float(e.rstrip('%')) / 100, 4) for e in rate]
        table_dict['rate_s'] = rate
        table_dict['fixed'] = [float(e.rstrip('@')) for e in fixed]
        table_dict['fixed_s'] = fixed
        tables_dict[file] = table_dict
    return tables_dict


def getPdfTables(files):
    """
    get all pdf tables, loop gain
    :param files:
    :return:
    """
    tables_dict = {}
    for file in files:
        tables_dict[file] = getPdfTable(file)
    return tables_dict


def save2Excel(dicts, name):
    """
    save parse dicts data to excel
    :param dicts:
    :param name:
    :return:
    """
    # create outputs dir if not exist
    if not os.path.exists('outputs'):
        os.makedirs('outputs')
    # save to excel
    df = pd.DataFrame(dicts)
    # df.to_excel("./outputs/" + name + '.xlsx', index=False, auto_column_width=True)
    # create Pandas ExcelWriter Object
    writer = pd.ExcelWriter("./outputs/" + name + '.xlsx', engine='xlsxwriter')

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
    for week, datas in tables.items():
        order_num = [x + y for x, y in zip(order_num, datas['txs'])]
        # amt_key = "w_" + week
        c_amt_w[week] = datas['amt']

    c_amt_w = {k: c_amt_w[k] for k in sorted(c_amt_w)}
    c_fixed_charge = [x * y for x, y in zip(order_num, fixed)]

    l = len(instalment_period)
    a_balance = [None] * l
    b_income = [None] * l
    c_amount = [None] * l
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

    print("\033[31m=========分期手续费求和===========\033[0m")
    print(round(sum(c_instalment_period_fee), 4))
    print("\033[31m=========Fixed Charge求和===========\033[0m")
    print(round(sum(c_fixed_charge), 4))
    dicts_c = {
        "C.1:其中分期手续费": c_instalment_period_fee,
        "C.2:Fixed Charge": c_fixed_charge,
        "C.3:提现到账金额": c_amount,
        "D.余额": d_balance
    }
    # 追加C列部分，该部分主要是计算获得
    dicts.update(dicts_c)

    return dicts


if __name__ == '__main__':
    # 1. 获取所有pdf文件并重命名 "VIS eStatement 20230907 - 88634961.pdf ==> 20230907.pdf"
    filenames = getFileNames()
    if not filenames:
        print("=========datas目录下无pdf文件==========")
        exit()
    print("=========文件重命名完成==========")
    print(filenames)
    # 2. 获取所有文件表格
    tables = getPdfTables(filenames)
    print("=========数据提取完成==========")
    print(tables)
    # 3. 格式化表格数据
    tables = formatTable(tables)
    print("=========格式化数据完成==========")
    print(tables)
    # 4. 解析数据
    dicts = parseDatas(tables)
    print("=========解析数据完成==========")
    print(dicts)
    # 3. 保存到文件
    save2Excel(dicts, 'output')
    print("=========保存到文件完成==========")
