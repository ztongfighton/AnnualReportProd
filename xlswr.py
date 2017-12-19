import xlrd
import xlwt
import pandas as pd

######################################################################
#将dict写入xls文件保存到本地
######################################################################
def writeDict2Xls(d, header, file_path):
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet("Sheet1")
    row = 0
    col = 0
    for item in header:
        worksheet.write(row, col, item)
        col += 1
    row += 1

    for key in d.keys():
        col = 0
        worksheet.write(row, col, key)
        for item in d[key]:
            col += 1
            worksheet.write(row, col, item)
        row += 1
    workbook.save(file_path)

######################################################################
#将xls文件读入dict
######################################################################
def readXls2Dict(file_path, sheet_index):
    workbook = xlrd.open_workbook(file_path)
    worksheet = workbook.sheet_by_index(sheet_index)
    nrows = worksheet.nrows
    ncols = worksheet.ncols
    d = dict()
    for i in range(1, nrows):
        key = worksheet.cell(i, 0).value
        item = []
        for j in range(1, ncols):
            item.append(worksheet.cell(i, j).value)
        d[key] = item
    return d

######################################################################
#将xls文件读入list
######################################################################
def readXls2List(file_path, col):
    data = pd.read_excel(file_path)
    l = data[col].tolist()
    return l

######################################################################
#将list写入xls文件保存到本地
######################################################################
def writeList2Xls(l, header, file_path):
    writer = pd.ExcelWriter(file_path)
    transaction = pd.DataFrame(l, columns=header)
    transaction.to_excel(writer, "Sheet1", index=False)
    writer.save()











