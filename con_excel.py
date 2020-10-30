import os, xlrd, xlwt

conExcel = xlwt.Workbook()
sheet = conExcel.add_sheet(u'合并', cell_overwrite_ok=True)
title = ['课文', '总字数','用字数', '平均字频', '总词数（有重复）', '总词数（无重复）', '平均词频', '平均句长-1', '平均句长-2', '最长句长']
for col in range(len(title)):
    sheet.write(0,col,title[col])

filePath = './result/'
excelList = os.listdir(filePath)

for row, fileName in enumerate(excelList):
    excel = xlrd.open_workbook(filePath+fileName)
    excel.sheet_names()
    table = excel.sheet_by_name('统计')
    data = table.col_values(1)

    for col in range(len(data)):
        sheet.write(row+1,0,fileName)
        sheet.write(row+1,col+1,data[col])

conExcel.save('conexcel.xls')
print('合并完成')