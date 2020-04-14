import pandas as pd
import openpyxl

table = pd.read_html('sale_order.xls')
table[0].to_excel('data.xlsx', index=False, header=False)
excelFile = openpyxl.load_workbook('data.xlsx')
sheet1 = excelFile.active
print(sheet1.max_row)

for row in range (2, sheet1.max_row+1):
 #print(sheet1['B' + str(row+1)].value) - проверка значение ячейки
 sheet1['B' + str(row)] = sheet1['B' + str(row)].value.replace('№', '')
 #print(sheet1['B' + str(row)].value) - проверка ИЗМЕНЁННОГО значение ячейки
 #print(sheet1['C' + str(row+1)].value)

 sheet1['C' + str(row)] = sheet1['C' + str(row)].value.replace('.00 руб.', '')
 #print(sheet1['C' + str(row)].value)

 row_date_order = 2
 date_order = sheet1['F' + str(row)].value.split().pop().split('-')
 #print(type(sheet1['F' + str(row)].value))
 #print(row)
 sheet1['F' + str(row)] = date_order[0]
 sheet1['G' + str(row)] = date_order[1]
excelFile.save('data.xlsx')

