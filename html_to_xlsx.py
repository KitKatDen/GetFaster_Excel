import pandas as pd
import openpyxl

table = pd.read_html('sale_order.xls')
table[0].to_excel('data.xlsx', index=False, header=False)
excelFile = openpyxl.load_workbook('data.xlsx')
sheet1 = excelFile.active

for row in range (2, sheet1.max_row+1):
 sheet1['B' + str(row)] = sheet1['B' + str(row)].value.replace('№', '')
 sheet1['C' + str(row)] = sheet1['C' + str(row)].value.replace('.00 руб.', '')
 date_order = sheet1['F' + str(row)].value.split().pop().split('-')
 sheet1['F' + str(row)] = date_order[0]
 sheet1['G' + str(row)] = date_order[1]
excelFile.save('data.xlsx')