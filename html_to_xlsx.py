import pandas as pd
import openpyxl

table = pd.read_html('sale_order.xls')
table[0].to_excel('data.xlsx', index=False, header=False)
excelFile = openpyxl.load_workbook('data.xlsx')
sheet1 = excelFile.active
col_letter = ['No Letter','A','B','C','D','E','F','G','H','I','J','K','L']

for col in range (1,sheet1.max_column+1):
 if sheet1[col_letter[col]+ '1'].value == 'ID':
  for row in range(2, sheet1.max_row + 1):
   sheet1[col_letter[col] + str(row)] = sheet1['B' + str(row)].value.replace('№', '')
 if sheet1[col_letter[col]+ '1'].value == 'Сумма':
  for row in range(2, sheet1.max_row + 1):
   sheet1['C' + str(row)] = sheet1['C' + str(row)].value.replace('.00 руб.', '')
 if sheet1[col_letter[col]+ '1'].value == 'Дата и время доставки':
  for row in range(2, sheet1.max_row + 1):
   date_order = sheet1[col_letter[col] + str(row)].value.split().pop().split('-')
   sheet1[col_letter[col] + str(row)] = date_order[0]
   sheet1[col_letter[col+1] + str(row)] = date_order[1]
excelFile.save('data.xlsx')