import xlrd
import os


#os.chdir('D:\!!!_KitKatDen_Yandex.Disk\YandexDisk\GetFaster\апрель\9.04')
xls_book = xlrd.open_workbook('sale_order.xls')

xls_sheet = xls_book.sheet_by_index(0)

for row_num in range(xls_sheet.nrows):
    print(xls_sheet.row_values(row_num))

