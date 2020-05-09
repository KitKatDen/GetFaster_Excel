import pandas as pd
import openpyxl

table = pd.read_html("sale_order.xls")
table[0].to_excel("data_refresh.xlsx", index=False, header=False)
excelFile = openpyxl.load_workbook("data_refresh.xlsx")
sheet1 = excelFile.active
col_letter = ("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L")


# поиск по верхней строчки столбца по содержанию текста, после замена в этом столбце одного символа на другой


def s_and_r_in_col(col_name, what_change_to, what_change_for):
    if sheet1[col_letter[col] + "1"].value == col_name:
        for row in range(2, sheet1.max_row + 1):
            sheet1[col_letter[col] + str(row)] = sheet1[col_letter[col] + str(row)].value.replace(what_change_to,
                                                                                                  what_change_for)


# поиск по верхней строчки столбца по содержанию текста, после разделение на два блока информации из ячейки на эту
# ячейку и ячейку справа


def s_and_split_in_2_cols(col_name_0, split_symbol):
    if sheet1[col_letter[col] + "1"].value == col_name_0:
        for row in range(2, sheet1.max_row + 1):
            date_order = sheet1[col_letter[col] + str(row)].value.split().pop().split(split_symbol)
            sheet1[col_letter[col] + str(row)] = date_order[0]
            sheet1[col_letter[col + 1] + str(row)] = date_order[1]


for col in range(0, sheet1.max_column + 1):
    s_and_r_in_col("ID", "№", "")
    s_and_r_in_col("Сумма", ".00 руб.", "")
    s_and_split_in_2_cols("Дата и время доставки", "-")
excelFile.save("data.xlsx")
