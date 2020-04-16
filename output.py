import openpyxl

excelFile = openpyxl.load_workbook('Доставка-30.04.2020_изм.xlsx')
sheets = excelFile.get_sheet_names()
col_letter = ['No Letter','A','B','C','D','E','F','G','H','I','J','K','L']
total_cash = 0
total_cc = 0

for i in range(len(sheets)):
  nq_orders = False
  money_from_clients = False
  active_sheet = excelFile.get_sheet_by_name(sheets[i])
  canceled_orders = 0
  not_taken_IQOS = 0
  cash = 0
  cc = 0
  for n in range (1, active_sheet.max_column+1):
   for n1 in range (1, active_sheet.max_row+1):
    if str(active_sheet[col_letter[n] + str(n1)].value).find('Всего заказов') == 0:
     n_orders = active_sheet[col_letter[n+1] + str(n1)].value #количество заказов у курьера
     nq_orders = True #проверка наличия пункта "Всего заказов"
     cancl_orders = False #проверка наличия пункта "Отменен"
    if str(active_sheet[col_letter[n] + str(n1)].value).find('Отменен') == 0 and cancl_orders == False:
     canceled_orders = active_sheet[col_letter[n+1] + str(n1)].value #количество отменённых заказов у курьера
     cancl_orders = True
    if str(active_sheet[col_letter[n] + str(n1)].value).find('Отменен') == 0 and active_sheet[col_letter[n+1] + str(n1)].value != canceled_orders and active_sheet[col_letter[n+1] + str(n1)].value == '-':
     not_taken_IQOS = not_taken_IQOS + 1
   if active_sheet[col_letter[n]+ '2'].value == 'Получено с клиента':
    money_from_clients = True #проверка наличия пункта "Получено с клиента"
    cash = 0 #сумма налички
    q_cash = 0 #кол-во заказов оплаченных наличкой
    cc = 0 #сумма карты
    q_cc = 0 #кол-во заказов оплаченных картой
    for p in range (1, n_orders+1):
     if_cash = active_sheet[col_letter[n+2]+ str(2+p)].value
     if_cc = active_sheet[col_letter[n+2]+ str(2+p)].value
     if if_cash.find('Нал') == 0 or if_cash.find('Нал') == 1 or if_cash.find('нал') == 0 or if_cash.find('нал') == 1:
      cash = cash + active_sheet[col_letter[n] + str(2+p)].value
      q_cash = q_cash + 1
     if if_cc.find('Карта') == 0 or if_cc.find('Карта') == 1 or if_cc.find('карта') == 0 or if_cc.find('карта') == 1:
      cc = cc + active_sheet[col_letter[n] + str(2+p)].value
      q_cc = q_cc + 1
  if nq_orders == False:
   print('Ошибка, не найдено количество заказов')
  if money_from_clients == False:
   print('Ошибка, не найдено ячейки "Получено с клиента"')
  if q_cc <= 4 and q_cc != 1:
   print(sheets[i] + ': ' + str(cash) + ' наличными, ' + str(q_cc) + ' чека, ' + str(canceled_orders - not_taken_IQOS) +' IQOS')
  elif q_cc == 1:
   print(sheets[i] + ': ' + str(cash) + ' наличными, ' + str(q_cc) + ' чек, ' + str(canceled_orders - not_taken_IQOS) +' IQOS')
  else:
   print(sheets[i] + ': ' + str(cash) + ' наличными, ' + str(q_cc) + ' чеков, ' + str(canceled_orders - not_taken_IQOS) +' IQOS')
  total_cash = total_cash + cash
  total_cc = total_cc + cc
print(str(total_cash) + ' наличными, ' + str(total_cc) + ' по карте')