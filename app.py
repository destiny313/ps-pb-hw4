import pandas as p
import collections as c

import openpyxl as opyx

excel_data = p.read_excel('logs.xlsx', sheet_name='log')
excel_data_dict = excel_data.to_dict(orient='records')

browsers = c.Counter()
items = c.Counter()
male_like = c.Counter()
female_like = c.Counter()

#Вычисляем списки браузеров и товаров
for session in excel_data_dict:
    browsers[session['Браузер']] += 1
    for item in session['Купленные товары'].split(','):
        items[item] += 1
        if session['Пол'] == 'м':
            male_like[item] += 1
        else:
            female_like[item] += 1

#Словарь для привязки сессий по браузерам к месяцам
browsers_date = {browsers.most_common(7)[a][0] : {b : 0 for b in range(1,13)} for a in range(len(browsers.most_common(7)))}


#Словарь для привязки сессий по товарам к месяцам
items_date = {items.most_common(7)[a][0] : {b : 0 for b in range(1,13)} for a in range(len(items.most_common(7)))}


#Открываем файл отчёта дл записи
excel_report = opyx.load_workbook(filename='report.xlsx')
sheet = excel_report['Лист1'];


#список столбцов для месяцев
month_count = [a for a in 'cdefghijklmn']


#Заполняем данные по браузерам
a_count = 5
for browser_name in browsers.most_common(7):
    for session in excel_data_dict:
        if session['Браузер'] == browser_name[0]:
            browsers_date[session['Браузер']][session['Дата посещения'].month] += 1
    sheet["A" + str(a_count)] = browser_name[0]

    for i in range(0,12):
        sheet[month_count[i] + str(a_count)] = browsers_date[browser_name[0]][i+1]

    a_count += 1

#Заполняем данные по товарам
a_count = 19
for item_name in items.most_common(7):
    for session in excel_data_dict:
        for item in session['Купленные товары'].split(','):
            if item == item_name[0]:
                items_date[item][session['Дата посещения'].month] += 1

    sheet["A" + str(a_count)] = item_name[0]

    for i in range(0,12):
        sheet[month_count[i] + str(a_count)] = items_date[item_name[0]][i+1]

    a_count += 1


#Заполняем данные по предпочтениям
sheet["B31"] = male_like.most_common(1)[0][0]
sheet["B32"] = female_like.most_common(1)[0][0]
sheet["B33"] = male_like.most_common()[:-2:-1][0][0]
sheet["B34"] = female_like.most_common()[:-2:-1][0][0]

#Закрываем и сохраняем файл отчета
excel_report.save('report.xlsx')
