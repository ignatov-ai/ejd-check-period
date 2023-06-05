##############################################################################################
### Программа для облегчения проверки из выгрузок ЭЖД отметок по средним баллам и периодам ###
##############################################################################################

import pandas as pd
import openpyxl

### Чтение периодов из выгрузки ###

# Укажите кол-во периодов для выгрузки
periods = 2

# Открытие файлов с периодами из выгрузки
url_period_1 = pd.ExcelFile('data/10Я-1.xlsx')
url_period_2 = pd.ExcelFile('data/10Я-2.xlsx')

if periods == 3:
    url_period_3 = pd.ExcelFile('data/10Я-3.xlsx')

# Открытие файла с итоговыми оценками за периоды
url_all_periods = pd.ExcelFile('data/10Я-все.xlsx')


book = openpyxl.load_workbook(filename='data/10Я-все.xlsx')
sheet = book.worksheets[0]

# создаем пустой список для хранения данных из таблицы
data = []

# проходимся по всем строкам таблицы
for row in sheet.iter_rows(values_only=True):
    # создаем пустой список для хранения данных из текущей строки
    row_data = []
    # проходимся по всем ячейкам текущей строки
    for cell in row:
        # добавляем значение ячейки в список данных текущей строки
        if cell != None:
            row_data.append(cell)
        else:
            row_data.append('')
    # добавляем список данных текущей строки в список данных всей таблицы
    data.append(row_data)

for i in range (1,len(data[0])):
    if data[0][i] == '':
        data[0][i] = data[0][i-1]

predmeti = ';'.join(data[0])

p1 = data[3][0]+';'
p2 = data[3][0]+';'
p3 = data[3][0]+';'
god = data[3][0]+';'
predmeti_p1 = ''
predmeti_p2 = ''
predmeti_p3 = ''
predmeti_god = ''

for i in range(len(data[2])):
    if data[2][i] == 'Аттестационный период 1':
        predmeti_p1 += data[0][i] + ';'
        p1 += data[3][i] + ';'
    if data[2][i] == 'Аттестационный период 2':
        predmeti_p2 += data[0][i] + ';'
        p2 += data[3][i] + ';'
    if data[2][i] == 'Аттестационный период 3':
        predmeti_p3 += data[0][i] + ';'
        p3 += data[3][i] + ';'
    if data[2][i] == 'Год':
        predmeti_god += data[0][i] + ';'
        god += data[3][i] + ';'

print(predmeti_p1)
print(p1)
print(predmeti_p2)
print(p2)
print(predmeti_p3)
print(p3)
print(predmeti_god)
print(god)
