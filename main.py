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

predmeti_p1 = ['']
for i in range(len(data[2])):
    if data[2][i] == 'Аттестационный период 1':
        predmeti_p1.append(data[0][i])

# Трассировка вывода предметов
# print(predmeti_p1)

# создание книги для проверенной выгрузки

out_book = openpyxl.Workbook()
out_book.remove(out_book.active)
out_sheet = out_book.create_sheet("Проверка ГОД")

# Обработка таблицы с периодами
def markToInt(x):
    marks = ['2','3','4','5']
    if x in marks:
        return int(x)
    return x

for fio in range(3,len(data)):
    p1 = ['Аттестационный период 1']
    p2 = ['Аттестационный период 2']
    if periods == 3:
        p3 = ['Аттестационный период 3']
    god = ['Год']
    predmeti_p1 = ['']
    
    for i in range(len(data[2])):
        if data[2][i] == 'Аттестационный период 1':
            predmeti_p1.append(data[0][i])
            p1.append(markToInt(data[fio][i]))
        if data[2][i] == 'Аттестационный период 2':
            p2.append(markToInt(data[fio][i]))
        if data[2][i] == 'Аттестационный период 3' and periods == 3:
            p3.append(markToInt(data[fio][i]))
        if data[2][i] == 'Год' and data[0][i] != 'Математика':
            god.append(markToInt(data[fio][i]))
    
    # Трассировка вывода данных
    # print(data[fio][0])
    # print(p1)
    # print(p2)
    # if periods == 3:
    #     print(p3)
    # print(god)

    out_sheet.append([data[fio][0]])

    out_sheet.append(predmeti_p1)
    out_sheet.append(p1)
    out_sheet.append(p2)
    if periods == 3:
        out_sheet.append(p3)
    out_sheet.append(god)

out_book.save('test.xlsx')
