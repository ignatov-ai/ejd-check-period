##############################################################################################
### Программа для облегчения проверки из выгрузок ЭЖД отметок по средним баллам и периодам ###
##############################################################################################

import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, PatternFill, Font, Fill, NamedStyle, Border
from math import ceil

### Чтение периодов из выгрузки ###

# Укажите кол-во периодов для выгрузки
periods = 2

# Открытие файлов с периодами из выгрузки
url_period_1 = pd.ExcelFile('data/10Я-1.xlsx')
url_period_2 = pd.ExcelFile('data/10Я-2.xlsx')

if periods == 3:
    url_period_3 = pd.ExcelFile('data/10Я-3.xlsx')
url_all_periods = pd.ExcelFile('data/10Я-все.xlsx')


# Открытие файла с итоговыми оценками за периоды
file_name_all_periods = '10Ч-Костина-все'

book = openpyxl.load_workbook(filename='data/' + file_name_all_periods + '.xlsx')
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
    if data[2][i] == 'Год':
        predmeti_p1.append(data[0][i])

# Трассировка вывода предметов
# print(predmeti_p1)

# создание книги для проверенной выгрузки
out_book = openpyxl.Workbook()
out_book.remove(out_book.active)
out_sheet = out_book.create_sheet("Проверка ГОД")
out_sheet = out_book.active
out_sheet.column_dimensions['A'].width = 30

for col in 'BCDEFGHIJKLMNOPQRSTUVWXWZ':
    out_sheet.column_dimensions[col].width = 4


# Обработка таблицы с периодами
def markToInt(x):
    marks = ['2','3','4','5']
    print(x)
    if x in marks:
        return int(x)
    elif x == '':
        return ''
    else:
        return x

# Количество выявленных ошибок
errors_count = 0

for fio in range(3,len(data)):
    p1 = ['Аттестационный период 1']
    p2 = ['Аттестационный период 2']
    if periods == 3:
        p3 = ['Аттестационный период 3']
    god = ['Год']
    predmeti_p1 = ['']

    ### НУЖНО ДОБАВИТЬ ОБРАБОЧИК ВСТАВКИ ПУСТЫХ ПЕРИОДОВ ###
    
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
    print(data[fio][0])
    print(p1)
    print(p2)
    if periods == 3:
        print(p3)
    print(god)

    out_sheet.append([data[fio][0]])

    out_sheet.append(predmeti_p1)
    out_sheet.append(p1)
    out_sheet.append(p2)
    if periods == 3:
        out_sheet.append(p3)
    out_sheet.append(god)

    # проверка выставления итоговых оценок
    cols = 'ABCDEFGHIJKLMNOPQRSTUVWXWZ'
    for i in range(len(predmeti_p1)-1):
        if god[i+1] != '':
            if periods != 3:
                a = 0
                b = 0
                if p1[i+1] == '':
                    a == 0
                elif p1[i+1] == 'НПА' or p1[i+1] == 'АЗ':
                    continue
                else:
                    a = p1[i+1]

                if p2[i+1] == '':
                    b == 0
                elif p2[i+1] == 'НПА' or p2[i+1] == 'АЗ':
                    continue
                else:
                    b = p2[i+1]

                if a == 0 or b == 0:
                    avg = a + b
                else:
                    avg = ceil((a + b) / 2)

                print('Предмет:',out_sheet[cols[i + 1] + str(2 + (fio - 3) * 5)].value,'| ячейка:', cols[i + 1] + str(5 + (fio - 3) * 5), 'a=',a,'b=',b, '(a+b)/ 2=',(a+b)/2,'avg=',avg,'god=',god[i+1])

                if avg == god[i+1]:
                    out_sheet[cols[i+1]+str(5+(fio-3)*5)].fill = PatternFill('solid',fgColor='A6F16C')
                else:
                    # Закрашиваем ячейку с ошибкой красным
                    out_sheet[cols[i + 1] + str(5 + (fio - 3) * 5)].fill = PatternFill('solid', fgColor='FF8B73')
                    errors_count += 1
                    # Закрашиваем ФИО ученика с ошибкой красным
                    out_sheet['A' + str(1+(fio - 3) * 5)].fill = PatternFill('solid', fgColor='FF8B73')

    # переворот названий предметов в ячейке на 90 градусов
    for row in out_sheet['B'+str(2+(fio-3)*5)+':Z'+str(2+(fio-3)*5)]:
        for cell in row:
            cell.alignment = Alignment(textRotation=90)

out_sheet.insert_rows(1)
out_sheet['A1'] = "Количество ошибок:"
out_sheet['B1'] = errors_count
if errors_count == 0:
    out_sheet['A1'].fill = PatternFill('solid', fgColor='A6F16C')
    out_sheet['B1'].fill = PatternFill('solid', fgColor='A6F16C')
else:
    out_sheet['A1'].fill = PatternFill('solid', fgColor='FF8B73')
    out_sheet['B1'].fill = PatternFill('solid', fgColor='FF8B73')

out_book.save('checked/' + file_name_all_periods + '-ПРОВЕРЕНО('+str(errors_count)+').xlsx')