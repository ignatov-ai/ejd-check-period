import itertools
import os
from itertools import permutations, product

import openpyxl
from openpyxl.styles import Alignment, PatternFill, Font, Fill, NamedStyle, Border
from math import ceil

file_name_all_periods = '9ч-расширенный.xlsx'

book = openpyxl.load_workbook(filename='data/' + file_name_all_periods)

sheet_names = book.sheetnames

for s in range(len(sheet_names)):
    sheet = book.worksheets[s]

    # Проверяем, является ли ячейка A1 объединенной
    merged_ranges = sheet.merged_cells.ranges.copy()
    for merged_range in merged_ranges:
        sheet.unmerge_cells(merged_range.coord)

    for i in range(21, 24):
        for j in range(1,40,2):
            sheet.unmerge_cells(start_row=j, start_column=i, end_row=j+1, end_column=i)

    sheet.delete_cols(20, 5)

    i = 0

    # Создаем список букв
    letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

    # Создаем комбинации из 1 и 2 символов
    col = itertools.chain(letters, itertools.product(letters, repeat=2))

    i = 0
    for s in col:
        t = ''.join(s)
        sheet.column_dimensions[t].width = 3
        if i == 400:
            break
        i += 1
    sheet.column_dimensions['A'].width = 4
    sheet.column_dimensions['B'].width = 30

    # Определяем количество заполненных строк в столбце A
    num_rows = len(list(sheet['A']))
    data_pages_count = num_rows // 50

    for i in range(1,data_pages_count):
        # Переносим область
        sheet.move_range('C' + str(1 + i*50) + ':S' + str((i+1)*50), rows=-(i*50), cols=(17 + (i-1)*17))

    # Удаляем перемещенные строки до конца таблицы
    sheet.delete_rows(41, sheet.max_row)

    # Меняем тип у оценок из str в int
    col = itertools.chain(letters, itertools.product(letters, repeat=2))
    for c in col:
        stop = False
        for row in range(2,40):
            cell = ''.join(c)+str(row)
            #print(cell,sheet[cell].value)
            if sheet[cell].value in ['1','2', '3', '4', '5']:
                sheet[cell].value = int(sheet[cell].value)
            elif sheet[cell].value == None:
                sheet[cell].value = sheet[cell].value or ''

            if cell[:2] == 'OK':
                stop = True
                break
        if stop:
            break

    for i in range(4,sheet.max_row):
        sum_period = 0
        count = 0
        for j in range(3,sheet.max_column-1):
            mark = sheet.cell(row=i, column=j).value
            koef = sheet.cell(row=i, column=j + 1).value

            if sheet.cell(row=3, column=j).value == 'оц' and (mark and koef) != '':
                print('оц =',mark,'коэф =',koef,'итог = ',mark*koef,sheet.cell(row=i, column=j))
                sum_period = mark*koef
                count += 1

            print(sheet.cell(row=2, column=j).value)
            if (sheet.cell(row=2, column=j).value).find('АП') != -1:
                print(sum_period,count,sum_period/count)



# Сохраняем книгу
book.save("example.xlsx")
