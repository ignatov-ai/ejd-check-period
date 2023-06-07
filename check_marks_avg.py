import itertools
import os
from itertools import permutations, product

import openpyxl
from openpyxl.styles import Alignment, PatternFill, Font, Fill, NamedStyle, Border
from math import ceil

file_name_all_periods = '9ч.xlsx'

book = openpyxl.load_workbook(filename='data/' + file_name_all_periods)

sheet_names = book.sheetnames
print(sheet_names)

for s in range(len(sheet_names)):
    sheet = book.worksheets[s]

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
        if i == 150:
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
        for row in range(3,40):
            cell = ''.join(c)+str(row)
            #print(cell,sheet[cell].value)
            if sheet[cell].value in ['2', '3', '4', '5']:
                sheet[cell].value = int(sheet[cell].value)
            if cell[:2] == 'EU':
                stop = True
                break
        if stop:
            break

    # записываем данные в массив data
    data = []
    for row in sheet.iter_rows(values_only=True):
        row_data = []
        for cell in row:
            if cell is not None:
                row_data.append(cell)
            else:
                row_data.append('')
        data.append(row_data)


# Сохраняем книгу
book.save("example.xlsx")
