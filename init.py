from __future__ import division
import openpyxl

# new workbook
resultBook = openpyxl.Workbook()
dest_filename = 'resultSet.xlsx'
sheetResult = resultBook.active
sheetResult.title = "results"
sheetResult = resultBook.get_sheet_by_name("results")

# retrieve data set
book = openpyxl.load_workbook('test.xlsx')

sheetNames = book.get_sheet_names()
toHeader = book.get_sheet_by_name(sheetNames[1])

# creating header line
a = []
for row in toHeader['A2':'A11']:
    for cell in row:
        a.append(cell.value)
# print a
h = []
for row in toHeader['H1':'H14']:
    for cell in row:
        h.append(cell.value)
# print h

h_line = 1
d_line = h_line + 1
for row in sheetResult['A' + str(h_line):'J' + str(h_line)]:
    index = 0
    for cell in row:
        cell.value = a[index]
        index = index + 1

for row in sheetResult['K' + str(h_line):'T' + str(h_line)]:
    index = 0
    for cell in row:
        cell.value = a[index] + '_angle'
        index = index + 1

for row in sheetResult['U' + str(h_line):'AC' + str(h_line)]:
    index = 1
    for cell in row:
        cell.value = a[index] + '_bala'
        index = index + 1

for row in sheetResult['AD' + str(h_line):'AQ' + str(h_line)]:
    index = 0
    for cell in row:
        cell.value = h[index]
        index = index + 1


# data lines creation
d_line = 2
for sheet in book.worksheets:
    a = []
    for row in sheet['A2':'A11']:
        for cell in row:
            a.append(cell.value)
    # print a
    b = []
    for row in sheet['B2':'B11']:
        for cell in row:
            b.append(cell.value)
    # print b
    c = []
    for row in sheet['C2':'C11']:
        for cell in row:
            c.append(cell.value)
    # print c
    d = []
    for row in sheet['D2':'D11']:
        for cell in row:
            d.append(cell.value)
    # print d
    e = []
    for row in sheet['E2':'E11']:
        for cell in row:
            e.append(cell.value)
    # print e
    f = []
    for row in sheet['F2':'F11']:
        for cell in row:
            f.append(cell.value)
    # print f
    h = []
    for row in sheet['H1':'H14']:
        for cell in row:
            h.append(cell.value)
    # print h
    i = []
    for row in sheet['I1':'I14']:
        for cell in row:
            i.append(cell.value)
            # print i

    # creating data line
    for row in sheetResult['A' + str(d_line):'J' + str(d_line)]:
        index = 0
        for cell in row:
            cell.value = b[index]
            index = index + 1

    for row in sheetResult['K' + str(d_line):'T' + str(d_line)]:
        index = 0
        for cell in row:
            cell.value = (c[index] * 60 + d[index]) / 60
            index = index + 1

    for row in sheetResult['U' + str(d_line):'AC' + str(d_line)]:
        index = 1
        for cell in row:
            cell.value = f[index]
            index = index + 1

    for row in sheetResult['AD' + str(d_line):'AQ' + str(d_line)]:
        index = 0
        for cell in row:
            cell.value = i[index]
            index = index + 1
    d_line = d_line + 1

resultBook.save(dest_filename)
# book.save('test.xlsx')
