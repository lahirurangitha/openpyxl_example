from __future__ import division
import openpyxl

book = openpyxl.load_workbook('test.xlsx')

sheetNames = book.get_sheet_names()
sheetOne = book.get_sheet_by_name(sheetNames[0])
sheetTwo = book.get_sheet_by_name(sheetNames[1])

# print (sheetOne['A1'].value)
# print (sheetTwo['A1'].value)

a = []
for row in sheetTwo['A2':'A11']:
    for cell in row:
        a.append(cell.value)
# print a

b = []
for row in sheetTwo['B2':'B11']:
    for cell in row:
        b.append(cell.value)
# print b

c = []
for row in sheetTwo['C2':'C11']:
    for cell in row:
        c.append(cell.value)
# print c

d = []
for row in sheetTwo['D2':'D11']:
    for cell in row:
        d.append(cell.value)
# print d

e = []
for row in sheetTwo['E2':'E11']:
    for cell in row:
        e.append(cell.value)
# print e

f = []
for row in sheetTwo['F2':'F11']:
    for cell in row:
        f.append(cell.value)
# print f

h = []
for row in sheetTwo['H1':'H14']:
    for cell in row:
        h.append(cell.value)
# print h

i = []
for row in sheetTwo['I1':'I14']:
    for cell in row:
        i.append(cell.value)
# print i



# creating header line
h_line = 5
d_line = h_line + 1
for row in sheetOne['A'+str(h_line):'J'+str(h_line)]:
    index = 0
    for cell in row:
        cell.value = a[index]
        index = index + 1

for row in sheetOne['K'+str(h_line):'T'+str(h_line)]:
    index = 0
    for cell in row:
        cell.value = a[index]+'_angle'
        index = index + 1

for row in sheetOne['U'+str(h_line):'AC'+str(h_line)]:
    index = 1
    for cell in row:
        cell.value = a[index]+'_bala'
        index = index + 1

for row in sheetOne['AD'+str(h_line):'AQ'+str(h_line)]:
    index = 0
    for cell in row:
        cell.value = h[index]
        index = index + 1

# creating data line
for row in sheetOne['A'+str(d_line):'J'+str(d_line)]:
    index = 0
    for cell in row:
        cell.value = b[index]
        index = index + 1

for row in sheetOne['K'+str(d_line):'T'+str(d_line)]:
    index = 0
    for cell in row:
        cell.value = (c[index]*60+d[index])/60
        index = index + 1

for row in sheetOne['U'+str(d_line):'AC'+str(d_line)]:
    index = 1
    for cell in row:
        cell.value = f[index]
        index = index + 1

for row in sheetOne['AD'+str(d_line):'AQ'+str(d_line)]:
    index = 0
    for cell in row:
        cell.value = i[index]
        index = index + 1


book.save('test.xlsx')




