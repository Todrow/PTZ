import openpyxl as oxl
from openpyxl.styles import Font, Alignment
from xls2xlsx import XLS2XLSX

paths = ['./w.xls', "D:/PTZ/Excel/b.xlsx"]



try:
    x2x = XLS2XLSX(paths[1])
    wb2 = x2x.to_xlsx()
except:
    wb2 = oxl.load_workbook(filename=paths[1])

names = {}
ws2 = wb2.active

for i, el in enumerate(ws2["B"]):
    if el.value[:2] == 'ПЭ':
        names[el.value[4:]] = ws2['U'+str(i+1)].value.split(', ')
print(names)
wb2.close()

try:
    x2x = XLS2XLSX(paths[0])
    wb1 = x2x.to_xlsx()
except:
    wb1 = oxl.load_workbook(filename=paths[0])
ws1 = wb1.active


# Создаём листы

for i, el in enumerate(ws1["F"]):
    knots = el.value.split('; ')
    for each in knots:
        wb1[el].append(ws1[i+1])
wb1.save('./ii.xlsx')
wb1.close()