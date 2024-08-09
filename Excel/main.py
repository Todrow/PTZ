from xl_work_class import Xl_work
from exlWrapper import ExcelWrapper

xl = Xl_work('C:/Users/pinchukna/Documents/GitHub/PTZ/Excel/w.xls', 'C:/Users/pinchukna/Documents/GitHub/PTZ/Excel/b.xlsx')
ew = ExcelWrapper()

xl.start()
pathDone = 'C:/Users/pinchukna/Documents/GitHub/PTZ/Excel/done.xlsx'
wb = xl.open_file(pathDone)

for sheet in wb.sheetnames[1:]:
    ew.formatTitles(wb[sheet], True)
    ew.formattingCells(wb[sheet])
wb.save(pathDone)
wb.close()