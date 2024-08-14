from xl_work_class import Xl_work
from exlWrapper import ExcelWrapper

xl = Xl_work(r'C:\Users\Aleksandr\Documents\Работа\Excel\Отчет по ПЭ от 02.08.2024.xls', r'C:\Users\Aleksandr\Documents\Работа\Excel\Выгрузка задачи Битрикс.xlsx')
ew = ExcelWrapper()

xl.start()
pathDone = r'C:\Users\Aleksandr\Documents\Работа\Excel\Report.xlsx'
wb = xl.open_file(pathDone)

for sheet in wb.sheetnames[1:-1]:
    ew.formatTitles(wb[sheet], True)
    ew.formattingCells(wb[sheet])
wb.save(pathDone)
wb.close()
