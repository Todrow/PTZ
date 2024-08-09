import openpyxl as oxl
from openpyxl.styles import Font, Alignment
from xls2xlsx import XLS2XLSX

class Xl_work:
    def __init__(self, web_src: str, bit_src: str) -> None:
        self.Paths = [web_src, bit_src]
        self.pathDone = './hh.xlsx'
    
    def __make_link_files(self) -> dict:
        try:
            x2x = XLS2XLSX(self.Paths[1])
            wb = x2x.to_xlsx()
        except:
            wb = oxl.load_workbook(filename=self.Paths[1])
        
        names = {}
        ws = wb.active

        for i, el in enumerate(ws["B"]):
            if el.value[:2] == 'ПЭ':
                names[el.value[4:]] = ws['U'+str(i+1)].value.split(', ')
        wb.close()
        return names

    def __create_sheets(self) -> None:
        try:
            x2x = XLS2XLSX(self.Paths[1])
            wb_bit = x2x.to_xlsx()
        except:
            wb_bit = oxl.load_workbook(filename=self.Paths[1])
        try:
            x2x = XLS2XLSX(self.Paths[0])
            wb_web = x2x.to_xlsx()
        except:
            wb_web = oxl.load_workbook(filename=self.Paths[0])
        ws_web = wb_web.active
        ws_bit = wb_bit.active
        for i in range(1, ws_bit.max_row):
            task = ws_bit.cell(column=2, row=i).value
            if task[:2] == 'ПЭ' and ws_bit.cell(column=10, row=i).value != 'Завершена':
                tags = ws_bit.cell(column=21, row=i).value.split(', ')
                for each in tags:
                    each = each[5:]
                    if each in wb_web.sheetnames:
                        continue
                    else: wb_web.create_sheet(each)
        wb_web.create_sheet('Мусорка')
        wb_web.save(self.pathDone)
        wb_web.close()
        wb_bit.close()
    
    def __spread_on_sheets(self, links: dict) -> None:
        try:
            x2x = XLS2XLSX(self.Paths[0])
            wb_web = x2x.to_xlsx()
        except:
            wb_web = oxl.load_workbook(filename=self.Paths[0])
        ws_web = wb_web.active
        first = True
        for i, el in enumerate(ws_web["F"]):
            if first:
                first = False
                continue
            knots = el.value.split('; ')
            for each in knots:
                if each in links.keys():
                    for byros in links[each]:
                        wb_web[byros[5:]].append([cell.value for cell in ws_web[i+1]])
                else:
                    wb_web['Мусорка'].append([cell.value for cell in ws_web[i+1]])
        wb_web.save(self.pathDone)
        wb_web.close()
    
    def start(self) -> None:
        self.__create_sheets()
        self.__spread_on_sheets(self.__make_link_files())
