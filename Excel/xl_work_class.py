import openpyxl as oxl
from xls2xlsx import XLS2XLSX

class Xl_work:
    def __init__(self, web_src: str, bit_src: str) -> None:
        self.paths = [web_src, bit_src]
        self.pathDone = 'C:/Users/pinchukna/Documents/GitHub/PTZ/Excel/done.xlsx'
    
    def __make_link_files(self) -> dict:
        try:
            x2x = XLS2XLSX(self.paths[1])
            wb = x2x.to_xlsx()
        except:
            wb = oxl.load_workbook(filename=self.paths[1])
        
        names = {}
        ws = wb.active

        for i, el in enumerate(ws["B"]):
            if el.value[:2] == 'ПЭ':
                names[el.value[4:]] = ws['U'+str(i+1)].value.split(', ')
        wb.close()
        return names

    def __create_sheets(self) -> None:
        try:
            x2x = XLS2XLSX(self.paths[1])
            wb_bit = x2x.to_xlsx()
        except:
            wb_bit = oxl.load_workbook(filename=self.paths[1])
        try:
            x2x = XLS2XLSX(self.paths[0])
            wb_web = x2x.to_xlsx()
        except:
            wb_web = oxl.load_workbook(filename=self.paths[0])
        ws_web = wb_web.active
        ws_bit = wb_bit.active
        for i in range(1, ws_bit.max_row):
            task = ws_bit.cell(column=2, row=i).value
            if task[:2] == 'ПЭ' and ws_bit.cell(column=10, row=i).value != 'Завершена':
                tags = ws_bit.cell(column=21, row=i).value.split(', ')
                for each in tags:
                    each = each[5:].capitalize()
                    if each in wb_web.sheetnames:
                        continue
                    else: wb_web.create_sheet(each)
        wb_web.create_sheet('Мусорка')
        wb_web.save(self.pathDone)
        wb_web.close()
        wb_bit.close()
    
    def __spread_on_sheets(self, links: dict) -> None:
        
        wb_web = self.open_file(self.paths[0])
        wb_done = self.open_file(self.pathDone)

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
                        our_row = []
                        for j, cell in enumerate(ws_web[i+1]):
                            if j == 5:
                                our_row.append(each)
                            elif j == ws_web.max_column-1:
                                continue
                            else:
                                our_row.append(cell.value)
                        wb_done[byros[5:].capitalize()].append(our_row)
                else:
                    wb_done['Мусорка'].append([cell.value for cell in ws_web[i+1]])
        wb_done.save(self.pathDone)
        wb_done.close()
        wb_web.close()
    
    def open_file(self, path:str):
        try:
            x2x = XLS2XLSX(path)
            wb = x2x.to_xlsx()
        except:
            wb = oxl.load_workbook(filename=path)
        return wb

    def start(self) -> None:
        self.__create_sheets()
        self.__spread_on_sheets(self.__make_link_files())
