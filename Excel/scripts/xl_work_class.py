import openpyxl as oxl
from xls2xlsx import XLS2XLSX
from openpyxl.styles import Font, Alignment, colors, Color, Border, Side
from openpyxl.chart import PieChart, Reference, Series
from openpyxl.chart.label import DataLabelList 
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.shapes import GraphicalProperties
import os
import logging

logging.basicConfig(level=logging.INFO, filename="prog_log.log",filemode="w", format="%(asctime)s %(levelname)s %(message)s")

class Xl_work:
    """Класс для создания неотформатированной версии отчета
        обладает публичными методами:
        open_file, для загрузки Excel файла в openpyxl
        start, для создания структуры итоговой версии отчета
    
    """
    def __init__(self, web_src: str, bit_src: str, done_src: str) -> None:
        """Создает атрибуты класса Xl_work

        :param web_src: путь к Excel файлу выгрузки с веб-системы, defaults to None
        :type web_src: str

        :param bit_src: путь к Excel файлу выгрузки с веб-системы, defaults to None
        :type bit_src: str

        :param done_src: путь к итоговому файлу, defaults to None

        :return None
        """
        self.paths = [web_src, bit_src]
        self.pathDone = done_src

    def __correct_file_B(self)->bool:
        """Проверяет соответсвие столбцов в файле из Битркса исходному шаблону
        :rtype: bool
        """
        try:
            x2x = XLS2XLSX(self.paths[1])
            wb = x2x.to_xlsx()
        except:
            wb = oxl.load_workbook(filename=self.paths[1])

        sheet = wb.active
        if sheet.cell(row=1, column=2).value == 'Название':
            wb.close()
            print('ok')
            return True
        else:
            wb.close()
            return False
           
    def __correct_file_W(self)->bool:
        """Проверяет соответсвие столбцов в файле из Веб-системы исходному шаблону
        :rtype: bool
        """
        try:
            x2x = XLS2XLSX(self.paths[0])
            wb = x2x.to_xlsx()
        except:
            wb = oxl.load_workbook(filename=self.paths[0])

        sheet = wb.active
        if sheet.cell(row=1, column=6).value == 'Опытный узел':
            wb.close()
            print('ok')
            return True
        else:
            wb.close()
            return False

    def __make_link_files(self) -> dict:
        """Создет словрь из названий бюро, задействованных в ПЭ

        :rtype: dict
        
        """
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
        """ Создает в листы с информацией для каждого бюро в итоговом файле
        При этом, передаются только незавершенные записи о ПЭ 

        :rtype:None

        """
        wb_bit = self.open_file(self.paths[1])
        wb_web = self.open_file(self.paths[0])
        
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
    
    def __count_tasks_in_departments(self)->dict:
        """Создает словарь с названиями всех бюро и количество программ ПЭ в каждом из них
        return:  
        rtype: dict
        """

        wb = oxl.load_workbook(self.paths[1])
        sheet = wb.active
        amount = {}
        for i in range(1, sheet.max_row):
            task_name = sheet.cell(column=2, row=i).value
            first_two = task_name[:2]
            if (first_two == 'ПЭ'):
                if (sheet.cell(column=10, row=i).value != 'Завершена'):
                        tag = sheet.cell(column=21, row=i).value
                        tags = tag.split(', ')
                        for each in tags:
                            if (each in amount):
                                amount[each] += 1
                            else:
                                amount[each] = 1
        
        return amount

    def __spread_on_sheets(self, links: dict) -> None:
        """Переносит информацию из веб-системы на лист соответсвующего бюро в итоговом файле

        :param links: словарь ключей-названий бюро
        :type links: dict
        
        """
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
          
    def __count_number_of_machines(self, path)->int:
        """Возвращает количество уникальных машин в ПЭ, взяв их из файла выгруженного из веб-системы

        :param path: Путь к файлу веб-системы
        :type path: str

        :rtype: int

        """
        
        wb = self.open_file(path)     

        sheet = wb.active 

        unique_machines = []

        for i in range(2, sheet.max_row):
            if (sheet.cell(column=2, row=i).value not in unique_machines):
                unique_machines.append(sheet.cell(column=2, row=i).value)

        wb.save(path)
        wb.close()
        return (len(unique_machines))

    def __stat(self)->None:
        """Создает в итоговом файле лист со статистикой
        На этом листе выводит общее количество тракторов, которое находит через функцию __count_number_of_machines,
        список всех бюро с количеством проектов ПЭ в каждом из них на основе словаря,
        полученного из функции __count_tasks_in_departments и создает круговую диграмму распределения проектов ПЭ по бюро

        :rtype: None
        
        """

        num_of_machines = self.__count_number_of_machines()
        path = self.pathDone
        data = self.__count_tasks_in_departments()

        wb = oxl.load_workbook(filename=path)
        wb.create_sheet('Статистика')
        wb.active=wb['Статистика']
        sheet = wb.active


        """Выведение числа тракторов"""
        sheet.cell(column=1, row=1).value = 'Количество тракторов'
        sheet.cell(column=1, row=1).font = Font(name='Times New Roman', bold=True, size=12)
        sheet.cell(column=2, row=1).font = Font(name='Times New Roman', bold=False, size=12)
        sheet.column_dimensions['A'].width = 30

        sheet.cell(column=2, row=1).value = num_of_machines

        """Таблица с числом проектов у каждого бюро"""
        for row_index, (key, value) in enumerate(data.items(), start=3):
            sheet[f'A{row_index}'] = key
            sheet[f'B{row_index}'] = value

        thin = Side(border_style="thin", color="000000")
        border = Border(top=thin, left=thin, right=thin, bottom=thin)

        """Создание диаграммы"""
        dll = DataLabelList(showVal=True)
        chart = PieChart()
        labels = Reference(sheet, min_col=1, min_row=3, max_row=sheet.max_row, max_col=1)
        info = Reference(sheet, min_col=2, min_row=2, max_row=sheet.max_row, max_col=2)
        chart.add_data(info, titles_from_data=True)
        chart.set_categories(labels)
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showVal = True
        chart.width = 20
        chart.height = 10
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showVal = True
        chart.dataLabels.showCatName = True
        chart.dataLabels.showPercentage = False
        data_label_font = Font(size=14)
        chart.dataLabels.font = data_label_font
        chart.legend = None

        sheet.add_chart(chart, 'D5')
        
        
        wb.save(self.pathDone)
        wb.close()
    
    def open_file(self, path):
        """Открывает файл как work_book в openpyxl

        :param path: путь к файлу, который необходимо открыть
        :type path: str 
        
        """

        try:
            x2x = XLS2XLSX(path)
            wb = x2x.to_xlsx()
        except Exception as e:
            print(e)
            wb = oxl.load_workbook(filename=path)
        return wb

    def start(self) -> None:
        """ Запускает весь процесс совмещения двух файлов, перед этим выполняя проверку на ошибки при добавлении исходных файлов
        в итогоговом файле распределяяет информацию по листам с бюро и создает лист статистики

        :rtype: None
        
        """

        if self.__correct_file_B() == False:
            logging.critical('Unexpected Bitrix file structure',exc_info=True)
        elif self.__correct_file_W() == False:
            logging.critical('Unexpected Web-sys file structure',exc_info=True)
        else:
            self.__create_sheets()
            self.__spread_on_sheets(self.__make_link_files())
            self.__stat()
            logging.info('Program completed successfully',exc_info=True)
