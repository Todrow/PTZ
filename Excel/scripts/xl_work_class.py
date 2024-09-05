import openpyxl as oxl
from openpyxl.styles import Font
from openpyxl.chart import PieChart3D, Reference
import logging
import copy
import os

logging.basicConfig(level=logging.INFO, filename="prog_log.log",filemode="w", format="%(asctime)s %(levelname)s %(message)s")

class Xl_work:
    """Класс для создания неотформатированной версии отчета
        обладает публичными методами:
        open_file, для загрузки Excel файла в openpyxl
        start, для создания структуры итоговой версии отчета
    
    """
    def __init__(self, web_src: str, bit_src: str, done_src: str) -> None:
        """Создает атрибуты класса Xl_work
        Также проверяет и классифицирует ошибки при загрузке фала пользователем, что в дальнейшем
        определяет выводимое на экран сообщение после ответа сервера

        Args:
            web_src (str):  путь к Excel файлу выгрузки с веб-системы, defaults to None
            bit_src (str): путь к Excel файлу выгрузки с Битрикс24, defaults to None
            done_src (str):  путь к итоговому файлу, defaults to None
        """
        self.paths = [web_src, bit_src]
        self.pathDone = done_src
        self.error = ''
        check1 = self.__correct_file(copy.deepcopy(self.paths[0]))
        check2 = self.__correct_file(copy.deepcopy(self.paths[1]))
        xls_check1 = self.__not_xls(self.paths[0])
        xls_check2 = self.__not_xls(self.paths[1])


        if check1 == 'web' and check2 == 'bitrix':
            self.error = ''
        elif xls_check1 == 'xls error':
            self.error = 'Файл выгрузки с веб-сист в формате xls, нужен xlsx'
        elif xls_check2 == 'xls error':
            self.error = 'Файл выгрузки с битрикса в формате xls, нужен xlsx'
        elif check1 == 'can not read' or check2 == 'can not read':
            if check1 == 'can not read': self.error += 'Файл web не читается'
            if check2 == 'can not read': self.error += 'Файл bitrix не читается'
        elif check1 == check2:
            self.error = 'Файлы одинаковы'
        elif check2 == 'web' and check1 == 'bitrix':
            self.error = 'Файлы перепутаны местами'
        else:
            if check1 == 'unknown':
                self.error += 'Файл web не тот'
            if check2 == 'unknown':
                self.error += 'Файл bitrix не тот'

    def __not_xls(self, path:str)->str:
        """Проверяет, что загружен xls файл, а не xlsx

        Args:
            path (str): Путь к файлу

        Returns:
            str: Возвращает сообщение, которое выводится на экран пользователя,
             если все хорошо - то пустую строку
        """
        ext = os.path.splitext(path)[1]

        if ext == '.xls':
            return 'xls error'
        else:
            return ''

    def __correct_file(self, path:str) -> str:
        """Проверяет соответсвие столбцов в файле исходному шаблону.

        Args:
            path (str): путь либо объект проверяемого файла

        Returns:
            str: Возвращает к какому из 3 типов принадлежит загруженный отчетов - из Битрикса
            из Веб-системы или неизвестного происхождения
        """
        wb = self.open_file(path)
        if wb == False:
            return 'can not read'
        else:
            sheet = wb.active
            if sheet.cell(row=1, column=2).value == 'Название':
                wb.close()
                return 'bitrix'
            elif sheet.cell(row=1, column=7).value == 'Опытный узел':
                wb.close()
                return 'web'
            else:
                wb.close()
                return 'unknown'

    def __delete_unwanted_rows(self)->None:
        """Удаляет строки, не содержащие необходимой информации, например
        строки, дублирующие шапку страницы
        """

        wb = self.open_file(self.paths[1])
        sheet = wb.active

        for i in range(sheet.max_row, 2, -1):
            if (sheet.cell(column=2, row=i).value == 'Название') or (sheet.cell(column=2, row=i).value == None):
                sheet.delete_rows(amount=1, idx=i)

        wb.save(self.paths[1])
        wb.close()

    def __make_link_files(self) -> dict:
        """Создет словрь из названий бюро, задействованных в ПЭ

        Returns:
            dict: Словарь программ ПЭ и бюро, занятых каждой конкретной программой
        """

        wb = self.open_file(self.paths[1])
        
        names = {}
        ws = wb.active

        for i, el in enumerate(ws["B"]):
            if el.value[:2] == 'ПЭ':
                try:
                    names[el.value[4:]] = ws['U'+str(i+1)].value.split(', ')
                except:
                    pass
        wb.close()
        return names
        
    def __create_sheets(self) -> None:
        """Создает в листы с информацией для каждого бюро в итоговом файле
        При этом, передаются только незавершенные записи о ПЭ 
        """
        wb_bit = self.open_file(self.paths[1])
        wb_web = self.open_file(self.paths[0])
        
        ws_web = wb_web.active
        ws_bit = wb_bit.active
        for i in range(1, ws_bit.max_row):
            task = ws_bit.cell(column=2, row=i).value
            if task[:2] == 'ПЭ' and ws_bit.cell(column=10, row=i).value != 'Завершена':
                try:
                    tags = ws_bit.cell(column=21, row=i).value.split(', ')
                    for each in tags:
                        each = each[5:].capitalize()
                        if each in wb_web.sheetnames:
                            continue
                        else: wb_web.create_sheet(each)
                except:
                    pass
        wb_web.create_sheet('Конфликты')
        wb_web.save(self.pathDone)
        wb_web.close()
        wb_bit.close()
    
    def __count_tasks_in_departments(self)->dict:
        """Создает словарь с названиями всех бюро и количество программ ПЭ в каждом из них

        Returns:
            dict: Словарь с названиями всех бюро и количество программ ПЭ в каждом из них
        """

        wb = self.open_file(self.paths[1])
        sheet = wb.active
        amount = {}
        for i in range(1, sheet.max_row):
            task_name = sheet.cell(column=2, row=i).value
            first_two = task_name[:2]
            if (first_two == 'ПЭ'):
                if (sheet.cell(column=10, row=i).value != 'Завершена'):
                        tag = sheet.cell(column=21, row=i).value
                        try:
                            tags = tag.split(', ')
                        except:
                            pass
                        for each in tags:
                            if (each in amount):
                                amount[each] += 1
                            else:
                                amount[each] = 1
        
        return amount

    def __spread_on_sheets(self, links: dict) -> None:
        """Переносит информацию из веб-системы на лист соответсвующего бюро в итоговом файле

        Args:
            links (dict): словарь ключей-названий бюро
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
                    our_row = []
                    for j, cell in enumerate(ws_web[i+1]):
                        if j == 5:
                            our_row.append(each)
                        elif j == ws_web.max_column:
                            continue
                        else:
                            our_row.append(cell.value)
                    wb_done['Конфликты'].append(our_row)
        wb_done.save(self.pathDone)
        wb_done.close()
        wb_web.close()

    def __count_unique(self, column: int, path = None, sheet = None)->int:
        """Рассчитывает число уникальных элементов, получив на вход номер столбца с идентификаторами, а 
        также путь к файлу или лист уже открытого

        Args:
            column (int): Номер колонки, где находятся индексы
            path (_type_, optional): Путь к файлу. Defaults to None.
            sheet (_type_, optional): Лист в открытом файле. Defaults to None.

        Returns:
            int: Число уникальных машин
        """
        if sheet:
            unique_elems = []

            for i in range(2, sheet.max_row):
                if (sheet.cell(column=column, row=i).value not in unique_elems):
                    unique_elems.append(sheet.cell(column=column, row=i).value)
            return(len(unique_elems))
        if path:
            wb = self.open_file(path)

            sheet = wb.active

            unique_elems = []

            for i in range(2, sheet.max_row):
                if (sheet.cell(column=column, row=i).value not in unique_elems):
                    unique_elems.append(sheet.cell(column=column, row=i).value)
            wb.save(path)
            wb.close()
            return (len(unique_elems))
        
    def __stat(self)->None:
        """Создает в итоговом файле лист со статистикой
        На этом листе выводит общее количество тракторов, которое находит через функцию __count_number_of_machines,
        список всех бюро с количеством проектов ПЭ в каждом из них на основе словаря,
        полученного из функции __count_tasks_in_departments и создает круговую диграмму распределения проектов ПЭ по бюро

        Returns:
            None
        
        """

        num_of_machines = self.__count_unique(path=self.paths[0], column=2)
        num_of_programms = self.__count_unique(path=self.paths[0], column=6)
        path = self.pathDone
        data = self.__count_tasks_in_departments()

        wb = oxl.load_workbook(filename=path)
        wb.create_sheet('Статистика')
        wb.active=wb['Статистика']
        sheet = wb.active

        sheet.column_dimensions['A'].width = 36.29
        sheet.column_dimensions['B'].width = 12
        sheet.column_dimensions['C'].width = 25

        """Выведение числа тракторов"""
        sheet.cell(column=1, row=1).value = 'Количество тракторов'
        sheet.cell(column=1, row=1).font = Font(name='Times New Roman', bold=True, size=12)
        sheet.cell(column=2, row=1).value = num_of_machines
        sheet.cell(column=2, row=1).font = Font(name='Times New Roman', bold=False, size=12)

        '''Выведение общ. числа ПЭ'''
        sheet.cell(column=1, row=2).value = 'Количество программ ПЭ'
        sheet.cell(column=1, row=2).font = Font(name='Times New Roman', bold=True, size=12)
        sheet.cell(column=2, row=2).value = num_of_programms
        sheet.cell(column=2, row=2).font = Font(name='Times New Roman', bold=False, size=12)

        """Таблица с числом проектов у каждого бюро"""
        sheet.cell(column=1, row=4).value = 'Бюро'
        sheet.cell(column=1, row=4).font = Font(name='Times New Roman', bold=True, size=12)
        sheet.cell(column=2, row=4).value = 'Кол-во ПЭ'
        sheet.cell(column=2, row=4).font = Font(name='Times New Roman', bold=True, size=12)
        sheet.cell(column=3, row=4).value = 'Кол-во тракторов в ПЭ'
        sheet.cell(column=3, row=4).font = Font(name='Times New Roman', bold=True, size=12)
        for row_index, (key, value) in enumerate(data.items(), start=5):
            sheet[f'A{row_index}'] = key
            sheet[f'B{row_index}'] = value
            sheet[f'C{row_index}'] = self.__count_unique(column=2, sheet=wb[wb.sheetnames[row_index-4]])

        for i in range(1, len(wb.sheetnames)-2):
            sheet.cell(row = 4+i, column=1).hyperlink = f"#'{wb.sheetnames[i]}'!A1"


        """Создание диаграммы"""
        chart = PieChart3D()
        labels = Reference(sheet, min_col=1, min_row=5, max_row=sheet.max_row, max_col=1)
        info = Reference(sheet, min_col=2, min_row=4, max_row=sheet.max_row, max_col=2)
        chart.add_data(info, titles_from_data=True)
        chart.set_categories(labels)

        chart.width = 15
        chart.height = 12

        chart.legend.position = 'b'

        chart.title = 'Программы ПЭ в бюро'
        chart.series[0].explosion = 10  


        sheet.add_chart(chart, 'E1')

        wb.move_sheet(wb.active, offset=-(len(wb.sheetnames) - 1))
        wb.active = wb['Статистика']
        
        wb.save(self.pathDone)
        wb.close()
    
    def open_file(self, path:str):
        """Открывает файл

        Args:
            path (str): путь к файлу

        Returns:
            _type_: Excel WorkBook
        """

        try:
            wb = oxl.load_workbook(filename=path)
        except:
            return False
        return wb

    def start(self) -> None:
        """ Запускает весь процесс совмещения двух файлов, перед этим выполняя проверку на ошибки при добавлении исходных файлов
        в итогоговом файле распределяяет информацию по листам с бюро и создает лист статистики
        
        """
        self.__delete_unwanted_rows()
        self.__create_sheets()
        self.__spread_on_sheets(self.__make_link_files())
        self.__stat()