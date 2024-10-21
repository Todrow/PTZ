# Django importing
from django.shortcuts import render
from django.http import HttpResponse, JsonResponse

# Python Libs
import os

# Project classes
from scripts.exlWrapper import ExcelWrapper
from scripts.xl_work_class import Xl_work

"""Views отвечает за принятие запросов, обработку данных из них и возрат ответов пользователю
"""

def merge_files(path_bitrix, path_web, path_done) -> (str, str):
    """Соединяет 2 загруженных файла в единый отчет, используя классы Xl_work и ExcelWrapper

    Args:
        path_bitrix (str): Путь к файлу с выгрузской из Битрикса
        path_web (str): Путь к файлу с выгрузкой из Веб-системы
        path_done (str): Путь к итоговому файл

    Returns:
        str: Возвращает значение ошибки, которое используется для вывода в всплывающем окне
        если ошибок нет, то значение - пустая строка
    """
    xl = Xl_work(path_web, path_bitrix, path_done)
    if xl.error == '':
        ew = ExcelWrapper(['Вложения', 'Последний раз обновлено', 'Статус', 'Наименование сервисного центра'], ['ПЭ: дата время', 'ПЭ: Комментарий', 'ПЭ: наработка м/ч'], path_web)
        ew.format()
        xl.start()
        wb = xl.open_file(path_done)
        for sheet in wb.sheetnames[2:]:
            ew.formatTitles(wb[sheet], True)
            ew.formattingCells(wb[sheet])
        wb.save(path_done)
        wb.close()
        xl.department_stat()
    return xl.error, xl.message


path_done = './uploads/'
def index(request):
    """Если приходит запрос, содержащий необходимые файлы, с методом POST, то возвращается json файл с id и
    значение ошибки. Значение ошибки формируется в функции merge_files, которая в свою очередь использует класс
    Xl_work для этого

    Если запрос не поступал, то пользователю выводится стандартная страница index.html

    Готовый файл отчета храниться на сервере 5 минут (300 секунд)

    Args:
        request (_type_): Запрос к серверу

    Returns:
        _type_: _description_
    """
    global path_done
    try:
        import time
        for filename in os.listdir(path_done):
            f = os.path.join(path_done, filename)
            if time.time() - os.path.getctime(f) > 300:
                os.remove(f)
    except:
        pass
    if request.method == 'POST' and request.FILES:
        if len(request.FILES) == 2:
            file1 = request.FILES['file_bitrix']
            ####
            file2 = request.FILES['file_web']
            ####
            error, message = merge_files(file1, file2, path_done+request.META['HTTP_ID']+'.xlsx')
        elif len(request.FILES) == 1:
            file = request.FILES['file_web']
            ####
            error, message = merge_files(
                None, file, path_done+request.META['HTTP_ID']+'.xlsx')
        ####
        try:
            os.remove(file1)
        except:
            pass
        try:
            os.remove(file2)
        except:
            pass
        return JsonResponse({"id": str(request.META['HTTP_ID']), "error": error, "message": message})
    else:
        return render(request, 'index.html')
    

def download_file(request, id):
    """Возвращает файл при нажатии на ссылку 'Скачать'

    Args:
        request (_type_): Запрос к серверу на скачивание
        id (_type_): id, по которому определяется путь к готовому файлу 
        (Это необходимо для загрузки от нескольких пользователей одновременно)

    Returns:
        _type_: Отправлемый файл готового отчета
    """
    global path_done
    with open(path_done+id+'.xlsx', 'rb') as file:
        response = HttpResponse(file.read())
        response['Content-Disposition'] = 'attachment; filename=' + 'finished_report.xlsx'
    os.remove(path_done+id+'.xlsx')
    return response
