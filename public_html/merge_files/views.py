# Django importing
from django.shortcuts import render
from django.http import HttpResponse, JsonResponse

# Python Libs
import os

# Project classes
from scripts.exlWrapper import ExcelWrapper
from scripts.xl_work_class import Xl_work

def merge_files(path_bitrix: str, path_web: str, path_done: str) -> str:
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
    return xl.error

path_done = '/var/www/PTZ/public_html/uploads/'
def index(request):
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
        file1 = request.FILES['file_bitrix']
        ####
        file2 = request.FILES['file_web']
        ####
        error = merge_files(file1, file2, path_done+request.META['HTTP_ID']+'.xlsx')
        ####
        try:
            os.remove(file1)
        except:
            pass
        try:
            os.remove(file2)
        except:
            pass
        return JsonResponse({"id": str(request.META['HTTP_ID']), "error": error})
    else:
        return render(request, 'index.html')

def download_file(request, id):
    global path_done
    with open(path_done+id+'.xlsx', 'rb') as file:
        response = HttpResponse(file.read())
        response['Content-Disposition'] = 'attachment; filename=' + 'finished_report.xlsx'
    os.remove(path_done+id+'.xlsx')
    return response