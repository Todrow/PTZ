# Django importing
from django.shortcuts import render
from django.http import HttpResponse, FileResponse, JsonResponse
from django.core.files.storage import FileSystemStorage
from .models import XL_file

# Python Libs
import os

# Project classes
from scripts.exlWrapper import ExcelWrapper
from scripts.xl_work_class import Xl_work

def merge_files(path_bitrix: str, path_web: str, path_done: str):
    xl = Xl_work(path_web, path_bitrix, path_done)
    ew = ExcelWrapper()
    xl.start()
    wb = xl.open_file(path_done)
    for sheet in wb.sheetnames[1:-2]:
        ew.formatTitles(wb[sheet], True)
        ew.formattingCells(wb[sheet])
    wb.save(path_done)
    wb.close()


def index(request):
    path_done = 'c:/Users/pinchukna/Documents/GitHub/PTZ/Excel/django/excel/done.xlsx'
    if request.method == 'POST' and request.FILES:
        file1 = request.FILES['file_bitrix']
        # fs = FileSystemStorage()
        # # filename = fs.save(file1.name, file1)
        # # file_bitrix_url = os.path.abspath(fs.url(filename))
        # ###
        file2 = request.FILES['file_web']
        # fs = FileSystemStorage()
        # # filename = fs.save(file2.name, file2)
        # # file_web_url = os.path.abspath(fs.url(filename))
        merge_files(file1, file2, path_done)
        return JsonResponse({"resp": str(request.FILES)})
    else:
        return render(request, 'index.html')
    # return render(request, 'index.html', context={'jj': str('NOGET')})

def download_file(request):
    path_done = 'c:/Users/pinchukna/Documents/GitHub/PTZ/Excel/django/excel/done.xlsx'
    with open(path_done, 'rb') as file:
        response = HttpResponse(file.read())
        response['Content-Disposition'] = 'attachment; filename=' + 'path_done.xlsx'
    os.remove(path_done)
    return response
