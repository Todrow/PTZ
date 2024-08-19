# Django importing
from django.shortcuts import render
from django.http import HttpResponse, FileResponse, JsonResponse
from django.core.files.storage import FileSystemStorage

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

path_done = 'c:/Users/pinchukna/Documents/GitHub/PTZ/Excel/done.xlsx'
def index(request):
    global path_done
    try:
        os.remove(path_done)
    except:
        pass
    if request.method == 'POST' and request.FILES:
        file1 = request.FILES['file_bitrix']
        if file1.name.split('.')[-1] == 'xls':
            fs = FileSystemStorage()
            fs.save(file1.name, file1)
            file1 = os.path.abspath('./uploads/'+file1.name)
        ####
        file2 = request.FILES['file_web']
        if file2.name.split('.')[-1] == 'xls':
            fs = FileSystemStorage()
            fs.save(file2.name, file2)
            file2 = os.path.abspath('./uploads/'+file2.name)
        ####
        merge_files(file1, file2, path_done)
        ####
        try:
            os.remove(file1)
        except:
            pass
        try:
            os.remove(file2)
        except:
            pass
        return JsonResponse({"resp": str(file1.name)})
    else:
        return render(request, 'index.html')

def download_file(request):
    global path_done
    with open(path_done, 'rb') as file:
        response = HttpResponse(file.read())
        response['Content-Disposition'] = 'attachment; filename=' + 'finished_report.xlsx'
    os.remove(path_done)
    return response
