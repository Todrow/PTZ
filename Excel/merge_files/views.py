# Django importing
from django.shortcuts import render
from django.http import HttpResponse, JsonResponse

# Python Libs
import os
import pika
from json import dumps
import asyncio

"""Views отвечает за принятие запросов, обработку данных из них и возрат ответов пользователю
"""

path_save = './uploads/'

# Create your views here.

def render_merge(request): 
    match request.method:
        case 'GET':
            return render(request, 'index.html')
        case 'POST':
            body = RequestParser.parse(request)
            p = Sender("merge")
            return p.send(message="123", id="1")
        
def download_file(request, id):
    match request.method:
        case 'GET':
            return FileTaker.get_file(path_save, id)
        case _:
            return None
    

class RequestParser():
    @staticmethod
    def parse(request) -> dict:
        response = {
            'path_bitrix': request.FILES['file_bitrix'] or None,
            'path_web': request.FILES['file_web'] or None,
            'path_done': path_save+request.META['HTTP_ID']+'.xlsx'
        }
        
        return response
        

class FileTaker():
    def get_file(path, id):
        with open(path+id+'.xlsx', 'rb') as file:
            response = HttpResponse(file.read())
            response['Content-Disposition'] = 'attachment; filename=' + 'finished_report.xlsx'
        os.remove(path+id+'.xlsx')
        return response
    
class Sender():
    def __init__(self, name):
        self.name = name
        credentials = pika.PlainCredentials('myuser', 'mypassword')
        self.connection = pika.BlockingConnection(
            pika.ConnectionParameters(
                host='localhost',
                port=5672,
                virtual_host='myvhost',  # Check if this exists
                credentials=credentials
            )
        )
        self.channel = self.connection.channel()
        self.sends = f"{name}_sends"
        self.responses = f"{name}_responses"
        
        # очередь в которую отправляются запросы
        self.channel.queue_declare(queue=self.sends)
        
        # очередь из которой ждет ответ
        self.channel.queue_declare(queue=self.responses)
        self.channel.basic_consume(queue=self.responses, on_message_callback=self._handle_response)
        
        
    def _send(self, message, id):
        self.channel.basic_publish(exchange='',
                                   routing_key=self.sends,
                                   properties=pika.BasicProperties(
                                       message_id=str(id),
                                   ),
                                   body=message)
        
        print("сообщение отправлено в очередь для обработчика")
        
    def _handle_response(self, ch, method, properties, body):
        print(f'Получил ответ: {body}')
        ch.basic_ack(delivery_tag=method.delivery_tag)  # Подтверждение
        # self.channel.stop_consuming()   Остановка после ответа
            
    async def _wait_for_response(self):
        await self.channel.start_consuming()
        
    def send(self, message, id):
        self._send(message, id)
        self._wait_for_response()

        

'''
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

    Если постцпает запрос с GET, то пользователю выводится стандартная страница index.html

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
    
    match request.method:
        case 'POST':
            if len(request.FILES) == 2:
                file1 = request.FILES['file_bitrix']
                ####
                file2 = request.FILES['file_web']
                ####
                error, message = merge_files(file1, file2, path_done+request.META['HTTP_ID']+'.xlsx')
                print('ОШИБКА:', error)
                print('Сообщение:', message)
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
        case 'GET':
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
'''