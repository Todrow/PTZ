import pika
from abc import abstractmethod, ABC
from ..scripts.xl_work_class import Xl_work
from ..scripts.exlWrapper import ExcelWrapper
from json import dumps



config = {
    'path_done': './uploads/',
}

class Consumer():
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
        
    def _callback(self, ch, method, properties, body):
        print(f'Обработчик получил запрос: {body}')
        
        id = properties.message_id
        job = body.decode('utf-8')['job']
        files = body.decode('utf-8')['files']
        error, message = WorkerFactory.chooseWorker(job, files, id).format()
        
        
        # Отправка ответа
        self.channel.basic_publish(
            exchange='',
            routing_key=self.responses,
            body=f"Ответ от {self.name}: {body}"
        )
        ch.basic_ack(delivery_tag=method.delivery_tag)  # Подтверждение
        
    def _recieve(self):
        self.channel.basic_consume(queue=self.sends, on_message_callback=self._callback)
        self.channel.start_consuming()
        
    def start(self):
        self._recieve()
        
class WorkerFactory():
    @staticmethod
    def chooseWorker(job:str, files, id:str):
        match job:
            case 'format':
                return FormatWorker(files, id)
            case 'merge':
                return MergeWorker(files, id)
            case _:
                return ErrorWorker(files, id)
            
class Worker(ABC):
    def __init__(self, files, id):
        self.files = files
        self.id = id
        self.error = str()
        self.message = str()
        
    @abstractmethod
    def format(self) -> (str, str):
        pass
        
class MergeWorker(Worker):
    
    def _merge(self, path_bitrix, path_web, path_done):
        # Объединяем два файла        
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
        
    
    def format(self):
        # Получаем пути и файлы
        path_web = self.files['file_web']
        path_bitrix = self.files['file_bitrix']
        path_done = self.files['file_done']
        
        return dumps({"id": self.id, "error": self.error, "message": self.message})
        
        
        
        
        
        
class FormatWorker(Worker):
    pass

class ErrorWorker(Worker):
    def format(self):
        error = "Ошибка"
        response = "Ошибка при работе над отчетом"
        return error, response






if __name__ == '__main__':
    c = Consumer('df')
    c.start()
    
    