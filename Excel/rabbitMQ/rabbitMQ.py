import pika

class RabitQueue():
    # Инициализация подключения и канала
    def __init__(self):
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

    # Создание очереди
    def create_queue(self, queue):
        self.channel.queue_declare(queue=queue)
        
    # Отправка сообщений в очередь
    def send(self, queue, message):
        self.channel.basic_publish(exchange='', 
                                  routing_key=queue, 
                                  body=message)
        
        print("[*] message send to queue")
    
    
    # при получении сообщения
    def callback(self,ch,method,properties,body):
        print ("[x] Received %r" % (body,))
        
        
    def recieve(self):    
        self.channel.basic_consume(on_message_callback=self.callback, 
                                  queue='queue1',
                                  auto_ack=True)
        
        self.channel.start_consuming()
        
class Producer():
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
        self.channel.basic_consume(queue=self.responses, on_message_callback=self.handle_response)
        
        
    def send(self, message):
        self.channel.basic_publish(exchange='',
                                   routing_key=self.sends,
                                   body=message)
        
        print("сообщение отправлено в очередь для обработчика")
        
    def handle_response(self,ch,method,properties,body):
        print(f'получил отвевт {body}')
        
    def wait_for_response(self):
        self.channel.start_consuming()
        
        
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
        
    def callback(self,ch,method,properties,body):
        print('Обработчик получил запрос')
        print('отправляю ответ')
        self.channel.basic_publish(exchange='',
                                   routing_key=f"{self.name}_responses",
                                   body=f"ответ от обработчика {self.name}: {body}")
    def recieve(self):
        self.channel.basic_consume(queue=self.sends, on_message_callback=self.callback)
        self.channel.start_consuming()
        
    
        
        
                                   
        
        
