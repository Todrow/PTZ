import pika
from json import dumps


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
            
    def _wait_for_response(self):
        self.channel.start_consuming()
        
    def send(self, message, id):
        self._send(message, id)
        self._wait_for_response()
        
        
if __name__ == '__main__':
    message = {
        'job': 'format',
        'files': {'file_w': 'file',
                'file_b': 'file'}
    }
    p = Producer('df')
    p.send(message=dumps(message), id=1)
    