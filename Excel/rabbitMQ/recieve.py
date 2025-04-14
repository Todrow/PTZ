import pika

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
        
    def start_consuming(self):
        self._recieve()
  
if __name__ == '__main__':
    c = Consumer('df')
    c.start_consuming()
    
    