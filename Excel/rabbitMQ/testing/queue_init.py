import pika

class Queue_Init():
    def __init__(self):
        credentials = pika.PlainCredentials('myuser', 'mypassword')
        self.connection = pika.BlockingConnection(
            pika.ConnectionParameters(
                host='localhost',
                port=5000,
                virtual_host='myvhost',  # Check if this exists
                credentials=credentials
            )
        )
        self.channel = self.connection.channel()
        self.channel.queue_declare(queue='sends')
        self.channel.queue_declare(queue='responses')