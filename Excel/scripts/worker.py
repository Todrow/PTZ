import pika
from abc import ABC, abstractmethod

class Consumer():
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
            self.channel.queue_declare(queue='merge_sends')
        
        def callback(self, ch, method, properties, body):
            # code, file = Employer.choose_worker(body)
            return code, file
        
        
        def recieve(self):
            self.channel.basic_consume(on_message_callback=self.callback, 
                                    queue='queue',
                                    auto_ack=True)
            
            self.channel.start_consuming()

class Worker():
    @abstractmethod
    @staticmethod
    def do_work():
        code = str()
        file = str()
        pass
        return code, file

class Employer():
    @staticmethod
    def choose_worker(body):
        job = str() # что-то из body достаем
        match job:
            case 'format':
                return FormatWorker.do_work(body)
            case 'merge':
                return MergeWorker.do_work(body)
            case _:
                return None



class MergeWorker(Worker):
    pass

class FormatWorker(Worker):
    pass

if __name__ == "__main__":
    Consumer.recieve()