import pika
from queue_init import Queue_Init

class Producer(Queue_Init):    
    def _send(self,msg):
        self.channel.basic_publish(exchange='',routing_key='sends',body=msg)
        self.channel.basic_consume(queue='responses', on_message_callback=self._callback, auto_ack=True)
        self.channel.start_consuming()
        
    def _callback(self,ch,method,properties,body):
        print("получен ответ, тело ответ - %s"%body)
        
if __name__ == '__main__':
    p = Producer()
    p._send