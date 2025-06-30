import pika
from queue_init import Queue_Init


class Consumer(Queue_Init):
    
    def recieve(self, channel, method, properties, body):
        self.channel.basic_consume(queue='sends', on_message_callback=self._callback, auto_ack=True)
        self.channel.start_consuming()
        
    def _callback(self, ch, method, properties, body):
        self.channel.basic_publish(exchange='',routing_key='responses',body=body)
        
if __name__ == "__main__":
    consumer = Consumer()
    consumer.recieve()
    