import logging
from logging.handlers import TimedRotatingFileHandler
from datetime import datetime
import os

class Logger:
    def __init__(self, name='TeamViewerPyLog', log_dir='Pylog', level=logging.DEBUG):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        now = datetime.now()
        log_dir = os.path.join(script_dir, log_dir, now.strftime('%Y-%m'))
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        log_file = os.path.join(log_dir, now.strftime('%Y-%m-%d') + '.log')

        self.logger = logging.getLogger(name)
        self.logger.setLevel(level)
        handler = TimedRotatingFileHandler(log_file, when='midnight', interval=1, backupCount=30, encoding='utf-8')
        handler.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s', '%Y-%m-%d %H:%M:%S'))
        self.logger.addHandler(handler)

    def get_logger(self):
        return self.logger

    @staticmethod
    def record_time():
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")