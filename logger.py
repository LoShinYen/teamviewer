import logging
from logging.handlers import TimedRotatingFileHandler
from datetime import datetime
import os

class Logger:
    _instance = None  # 创建类属性以存储单例

    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super(Logger, cls).__new__(cls)
            cls._instance.setup_logger()
        return cls._instance

    def setup_logger(self):
        name = 'TeamViewerPyLog'
        log_dir = 'Pylog'
        level = logging.DEBUG
        script_dir = os.path.dirname(os.path.abspath(__file__))
        now = datetime.now()
        log_dir = os.path.join(script_dir, log_dir, now.strftime('%Y-%m'))
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        log_file = os.path.join(log_dir, now.strftime('%Y-%m-%d') + '.log')

        self.logger = logging.getLogger(name)
        self.logger.setLevel(level)
        if not self.logger.handlers:  
            handler = TimedRotatingFileHandler(log_file, when='midnight', interval=1, backupCount=30, encoding='utf-8')
            handler.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s', '%Y-%m-%d %H:%M:%S'))
            self.logger.addHandler(handler)

    def get_logger(self):
        return self.logger
