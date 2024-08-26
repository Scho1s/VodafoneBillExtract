import logging
from datetime import datetime

class Logger:
    def __init__(self, name, filename='activity.log', level=logging.INFO):
        self.name = name
        self.logger = logging.getLogger(name)
        logging.basicConfig(filename=filename, level=level)

    def info(self, message):
        self.logger.info(f"{datetime.now()}: {message}")
