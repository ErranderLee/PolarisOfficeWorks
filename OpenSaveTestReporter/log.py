import logging

class loggingman():
    def __init__(self):
        self.logger = logging.getLogger()
        self.logger.setLevel(logging.ERROR)
        self.formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        self.file_handler = logging.FileHandler('maillog.log')
        self.file_handler.setFormatter(self.formatter)
        self.logger.addHandler(self.file_handler)

    def write(self, msg):
        self.logger.error(msg)
