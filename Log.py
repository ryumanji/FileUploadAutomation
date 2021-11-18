from datetime import datetime

class Log():
    def __init__(self):
        self.INFO = '[LOG][INFO] '
        self.WARN = '[LOG][WARN] '
        self.ERROR = '[LOG][ERROR] '

    def log_info(self, message):
        print(self.INFO + self.get_time_now() + ' ' + message)

    def log_warn(self, message):
        print(self.WARN + self.get_time_now() + ' ' + message)

    def log_error(self, message):
        print(self.ERROR + self.get_time_now() + ' ' + message)

    def get_time_now(self):
        now = datetime.now()
        return now.strftime('%Y/%m/%d %H:%M:%S')
