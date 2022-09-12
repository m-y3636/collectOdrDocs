import logging

class ExpLog():
    def __init__(self, log_path, log_format):
        self.log_path = log_path
        self.log_format = log_format

        self.log = logging.getLogger(__name__)
        log_formatter = logging.Formatter(
            log_format
        )
        # '%(asctime)s <%(levelname)s> : %(message)s'
        filehandler = logging.FileHandler(log_path, mode='a')
        filehandler.setFormatter(log_formatter)
        self.log.setLevel(logging.DEBUG)
        self.log.addHandler(filehandler)

    def logf_exp(self, *args, **kwargs):
        message = ','.join(args)
        self.log.debug(message)