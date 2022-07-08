import logging
import sys
import os
from __path__ import EXE_PATH


def setup_logger(log_file: str, stream=sys.stdout, file_mode: str='a+'):
    formatter = logging.Formatter(
        fmt='[{asctime}] [{name}] {levelname}: {message}',
        datefmt='%m/%d/%Y %H:%M:%S',
        style='{'
    )
    logger = logging.getLogger()
    logger.setLevel(logging.NOTSET)

    stream_handler = logging.StreamHandler(stream=stream)
    stream_handler.setLevel(logging.INFO)
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)

    file_handler = logging.FileHandler(log_file, mode=file_mode, encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    return logger


LOGGER = setup_logger(os.path.join(EXE_PATH, 'log.log'), file_mode='w')

