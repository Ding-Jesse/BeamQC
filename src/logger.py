# logger.py
import logging
from logging.handlers import RotatingFileHandler
from flask import session


class ContextualFilter(logging.Filter):
    def filter(self, record):
        record.user_id = session.get('client_id')
        # record.user_id = getattr(local, 'user_id', 'unknown-user')
        return True


def setup_custom_logger(name, client_id):
    formatter = logging.Formatter(
        fmt='%(asctime)s - %(levelname)s - %(module)s : %(message)s')
    handler = RotatingFileHandler(
        f'logs/app_{client_id}.log', maxBytes=10000000, backupCount=10)
    handler.setFormatter(formatter)

    logger = logging.getLogger(name)
    if not logger.handlers:
        logger.addHandler(handler)
    logger.setLevel(logging.DEBUG)

    # logger.addFilter(ContextualFilter())
    return logger
