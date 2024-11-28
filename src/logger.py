# logger.py
import logging
import threading
from logging.handlers import RotatingFileHandler
from flask import session

# Thread-local storage
thread_local = threading.local()


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


def get_thread_logger(client_id):
    """
    Get or create a thread-specific logger for the current thread.
    """
    if not hasattr(thread_local, 'main_logger'):
        thread_local.main_logger = setup_custom_logger(__name__, client_id)
    return thread_local.main_logger
