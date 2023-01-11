from os import environ, path
from datetime import timedelta
class Config:
    # TESTING = True
    # permanent_session_lifetime = timedelta(minutes=60)
    SESSION_PERMANENT = True
    SESSION_TYPE = "filesystem"
    PERMANENT_SESSION_LIFETIME = timedelta(minutes=60)
    SESSION_USE_SIGNER = True
    PROGRESS_FILE = './TEST/OUTPUT'
    MAX_CONTENT_LENGTH = 100 * 1024 * 1024  # 100MB
    MAIL_SERVER='smtp.gmail.com'
    MAIL_PORT=465
    MAIL_USE_SSL=True
    MAIL_DEFAULT_SENDER=('freerccheck', 'elements.users27@gmail.com')
    MAIL_MAX_EMAILS=10
    MAIL_USERNAME='elements.users27@gmail.com'
    MAIL_PASSWORD='vzxwtbnbssfnfukq'
    MAIL_ASCII_ATTACHMENTS  =True
    DEBUG = False
    TESTING = False
    # GEVENT_SUPPORT =True
    
class DevConfig(Config):
    UPLOAD_FOLDER = 'D:/Desktop/BeamQC/TEST/INPUT'
    OUTPUT_FOLDER = 'D:/Desktop/BeamQC/TEST/OUTPUT'
    PROGRESS_FILE = 'D:/Desktop/BeamQC/TEST/OUTPUT'
    # FLASK_ENV = 'development'
    SECRET_KEY = 'dev2'
    DEBUG = True
    TESTING = True
    

class ProdConfig(Config):
    UPLOAD_FOLDER = 'C:/Users/User/Desktop/BeamQC/INPUT'
    OUTPUT_FOLDER = 'C:/Users/User/Desktop/BeamQC/OUTPUT'
    PROGRESS_FILE = 'C:/Users/User/Desktop/BeamQC/OUTPUT'
    FLASK_ENV = 'production'
    SECRET_KEY = 'bbdb12eeb63aeb29a9535999e091b5f6de228d9e099575f92f29c10cc0a13c06'
    DEBUG = False
    TESTING = False