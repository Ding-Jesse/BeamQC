from concurrent.futures import thread
from waitress import serve
 
import app


if __name__ == '__main__':
    app.app.secret_key = b'_5#y2L"F4Q8z\n\xda]/'
    serve(app.app, host='192.168.0.143', port=8080,threads = 8)
