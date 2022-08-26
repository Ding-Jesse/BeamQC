from concurrent.futures import thread
from waitress import serve
from werkzeug.middleware.proxy_fix import ProxyFix
import app


if __name__ == '__main__':
    app.app.secret_key = b'_5#y2L"F4Q8z\n\xda]/'
    app.app.wsgi_app = ProxyFix(
        app.app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_prefix=1
    )
    serve(app.app, host='192.168.1.102', port=8081,threads = 8)
