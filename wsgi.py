from concurrent.futures import thread
from waitress import serve
from werkzeug.middleware.proxy_fix import ProxyFix
import app


if __name__ == '__main__':
    app.app.config.from_object('config.config.ProdConfig')
    # app.app.secret_key = b'_5#y2L"F4Q8z\n\xda]/'
    app.app.wsgi_app = ProxyFix(
        app.app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_prefix=1
    )
    serve(app.app, host='192.168.0.143', port=8080,threads = 8)
