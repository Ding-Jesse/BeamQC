from waitress import serve
 
from app import app
 
serve(app, host='192.168.0.143', port=8080)
