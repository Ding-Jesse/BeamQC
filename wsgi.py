from waitress import serve
 
from app import app
 
serve(app, host='192.168.0.189', port=5002)
