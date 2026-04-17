import sys
import os

# Add your app directory to the path
INTERP = os.path.join(os.environ.get('HOME'), 'virtualenv', 'bin', 'python')
if os.path.exists(INTERP):
    sys.executable = INTERP

sys.path.insert(0, os.path.dirname(__file__))

from app import app as application

# Override the script name to handle the /SS subpath
def application(environ, start_response):
    # Remove the /SS prefix for Flask to handle
    path_info = environ.get('PATH_INFO', '')
    if path_info.startswith('/SS'):
        environ['PATH_INFO'] = path_info[3:]  # Remove /SS
        environ['SCRIPT_NAME'] = '/SS'
    return app(environ, start_response)
    return app(environ, start_response)
