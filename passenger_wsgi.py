import sys
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

HOME = os.environ.get("HOME")
INTERP = os.path.join(HOME, "virtualenv", "bin", "python") if HOME else None
if INTERP and os.path.exists(INTERP):
    sys.executable = INTERP

if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from app import app as flask_app

def application(environ, start_response):
    path_info = environ.get("PATH_INFO") or ""

    if path_info == "/SS" or path_info.startswith("/SS/"):
        environ = environ.copy()
        environ["SCRIPT_NAME"] = "/SS"
        environ["PATH_INFO"] = path_info[3:] or "/"

    return flask_app(environ, start_response)
