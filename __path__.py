import sys

import os

if getattr(sys, 'frozen', False):
    RUN_PATH = sys._MEIPASS
    EXE_PATH = os.path.dirname(sys.executable)
else:
    EXE_PATH = RUN_PATH = os.path.dirname(__file__)


def resource(path, replace=False):
    path = os.path.join(RUN_PATH, path)
    return path.replace('\\', '/') if replace else path
