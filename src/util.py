import os
import sys


def getCurRootPath():
    if getattr(sys, 'frozen', False):
        # rodando executavel
        curPath = os.path.abspath(sys.executable)
    else:
        # rodando py
        curPath = os.path.abspath(__file__)
    return os.path.dirname(curPath)
