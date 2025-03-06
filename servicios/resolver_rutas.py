"""modulo para resolver rutas base y rutas relativas"""
import os
import sys

def resource_path(ruta_relativa):
    """resuelve las rutas"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path,ruta_relativa)
