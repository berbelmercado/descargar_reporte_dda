"""
Módulo para manejar el registro de logs en la aplicación.

Este módulo contiene la clase Logger, que proporciona métodos para registrar mensajes de información y errores
en un archivo de log.

Clases:
-------
- Logger: Clase para manejar el registro de logs.

Dependencias:
-------------
- logging
- os
- servicios.resolver_rutas.resource_path
"""
import logging
import os
from servicios.resolver_rutas import resource_path

class Logger:
    """
    Clase para manejar el registro de logs en la aplicación.

    Atributos:
    ----------
    Ninguno

    Métodos:
    --------
    __init__():
        Inicializa la configuración del logger.
    log(mensaje: str):
        Registra un mensaje de información en el archivo de log.
    error(mensaje: str):
        Registra un mensaje de error en el archivo de log.
    """
    def __init__(self):
        """
        Inicializa la configuración del logger.
        """
        logging.basicConfig(filename=resource_path(os.getenv('ARCHIVO_LOG'))
                        ,level = logging.INFO
                        ,format = '%(asctime)s - %(levelname)s %(message)s'
                        )
    def log(self,mensaje):
        """
        Registra un mensaje de información en el archivo de log.

        Parámetros:
        -----------
        mensaje : str
            El mensaje de información a registrar.
        """
        logging.info(mensaje)
    
    def error(self,mensaje):
        """
        Registra un mensaje de error en el archivo de log.

        Parámetros:
        -----------
        mensaje : str
            El mensaje de error a registrar.
        """
        logging.error(mensaje)