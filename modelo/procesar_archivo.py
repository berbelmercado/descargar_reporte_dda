"""
Módulo para procesar archivos y manejar la inserción de datos en una base de datos SQL Server.

Este módulo contiene la clase ProcesarArchivo, que proporciona métodos para leer, procesar y eliminar archivos,
así como insertar datos filtrados en una base de datos SQL Server.

Clases:
-------
- ProcesarArchivo: Clase para procesar archivos y manejar la inserción de datos en una base de datos SQL Server.

Dependencias:
-------------
- pandas
- vista.logger.Logger
- servicios.resolver_rutas.resource_path
- os.getenv
- os.path
- os.remove
- modelo.inserta_datos.InsertaDatos
"""
import pandas as pd
from vista.logger import Logger
from servicios.resolver_rutas import resource_path
from os import getenv,path,remove
from modelo.inserta_datos import InsertaDatos

class ProcesarArchivo:
    """
    Clase para procesar archivos y manejar la inserción de datos en una base de datos SQL Server.

    Atributos:
    ----------
    archivo_reporte : str
        Ruta del archivo de reporte.
    obj_log : Logger
        Objeto para manejar el registro de logs.
    dataframe_filtrado : pandas.DataFrame
        DataFrame para almacenar los datos filtrados.
    obj_inserta_datos : InsertaDatos
        Objeto para manejar la inserción de datos en la base de datos.
    """
    def __init__(self):
        """
        Inicializa la clase ProcesarArchivo con la ruta del archivo de reporte y un objeto de registro de logs.
        """
        self.archivo_reporte = resource_path(getenv('ARCHIVO_REPORTE'))#leer ruta archivo
        self.obj_log = Logger() #Creamos objeto log
        self.dataframe_filtrado = pd.DataFrame()
        self.obj_inserta_datos = InsertaDatos()

    def leer_modificar_archivo(self):
        """
        Lee y procesa el archivo de reporte si no está vacío.
        """
        if not self._archivo_vacio():
            try:
                self._leer_archivo()
            except Exception as ex:
                self.obj_log.error(f'Error al leer el archivo: {ex}\n')
            else:
                self._procesar_archivo()

    def archivo_vacio(self):
        """
        Verifica si el archivo de reporte está vacío.

        Returns:
        --------
        bool
            True si el archivo está vacío, False en caso contrario.
        """
        if path.getsize(self.archivo_reporte) == 0:
            self.obj_log.error("Archivo vacío\n")
            return True
        return False

    def leer_archivo(self):
        """
        Lee el archivo de reporte y lo carga en un DataFrame.
        """
        self.dataframe = pd.read_excel(self.archivo_reporte)
        self.obj_log.log('Se lee correctamente el archivo')

    def procesar_archivo(self):
        """
        Filtra los datos del archivo y los inserta en la base de datos.
        """
        try:
            #Filtramos los datos del archivo y lo guardamos en un nuevo Dataframe
            self.dataframe_filtrado = self.dataframe.loc[(self.dataframe['Process'] == 'Delivery')
                                                            ,['VIN','Date']
                                                        ]
            #Conectamos a la base de datos
            if self.obj_inserta_datos.conexion_sql():
                #Insertamos los datos filtrados ala base de datos
                self.obj_inserta_datos.insertar_datos(self.dataframe_filtrado)

        except Exception as ex:
            #Registramos log
            self.obj_log.log(f'Error al procesar el archivo: {ex}')

        finally:
                #Eliminamos archivo
                self.eliminar_archivo()

    def eliminar_archivo(self):
        """
        Elimina el archivo de reporte.
        """
        try:
            remove(self.archivo_reporte)
            self.obj_log.log("El archivo ha sido eliminado\n")
        except Exception as err:
            #Registramos log
            self.obj_log.error(f'Error al eliminar el archivo {err}\n')

    def validar_archivo(self):
        """
        Verifica si el archivo de reporte existe.

        Returns:
        --------
        bool
            True si el archivo existe, False en caso contrario.
        """
        if path.exists(self.archivo_reporte):
            return True
        else:
            return False
