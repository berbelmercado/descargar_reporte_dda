"""
Módulo principal para ejecutar la automatización de descarga y procesamiento de reportes desde Spotfire.

Este módulo carga las variables de entorno, crea un objeto ScraperController para manejar la descarga de reportes,
e intenta descargar y procesar el archivo hasta 3 veces antes de finalizar la ejecución.

Funciones:
----------
- main(): Función principal que ejecuta la automatización de descarga y procesamiento de reportes.

Dependencias:
-------------
- controlador.scraping_spotfire.ScraperController
- dotenv.load_dotenv
- servicios.resolver_rutas.resource_path
- os.getenv
- modelo.procesar_archivo.ProcesarArchivo
"""
from controlador.scraping_spotfire import ScraperController
from dotenv import load_dotenv
from servicios.resolver_rutas import resource_path
from os import getenv
from modelo.procesar_archivo import ProcesarArchivo

def main():
    """
    Función principal que ejecuta la automatización de descarga y procesamiento de reportes.

    Carga las variables de entorno, crea un objeto ScraperController para manejar la descarga de reportes,
    e intenta descargar y procesar el archivo hasta 3 veces antes de finalizar la ejecución.
    """
    #Cargar variables de entorno
    load_dotenv(resource_path('.env'))

    #Creamos objeto para scraping
    obj_scraping_controller = ScraperController()

    intento=0
    #Se intenta descargar el archivo 3 veces antes de finalizar ejecución
    while intento<=3:
        if obj_scraping_controller.descargar_reporte_dda():
            obj_scraping_controller.procesar_insertar_data()
            intento=4
        else:
            intento+=1
if __name__ == "__main__":
    main()
