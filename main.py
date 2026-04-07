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
- dotenv.load_dot00env
- servicios.resolver_rutas.resource_path
- os.getenv
- modelo.procesar_archivo.ProcesarArchivo
"""

from controlador.scraping_spotfire import ScraperController
from dotenv import load_dotenv
from servicios.resolver_rutas import resource_path
from os import getenv
from modelo.procesar_archivo import ProcesarArchivo

import time
from dotenv import load_dotenv
from servicios.resolver_rutas import resource_path
from controlador.scraping_spotfire import ScraperController
from vista.logger import Logger


MAX_INTENTOS = 2  # Número máximo de reintentos
WAIT_REINTENTO = 30  # Segundos de espera entre intentos


def main() -> None:
    """
    Función principal que ejecuta la descarga y procesamiento del reporte DDA.
    """

    # Cargar variables de entorno
    load_dotenv(resource_path(".env"))

    scraper = ScraperController()
    log = Logger()
    for intento in range(1, MAX_INTENTOS + 1):
        log.log(f"Intento {intento} de {MAX_INTENTOS}...")

        if scraper.descargar_reporte_dda():
            scraper.procesar_insertar_data()
            log.log("Reporte descargado y procesado exitosamente.\n")
            return

        log.error("Fallo en la descarga. Reintentando...")
        time.sleep(WAIT_REINTENTO)

    log.error("No fue posible descargar el reporte tras varios intentos.")


if __name__ == "__main__":
    main()
