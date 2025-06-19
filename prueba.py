from controlador.scraping_spotfire import ScraperController
from servicios.resolver_rutas import resource_path
from dotenv import load_dotenv
load_dotenv(resource_path('.env'))

obj_controler = ScraperController()
obj_controler.procesar_insertar_data()