"""
Módulo para automatizar la descarga de reportes desde Spotfire utilizando Selenium.

Este módulo contiene la clase ScraperController, que proporciona métodos para automatizar la descarga de reportes
desde Spotfire y procesar los datos descargados.

Clases:
-------
- ScraperController: Clase para automatizar la descarga de reportes desde Spotfire.

Dependencias:
-------------
- pyautogui
- selenium
- servicios.resolver_rutas.resource_path
- os.getenv
- vista.logger.Logger
- modelo.procesar_archivo.ProcesarArchivo
"""

import pyautogui
import pygetwindow as gw
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from servicios.resolver_rutas import resource_path
import time
from os import getenv
from vista.logger import Logger
from modelo.procesar_archivo import ProcesarArchivo

class ScraperController:
    """
    Clase para automatizar la descarga de reportes desde Spotfire.

    Atributos:
    ----------
    ruta_driver : str
        Ruta al ejecutable del driver de Chrome.
    url_sopotfire : str
        URL de Spotfire.
    anno_reporte : str
        XPath del año del reporte.
    ruta_descarga_reporte : str
        Ruta donde se descargará el reporte.
    obj_log : Logger
        Objeto para manejar el registro de logs.
    obj_procesar_archivo : ProcesarArchivo
        Objeto para procesar el archivo descargado.
    """
    def __init__(self):
        """
        Inicializa la clase ScraperController con las rutas y configuraciones necesarias.
        """
        self.ruta_driver = resource_path(getenv('RUTA_WEBDRIVER'))
        self.url_sopotfire = getenv('RUTA_SPOTFIRE')
        #Rutas de carpetas
        self.carp_reports = getenv('XPATH_CARPETA_REPORTS')
        self.carp_reports_published = getenv('XPATH_CARPETA_REPORTS_PUBLISHED')
        self.cartp_dda_activity_report = getenv('XPATH_CARPETA_DDA_ACTIVITY_REPORT')

        self.anno_reporte = getenv('XPATH_YEAR_REPORTE')
        self.__ruta_descarga_reporte = resource_path(getenv('DESCARGA_REPORTE'))
        self.obj_log = Logger()
        self.obj_procesar_archivo = ProcesarArchivo()

    def descargar_reporte_dda(self):
        """
        Automatiza la descarga del reporte DDA desde Spotfire.
        
        Returns:
        --------
        bool
            True si la descarga es exitosa, False en caso contrario.
        """
        try:
            #Ruta al ejecutable de Chrome
            service = Service(executable_path = self.ruta_driver)

            #Cramos instancia de opciones
            options = webdriver.ChromeOptions()

            #indicamos que no se cierre el navegador a finalizar la ejecución
            options.add_experimental_option('detach'
                                            ,True)

            # Configurar las preferencias de descarga
            prefs = {
                'download.default_directory': self.__ruta_descarga_reporte,
                'download.prompt_for_download': False,
                'download.directory_upgrade': True,
                'safebrowsing.enabled': True
            }

            #Indicamos la ruta donde se descargará el archivo
            options.add_experimental_option('prefs', prefs)

            # options.add_argument('--log-level=3')
            # options.add_argument('--disable-browser-switcher')

            nav_driver = webdriver.Chrome(service = service
                                        ,options = options)

            #Instacia de actionchains
            nav_driver.implicitly_wait(30)

            #Obtenemos y naevegamos a la url
            nav_driver.get(self.url_sopotfire)

            #Instancia actionchains
            accion = ActionChains(nav_driver)

            #clic en IDP-renault
            nav_driver.find_element(By.CLASS_NAME,'w-36.mb-4.tss-button.compact.ng-star-inserted').click()
            time.sleep(10)

            try:

                # Traer la ventana de Chrome al frente
                chrome_windows = [w for w in gw.getWindowsWithTitle("Chrome") if w.isActive == False]
                if chrome_windows:
                    chrome_windows[0].activate()

            except:
                self.obj_log.log('No hay ventana de dialogo')
            #Usar pyautogui para interactuar con el cuadro de diálogo del sistema
            pyautogui.press('down')  # Presionar la tecla de flecha abajo
            pyautogui.press('enter')  # Presionar la tecla Enter

            #Clic en iniciar cesiósn
            nav_driver.find_element(By.ID, 'loginButton2').click()

            #Encontrar el elemento que queremos hacer doble clic (carpeta DIA)
            elemento_car_dia = nav_driver.find_element(By.ID, '8ac0a341-c82d-418b-bd0c-7a00d1209502')
            #Realizar doble clic
            accion.double_click(elemento_car_dia).perform()

            #Encontrar carpeta ope
            elemento_car_ope = nav_driver.find_element(By.ID, '636d33be-0069-4348-a765-6fa5e7fc7f11')
            #Doble clic elemento
            accion.double_click(elemento_car_ope).perform()

            #Encontrar carpeta Reports
            elemento_car_reports = nav_driver.find_element(By.XPATH, self.carp_reports)
            accion.double_click(elemento_car_reports).perform()

            #Encontrar carpeta Report_published
            elemento_car_reports_published = nav_driver.find_element(By.XPATH, self.carp_reports_published)
            accion.double_click(elemento_car_reports_published).perform()


            #Encontrar dda activiti report
            elemento_dda_activity_reports = nav_driver.find_element(By.XPATH, self.cartp_dda_activity_report)
            accion.double_click(elemento_dda_activity_reports).perform()

            # Obtener los identificadores de las ventanas
            ventanas = nav_driver.window_handles

            # Cambiar a la nueva pestaña (la última en la lista)
            nav_driver.switch_to.window(ventanas[1])

            #Encontrar elemento delailed view
            elemento_detail_view =  WebDriverWait(nav_driver, 120).until(
            EC.presence_of_element_located((By.ID, '0b9249dba6e945eaa3838b1aad644c5c')))
            elemento_detail_view.click()

            ##seleccionar pivotage
            element_pivotage =   WebDriverWait(nav_driver, 120).until(
            EC.presence_of_element_located((By.ID, '002323c710d546bdb354a77e6024e438')))
            element_pivotage.click()

            # #Seleccionar ojo sf-element-dropdown-list-item
            nav_driver.find_element(By.CLASS_NAME, 'sf-element-dropdown-list-item').click()
            time.sleep(3)

            # #Abrir panel paises 
            WebDriverWait(nav_driver, 120).until(EC.presence_of_element_located((By.XPATH,'/html/body/div/div[2]/div/div[1]/div/div/div[1]/div/div[2]/div/div/div/div/div/div/div/div/div[3]/div[1]/div[1]/div[1]/button[1]'))).click()
            time.sleep(60)

            #Seleccionar CO
            WebDriverWait(nav_driver, 120).until(EC.presence_of_element_located((By.XPATH,'/html/body/div/div[2]/div/div[1]/div/div/div[1]/div/div[2]/div/div/div/div/div/div/div/div/div[3]/div[1]/div[1]/div[1]/div[3]/span/div[1]/div[2]/div[1]/div[1]/div/div/div[7]'))).click()
            time.sleep(3)

            #Bajar Scroll bar
            elemento_scroll = WebDriverWait(nav_driver, 120).until(EC.presence_of_element_located((By.XPATH,'/html/body/div/div[2]/div/div[1]/div/div/div[1]/div/div[2]/div/div/div/div/div/div/div/div/div[3]/div[1]/div[1]/div[2]/div[2]')))
            contador=0
            while contador<= 12:
                elemento_scroll.click()
                contador+=1
            time.sleep(3)

            #Bajar 1 clic scroll
            elemento_scroll.click()

            #Abrir panel año 
            WebDriverWait(nav_driver, 120).until(EC.presence_of_element_located((By.XPATH,'/html/body/div/div[2]/div/div[1]/div/div/div[1]/div/div[2]/div/div/div/div/div/div/div/div/div[3]/div[1]/div[1]/div[1]/button[8]'))).click()
            time.sleep(2)

            elemento_scroll.click()

            contador=0
            while contador<= 12:
                elemento_scroll.click()
                contador+=1  
            time.sleep(3)

            # seleccionar año
            #nav_driver.find_element(By.XPATH,'/html/body/div/div[2]/div/div[1]/div/div/div[1]/div/div[2]/div/div/div/div/div/div/div/div/div[3]/div[1]/div[1]/div[1]/div[10]/span/div[1]/div[2]/div[1]/div[1]/div/div/div[3]').click()
            nav_driver.find_element(By.XPATH,self.anno_reporte).click()

            #Clic derecho
            time.sleep(60)
            elemento_clic_derecho = WebDriverWait(nav_driver, 300).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[1]/div[2]/div/div[1]/div/div/div[1]/div/div[2]/div/div/div/div/div/div/div/div/div[1]/div[1]/div[1]/div/div[2]/div[1]/div[2]/div[1]/div/div[4]/div'))).click()
            accion.context_click(elemento_clic_derecho).perform()

            #Exportar
            WebDriverWait(nav_driver, 120).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[2]/div[1]/div[1]/div'))).click()

            #Exportar tabla (Descargar archivo)
            WebDriverWait(nav_driver, 120).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div[1]/div[2]/div'))).click()

            #Esperar a que se realice la     descarga
            time.sleep(120)

            nav_driver.quit()

            self.obj_log.log('Se descarga correctamente el archivo')
            return True

        except Exception as ex:
            nav_driver.quit()
            #Resgitro log
            self.obj_log.error(f'Error al descargar el archivo de spotfire {ex}\n')
            return False

    def procesar_insertar_data(self):
        """
        Valida y procesa el archivo descargado, insertando los datos en la base de datos.
        """
        if self.obj_procesar_archivo.validar_archivo():
            if not self.obj_procesar_archivo.archivo_vacio():
                try:
                    self.obj_procesar_archivo.leer_archivo()
                except Exception as ex:
                    self.obj_log.error(f'Error al leer el archivo: {ex}\n')
                else:
                    self.obj_procesar_archivo.procesar_archivo()
        else:
            self.obj_log.error('El archivo no existe\n')
