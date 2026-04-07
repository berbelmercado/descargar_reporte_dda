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

import time
from os import getenv

import pyautogui
import pygetwindow as gw
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

from servicios.resolver_rutas import resource_path
from vista.logger import Logger
from modelo.procesar_archivo import ProcesarArchivo


# ---------------------------------------------------------------------------
# Tiempos de espera (segundos) – centralizar facilita ajustarlos sin tocar
# la lógica de negocio.
# ---------------------------------------------------------------------------
_WAIT_IMPLICIT = 30  # espera implícita global del driver
_WAIT_SHORT = 3  # pausa breve tras interacciones de UI
_WAIT_LOGIN = 10  # tiempo para que cargue el IDP
_WAIT_REPORT = 40  # carga inicial del reporte
_WAIT_COUNTRY = 60  # actualización del filtro de países
_WAIT_YEAR = 60  # actualización del filtro de año
_WAIT_DOWNLOAD = 120  # tiempo de descarga del archivo
_WAIT_EXPLICIT = 120  # timeout estándar para WebDriverWait
_WAIT_LONG = 300  # timeout para elementos de carga lenta
_SCROLL_CLICKS = 12  # cantidad de clics en la barra de desplazamiento


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
        self.ruta_driver = resource_path(getenv("RUTA_WEBDRIVER"))
        self.url_sopotfire = getenv("RUTA_SPOTFIRE")

        # Inicio sesión Spotfire
        self.input_ipn = getenv("ID_INPUT_IPN")
        self.selec_clave = getenv("XPATH_INGRESAR_CLAVE")
        self.input_clave = getenv("ID_INPUT_CLAVE")
        self.siguiente = getenv("XPATH_SIGUIENTE")
        self.verificar = getenv("XPATH_VERIFICAR")

        # Rutas de carpetas
        self.carp_dia2 = getenv("XPATH_CARPETA_DIA2")
        self.carp_reports = getenv("XPATH_CARPETA_REPORTS")
        self.carp_reports_published = getenv("XPATH_CARPETA_REPORTS_PUBLISHED")
        self.cartp_dda_activity_report = getenv("XPATH_CARPETA_DDA_ACTIVITY_REPORT")

        self.anno_reporte = getenv("XPATH_YEAR_REPORTE")
        self.__ruta_descarga_reporte = resource_path(getenv("DESCARGA_REPORTE"))

        # Rutas accesos
        self.ipn = getenv("IPN")
        self.clave = getenv("CLAVE")

        # Certificados en servidor
        self.validar_certificado = getenv("SSL_VERIFY")

        self.obj_log = Logger()
        self.obj_procesar_archivo = ProcesarArchivo()

    # ------------------------------------------------------------------
    # Helpers privados
    # ------------------------------------------------------------------

    def _crear_driver(self) -> webdriver.Chrome:
        """
        Crea y configura una instancia de ChromeDriver con las preferencias
        de descarga necesarias.

        Returns:
        --------
        webdriver.Chrome
            Instancia configurada del navegador.
        """
        dict_validacion_driver = self.validar_driver()
        service = (
            dict_validacion_driver["driver"]
            if dict_validacion_driver["estado"]
            else Service(executable_path=self.ruta_driver)
        )

        options = webdriver.ChromeOptions()

        # Indicamos que no se cierre el navegador al finalizar la ejecución
        options.add_experimental_option("detach", True)

        # Configurar las preferencias de descarga
        prefs = {
            "download.default_directory": self.__ruta_descarga_reporte,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
        }
        options.add_experimental_option("prefs", prefs)

        # options.add_argument('--log-level=3')
        # options.add_argument('--disable-browser-switcher')

        driver = webdriver.Chrome(service=service, options=options)
        driver.implicitly_wait(_WAIT_IMPLICIT)
        return driver

    def _aceptar_sitio_no_seguro(self, driver: webdriver.Chrome) -> None:
        """
        Maneja la advertencia de seguridad del navegador cuando el sitio
        no tiene un certificado de confianza.
        """
        if self.validar_exitencia(driver, "details-button", "ID"):
            WebDriverWait(driver, _WAIT_IMPLICIT).until(
                EC.presence_of_element_located((By.ID, "details-button"))
            ).click()

            # Accedemos a la URL a pesar de la advertencia de seguridad
            WebDriverWait(driver, _WAIT_IMPLICIT).until(
                EC.presence_of_element_located((By.ID, "proceed-link"))
            ).click()

    def _validar_certificado_servidor(self, driver: webdriver.Chrome) -> None:
        """
        Interactúa con el cuadro de diálogo del sistema operativo para
        aceptar el certificado del servidor cuando SSL_VERIFY es True.
        """
        if self.validar_certificado != "True":
            return

        # Traer la ventana de Chrome al frente
        chrome_windows = [w for w in gw.getWindowsWithTitle("Chrome") if not w.isActive]
        if chrome_windows:
            chrome_windows[0].activate()

        # Usar pyautogui para interactuar con el cuadro de diálogo del sistema
        pyautogui.press("down")  # Presionar la tecla de flecha abajo
        pyautogui.press("enter")  # Presionar la tecla Enter
        self.obj_log.log("Valida certificado del servidor")

    def _navegar_carpetas(self, driver: webdriver.Chrome, accion: ActionChains) -> None:
        """
        Navega por la estructura de carpetas de Spotfire hasta llegar
        al reporte DDA Activity Report.
        """
        # Encontrar el elemento que queremos hacer doble clic (carpeta DIA)
        elemento_car_dia = driver.find_element(
            By.ID, "8ac0a341-c82d-418b-bd0c-7a00d1209502"
        )
        accion.double_click(elemento_car_dia).perform()

        # Encontrar carpeta OPE
        elemento_car_ope = driver.find_element(
            By.ID, "636d33be-0069-4348-a765-6fa5e7fc7f11"
        )
        accion.double_click(elemento_car_ope).perform()
        # Encontrar carpeta dia2
        elemento_car_dia2 = driver.find_element(By.XPATH, self.carp_dia2)
        accion.double_click(elemento_car_dia2).perform()

        # Encontrar carpeta Reports
        elemento_car_reports = driver.find_element(By.XPATH, self.carp_reports)
        accion.double_click(elemento_car_reports).perform()

        # Encontrar carpeta Reports_published
        elemento_car_reports_published = driver.find_element(
            By.XPATH, self.carp_reports_published
        )
        accion.double_click(elemento_car_reports_published).perform()

        # Encontrar DDA Activity Report
        elemento_dda_activity_reports = driver.find_element(
            By.XPATH, self.cartp_dda_activity_report
        )
        accion.double_click(elemento_dda_activity_reports).perform()

    def _cambiar_a_nueva_ventana(self, driver: webdriver.Chrome) -> None:
        """
        Espera a que se abra una nueva pestaña y cambia el foco a ella.
        """
        ventana_original = driver.current_window_handle

        # Espera a que haya más de una ventana
        WebDriverWait(driver, _WAIT_SHORT * 3).until(
            lambda d: len(d.window_handles) > 1
        )

        # Cambia a la nueva ventana
        for handle in driver.window_handles:
            if handle != ventana_original:
                driver.switch_to.window(handle)
                break

    def _aplicar_filtros(self, driver: webdriver.Chrome, accion: ActionChains) -> None:
        """
        Aplica los filtros de vista, país y año dentro del reporte de Spotfire.
        """
        time.sleep(_WAIT_REPORT)

        # Encontrar elemento Detailed View
        # elemento_detail_view = WebDriverWait(driver, _WAIT_EXPLICIT).until(
        #     EC.element_to_be_clickable((By.ID, '0b9249dba6e945eaa3838b1aad644c5c')))
        # elemento_detail_view.click()

        elemento = WebDriverWait(driver, _WAIT_EXPLICIT).until(
            EC.presence_of_element_located((By.ID, "0b9249dba6e945eaa3838b1aad644c5c"))
        )
        driver.execute_script("arguments[0].scrollIntoView();", elemento)
        WebDriverWait(driver, _WAIT_EXPLICIT).until(
            EC.element_to_be_clickable((By.ID, "0b9249dba6e945eaa3838b1aad644c5c"))
        ).click()

        # Seleccionar pivotage
        element_pivotage = WebDriverWait(driver, _WAIT_EXPLICIT).until(
            EC.presence_of_element_located((By.ID, "002323c710d546bdb354a77e6024e438"))
        )
        element_pivotage.click()

        # Seleccionar ojo (sf-element-dropdown-list-item)
        driver.find_element(By.CLASS_NAME, "sf-element-dropdown-list-item").click()
        time.sleep(_WAIT_SHORT)

        # Abrir panel países
        WebDriverWait(driver, _WAIT_EXPLICIT).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "/html/body/div/div[2]/div/div[1]/div/div/div[1]/div/div[2]"
                    "/div/div/div/div/div/div/div/div/div[3]/div[1]/div[1]/div[1]/button[1]",
                )
            )
        ).click()
        time.sleep(_WAIT_COUNTRY)

        # Seleccionar CO
        WebDriverWait(driver, _WAIT_EXPLICIT).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "/html/body/div/div[2]/div/div[1]/div/div/div[1]/div/div[2]"
                    "/div/div/div/div/div/div/div/div/div[3]/div[1]/div[1]/div[1]"
                    "/div[3]/span/div[1]/div[2]/div[1]/div[1]/div/div/div[7]",
                )
            )
        ).click()
        time.sleep(_WAIT_SHORT)

        # Bajar Scroll bar
        elemento_scroll = WebDriverWait(driver, _WAIT_EXPLICIT).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "/html/body/div/div[2]/div/div[1]/div/div/div[1]/div/div[2]"
                    "/div/div/div/div/div/div/div/div/div[3]/div[1]/div[1]/div[2]/div[2]",
                )
            )
        )

        for _ in range(_SCROLL_CLICKS + 1):
            elemento_scroll.click()
        time.sleep(_WAIT_SHORT)

        # Bajar 1 clic scroll adicional
        elemento_scroll.click()

        # Abrir panel año
        WebDriverWait(driver, _WAIT_EXPLICIT).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "/html/body/div/div[2]/div/div[1]/div/div/div[1]/div/div[2]"
                    "/div/div/div/div/div/div/div/div/div[3]/div[1]/div[1]/div[1]/button[8]",
                )
            )
        ).click()
        time.sleep(_WAIT_SHORT - 1)

        elemento_scroll.click()
        for _ in range(_SCROLL_CLICKS + 1):
            elemento_scroll.click()
        time.sleep(_WAIT_SHORT)

        # Seleccionar año
        # driver.find_element(
        #     By.XPATH,
        #     '/html/body/div/div[2]/div/div[1]/div/div/div[1]/div/div[2]/div/div/div'
        #     '/div/div/div/div/div/div[3]/div[1]/div[1]/div[1]/div[10]/span/div[1]'
        #     '/div[2]/div[1]/div[1]/div/div/div[3]'
        # ).click()
        driver.find_element(By.XPATH, self.anno_reporte).click()

    def _exportar_tabla(self, driver: webdriver.Chrome, accion: ActionChains) -> None:
        """
        Realiza clic derecho sobre la tabla del reporte y ejecuta la
        secuencia de exportación para descargar el archivo.
        """
        time.sleep(_WAIT_YEAR)

        elemento_clic_derecho = (
            WebDriverWait(driver, _WAIT_LONG)
            .until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        "/html/body/div[1]/div[2]/div/div[1]/div/div/div[1]/div/div[2]"
                        "/div/div/div/div/div/div/div/div/div[1]/div[1]/div[1]/div/div[2]"
                        "/div[1]/div[2]/div[1]/div/div[4]/div",
                    )
                )
            )
            .click()  # .click() devuelve None; accion.context_click recibe None, lo que es correcto
        )
        accion.context_click(elemento_clic_derecho).perform()

        # Exportar (opción del menú contextual)
        WebDriverWait(driver, _WAIT_EXPLICIT).until(
            EC.presence_of_element_located(
                (By.XPATH, "/html/body/div[2]/div[1]/div[1]/div")
            )
        ).click()

        # Exportar tabla (descarga el archivo)
        WebDriverWait(driver, _WAIT_EXPLICIT).until(
            EC.presence_of_element_located(
                (By.XPATH, "/html/body/div[3]/div[1]/div[2]/div")
            )
        ).click()

        # Esperar a que se realice la descarga
        time.sleep(_WAIT_DOWNLOAD)

    # ------------------------------------------------------------------
    # API pública
    # ------------------------------------------------------------------

    def descargar_reporte_dda(self) -> bool:
        """
        Automatiza la descarga del reporte DDA desde Spotfire.

        Returns:
        --------
        bool
            True si la descarga es exitosa, False en caso contrario.
        """
        driver = None
        try:
            driver = self._crear_driver()
            accion = ActionChains(driver)

            # Navegar a la URL de Spotfire
            driver.get(self.url_sopotfire)

            # Validación cuando el sitio es no seguro
            self._aceptar_sitio_no_seguro(driver)

            # Clic en IDP-Renault
            driver.find_element(
                By.CLASS_NAME, "w-36.mb-4.tss-button.compact.ng-star-inserted"
            ).click()
            time.sleep(_WAIT_LOGIN)

            # Validar certificado del servidor si corresponde
            self._validar_certificado_servidor(driver)

            # Clic en iniciar sesión para token
            # driver.find_element(
            #     By.XPATH,
            #     "/html/body/div[2]/main/div[2]/div/div/div[2]/form/div[1]/a",
            # ).click()

            # Ingresamos la IPN
            input_ipn = driver.find_element(By.ID, self.input_ipn)
            input_ipn.send_keys(self.ipn)

            # Clic en siguiente
            driver.find_element(By.XPATH, self.siguiente).click()
            time.sleep(1)

            # Seleccionamos metodo de inicio con clave de acceso
            driver.find_element(By.XPATH, self.selec_clave).click()

            # capturamos input para ingresar la clave
            input_clave = driver.find_element(By.ID, self.input_clave)
            # Ingresamos la clave
            input_clave.send_keys(self.clave)
            time.sleep(3)

            # Clic en verificar
            driver.find_element(By.XPATH, self.verificar).click()
            time.sleep(3)

            # Navegar por la estructura de carpetas
            self._navegar_carpetas(driver, accion)

            # Cambiar al reporte que se abre en nueva pestaña
            self._cambiar_a_nueva_ventana(driver)

            # Aplicar filtros de país y año
            self._aplicar_filtros(driver, accion)

            # Exportar y descargar la tabla
            self._exportar_tabla(driver, accion)

            driver.quit()
            self.obj_log.log("Se descarga correctamente el archivo")
            return True

        except Exception as ex:
            if driver:
                driver.quit()
            self.obj_log.error(f"Error al descargar el archivo de Spotfire: {ex}\n")
            return False

    def procesar_insertar_data(self) -> None:
        """
        Valida y procesa el archivo descargado, insertando los datos en la base de datos.
        """
        if not self.obj_procesar_archivo.validar_archivo():
            self.obj_log.error("El archivo no existe\n")
            return

        if self.obj_procesar_archivo.archivo_vacio():
            return

        try:
            self.obj_procesar_archivo.leer_archivo()
        except Exception as ex:
            self.obj_log.error(f"Error al leer el archivo: {ex}\n")
            return

        self.obj_procesar_archivo.procesar_archivo()

    def validar_exitencia(
        self, nav_driver: webdriver.Chrome, identificador: str, tipo_id: str
    ) -> bool:
        """
        Verifica si un elemento existe en la página dentro de un breve timeout.

        Parameters:
        -----------
        nav_driver : webdriver.Chrome
            Instancia activa del navegador.
        identificador : str
            Valor del selector (ID, XPATH, etc.).
        tipo_id : str
            Tipo de selector como string (ej. "ID", "XPATH").

        Returns:
        --------
        bool
            True si el elemento existe, False en caso contrario.
        """
        try:
            by = getattr(By, tipo_id)
            WebDriverWait(nav_driver, _WAIT_SHORT).until(
                EC.presence_of_element_located((by, identificador))
            )
            return True
        except Exception:
            return False

    def validar_driver(self) -> dict:
        """
        Intenta obtener el ChromeDriver actualizado vía ChromeDriverManager.

        Returns:
        --------
        dict
            {"estado": True, "driver": service} si tiene éxito,
            {"estado": False} si falla y se debe usar el driver local.
        """
        try:
            service = Service(ChromeDriverManager().install())
            return {"estado": True, "driver": service}
        except Exception:
            return {"estado": False}
