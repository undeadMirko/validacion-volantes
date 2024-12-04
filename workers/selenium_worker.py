from PyQt5.QtCore import QThread, pyqtSignal
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from browsermobproxy import Server
import time

# Rutas necesarias para WebDriver y BrowserMob Proxy
webdriver_path = "C:\\webdriver\\msedgedriver.exe"
browsermob_path = "C:\\utils\\browsermob-proxy\\bin\\browsermob-proxy"

class SeleniumWorker(QThread):
    cookie_capturado = pyqtSignal(str)  # Señal para emitir el valor de la cookie capturada

    def __init__(self):
        super().__init__()
        self.driver = None
        self.proxy = None
        self.server = None

    def run(self):
        """Configuración y ejecución del navegador con BrowserMob Proxy."""
        try:
            # Iniciar el servidor de BrowserMob Proxy
            self.server = Server(browsermob_path)
            self.server.start()
            self.proxy = self.server.create_proxy()

            # Configurar el navegador Edge con el proxy
            edge_options = webdriver.EdgeOptions()
            edge_options.add_argument(f'--proxy-server={self.proxy.proxy}')
            edge_options.add_argument('--disable-gpu')
            service = Service(webdriver_path)
            self.driver = webdriver.Edge(service=service, options=edge_options)

            # Configurar BrowserMob Proxy para capturar tráfico HTTP
            self.proxy.new_har("volantes", options={"captureHeaders": True, "captureContent": True})

            # Navegar a la página de inicio de sesión
            self.driver.get("https://poliedrodist.comcel.com.co/LoginPoliedro/Login.aspx")

            # Esperar a que el usuario navegue a la URL objetivo
            target_url = "https://poliedrodist.comcel.com.co/Recaudo.PS/VolantesNIT/ResumenVolantes"
            self.esperar_en_url(target_url)

            # Monitorear solicitudes hasta capturar la cookie
            self.capturar_cookie()
        except Exception as e:
            print(f"Error en SeleniumWorker: {e}")
        finally:
            # Mantener el navegador abierto (según la preferencia del usuario)
            pass

    def esperar_en_url(self, url_objetivo):
        """Espera a que el usuario navegue a la URL objetivo."""
        while True:
            if self.driver.current_url == url_objetivo:
                print("En la página objetivo. Esperando captura de solicitud...")
                break
            time.sleep(2)  # Revisar cada 2 segundos

    def capturar_cookie(self):
        """Captura la cookie de la solicitud POST a la URL objetivo."""
        while True:
            try:
                har_data = self.proxy.har
                for entry in har_data["log"]["entries"]:
                    request = entry["request"]
                    if "ResumenVolantes" in request["url"] and request["method"] == "POST":
                        # Extraer la cookie de los encabezados de la solicitud
                        cookies = next(
                            (header["value"] for header in request["headers"] if header["name"].lower() == "cookie"),
                            None
                        )
                        if cookies:
                            self.cookie_capturado.emit(cookies)
                            print(f"Cookie capturada: {cookies}")
                            return
                time.sleep(2)  # Revisar cada 2 segundos
            except Exception as e:
                print(f"Error al capturar la cookie: {e}")

    def stop(self):
        """Cierra el servidor de proxy si existe, pero no cierra el navegador ni el driver."""
        try:
            if self.proxy:
                self.proxy.close()
            if self.server:
                self.server.stop()
        except Exception as e:
            print(f"Error al detener el proxy o el servidor: {e}")
