from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QLabel, QLineEdit, QPushButton, QWidget, QFormLayout, QDateEdit, QComboBox, QHBoxLayout
from PyQt5.QtCore import QThread, pyqtSignal, QDate
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from browsermobproxy import Server
import sys
import time

# Ruta al Edge WebDriver
webdriver_path = "C:\\webdriver\\msedgedriver.exe"  # Cambia esta ruta si es necesario
browsermob_path = "C:\\utils\\browsermob-proxy\\bin\\browsermob-proxy"  # Cambia esta ruta si es necesario

# Selenium Worker en un hilo separado
class SeleniumWorker(QThread):
    distribuidor_obtenido = pyqtSignal(str)  # Señal para enviar el distribuidor obtenido
    cookie_obtenida = pyqtSignal(str)  # Señal para enviar la cookie obtenida

    def __init__(self):
        super().__init__()
        self.driver = None
        self.proxy = None
        self.server = None

    def run(self):
        # Configurar BrowserMob Proxy
        self.server = Server(browsermob_path)
        self.server.start()
        self.proxy = self.server.create_proxy()

        # Configurar el servicio y el navegador
        edge_options = webdriver.EdgeOptions()
        edge_options.add_argument(f'--proxy-server={self.proxy.proxy}')
        edge_options.add_argument('--disable-gpu')  # Opcional
        service = Service(webdriver_path)
        self.driver = webdriver.Edge(service=service, options=edge_options)

        # Iniciar el proxy para capturar las peticiones
        self.proxy.new_har("volantes", options={'captureHeaders': True, 'captureContent': True})

        # Navegar a la página de inicio de sesión
        self.driver.get("https://poliedrodist.comcel.com.co/LoginPoliedro/Login.aspx")

        # Esperar hasta estar en la página objetivo
        target_url = "https://poliedrodist.comcel.com.co/Recaudo.PS/VolantesNIT/"
        try:
            while True:
                if self.driver.current_url.startswith(target_url):
                    # Capturar distribuidor
                    try:
                        distribuidor_element = self.driver.find_element(By.ID, "userDataCodDistribuidor")
                        distribuidor = distribuidor_element.text
                    except Exception as e:
                        distribuidor = None
                        print(f"Error al obtener distribuidor: {e}")

                    # Capturar cookie en las solicitudes
                    time.sleep(3)  # Dar tiempo para capturar solicitudes
                    har_data = self.proxy.har
                    cookie_header = None
                    for entry in har_data['log']['entries']:
                        request_url = entry['request']['url']
                        headers = entry['request']['headers']

                        if "Recaudo.PS/VolantesNIT/Index" in request_url:
                            for header in headers:
                                if header['name'].lower() == "cookie":
                                    cookie_header = header['value']
                                    break
                            if cookie_header:
                                break

                    # Emitir las señales si los datos se obtuvieron
                    if distribuidor:
                        self.distribuidor_obtenido.emit(distribuidor)
                    if cookie_header:
                        self.cookie_obtenida.emit(cookie_header)

                    # Cerrar el navegador y salir
                    self.stop()
                    break
                time.sleep(2)
        except Exception as e:
            print(f"Error en SeleniumWorker: {e}")

    def stop(self):
        """Cierra el navegador, el proxy y el servidor."""
        if self.proxy:
            self.proxy.close()
        if self.server:
            self.server.stop()
        if self.driver:
            self.driver.quit()

    def cambiar_distribuidor(self):
        """Abre una nueva ventana o pestaña con el enlace para cambiar de distribuidor y espera que el usuario cierre la pestaña."""
        try:
            # Si no hay un navegador abierto o el anterior ya está cerrado, crear uno nuevo
            if self.driver is None or not self.driver.service.is_connectable():
                print("No hay navegador abierto o la sesión anterior está cerrada, creando una nueva instancia.")
                
                # Configurar nuevamente el servicio y el navegador
                edge_options = webdriver.EdgeOptions()
                edge_options.add_argument('--disable-gpu')
                service = Service(webdriver_path)
                self.driver = webdriver.Edge(service=service, options=edge_options)

            # Abrir una nueva pestaña en el navegador
            self.driver.execute_script(
                "window.open('https://convergencia.claro.com.co/sigma/app/index?v=1732816094428&next=%2Fsigma%2Fapp%2Findex%2Fmodules%2F8977#/login', '_blank');")
            
            window_handles = self.driver.window_handles
            self.driver.switch_to.window(window_handles[1])

            target_url = "https://convergencia.claro.com.co/sigma/app/index?v=1732816094428#/modules/8977/9077/form"
            WebDriverWait(self.driver, 30).until(EC.url_contains(target_url))

            print("Página cargada correctamente en la nueva pestaña. Esperando que el usuario cierre la pestaña.")

            # Esperar hasta que el usuario cierre la ventana
            while len(self.driver.window_handles) > 1:
                time.sleep(1)  # Comprobar cada segundo si la ventana sigue abierta

            # Una vez la ventana se ha cerrado, vuelve a la ventana principal
            self.driver.switch_to.window(window_handles[0])

            print("El usuario ha cerrado la pestaña. Volviendo a la ventana principal.")

        except Exception as e:
            print(f"Error al cambiar distribuidor: {e}")


# Interfaz gráfica
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Validación de Volantes")
        self.setGeometry(100, 100, 1024, 768)

        main_layout = QVBoxLayout()
        title_label = QLabel("Bienvenido a la validación de volantes")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; margin-bottom: 20px;")
        main_layout.addWidget(title_label)

        # Botones principales
        button_layout = QHBoxLayout()
        self.iniciar_sesion_button = QPushButton("Entrar en Poliedro")
        self.iniciar_sesion_button.setStyleSheet("background-color: #4CAF50; color: white; font-size: 16px; padding: 10px;")
        self.iniciar_sesion_button.clicked.connect(self.iniciar_sesion)

        self.cambiar_distribuidor_button = QPushButton("Cambiar Distribuidor")
        self.cambiar_distribuidor_button.setStyleSheet("background-color: #FF5733; color: white; font-size: 16px; padding: 10px;")
        self.cambiar_distribuidor_button.clicked.connect(self.cambiar_distribuidor)

        button_layout.addWidget(self.iniciar_sesion_button)
        button_layout.addWidget(self.cambiar_distribuidor_button)
        main_layout.addLayout(button_layout)

        # Formulario de campos
        form_layout = QFormLayout()

        # Campos de texto y selección
        self.distribuidor_label = QLabel("Distribuidor:")
        self.distribuidor_field = QLineEdit()
        self.distribuidor_field.setReadOnly(True)
        self.tipo_label = QLabel("Tipo:")
        self.tipo_combobox = QComboBox()
        self.tipo_combobox.addItems(["Voz", "Reposición"])

        # Campos de fechas
        self.fecha_inicio_label = QLabel("Fecha inicio:")
        self.fecha_inicio_edit = QDateEdit()
        self.fecha_inicio_edit.setDate(QDate.currentDate())
        self.fecha_inicio_edit.setCalendarPopup(True)

        self.fecha_fin_label = QLabel("Fecha fin:")
        self.fecha_fin_edit = QDateEdit()
        self.fecha_fin_edit.setDate(QDate.currentDate())
        self.fecha_fin_edit.setCalendarPopup(True)

        # Botones de acción
        self.consulta_button = QPushButton("Hacer Consulta")
        self.exportar_excel_button = QPushButton("Exportar a Excel")

        for button in [self.consulta_button, self.exportar_excel_button]:
            button.setEnabled(False)
            button.setStyleSheet("background-color: #008CBA; color: white; font-size: 14px; padding: 10px;")

        # Agregar widgets al formulario
        form_layout.addRow(self.distribuidor_label, self.distribuidor_field)
        form_layout.addRow(self.tipo_label, self.tipo_combobox)
        form_layout.addRow(self.fecha_inicio_label, self.fecha_inicio_edit)
        form_layout.addRow(self.fecha_fin_label, self.fecha_fin_edit)
        form_layout.addRow(self.consulta_button, self.exportar_excel_button)

        form_widget = QWidget()
        form_widget.setLayout(form_layout)
        main_layout.addWidget(form_widget)

        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

        self.selenium_worker = SeleniumWorker()
        self.selenium_worker.distribuidor_obtenido.connect(self.mostrar_distribuidor)

    def mostrar_distribuidor(self, distribuidor):
        self.distribuidor_field.setText(distribuidor)
        self.habilitar_formulario()

    def habilitar_formulario(self):
        self.consulta_button.setEnabled(True)
        self.exportar_excel_button.setEnabled(True)

    def iniciar_sesion(self):
        self.selenium_worker.start()

    def cambiar_distribuidor(self):
        self.selenium_worker.cambiar_distribuidor()

    def closeEvent(self, event):
        self.selenium_worker.stop()
        self.selenium_worker.wait()
        super().closeEvent(event)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
