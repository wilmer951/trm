import win32com.client
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time
import datetime
import pandas as pd
import requests
from pathlib import Path
import psutil  # Para verificar si Outlook está abierto

fecha_actual = datetime.datetime.now()

# Ruta de almacenamiento de captura
filecaptura = r'C:\trm\captura\captura.png'

file_destinatarios = r'C:\trm\correos\destinatarios.txt'
file_asunto = r'C:\trm\correos\asunto.txt'
file_cuerpo = r'C:\trm\correos\cuerpo.txt'
file_log = r'C:\trm\log\log.txt'
# Ruta de la página
url = 'https://www.banguat.gob.gt/tipo_cambio/'

# Función para obtener el nombre del archivo de log con la fecha (sin hora)
def obtener_nombre_log():
    fecha_actual = datetime.datetime.now().strftime("%Y-%m-%d") 
    return f"C:\\trm\\log\\log_{fecha_actual}.txt"

# Creación de archivo log
def log_result(msm):
    file_log = obtener_nombre_log()
    with open(file_log, 'a') as file:       
        file.write(f'{fecha_actual} {str(msm)} \n')
        file.write('************* \n')

# Función para verificar la conexión a la URL
def verificar_conexion(url):
    mensaje = r'inicio de proceso'
    log_result(mensaje)
    try:
        respuesta = requests.get(url, timeout=5)
        if respuesta.status_code == 200:
            mensaje = r'Conexión a página exitosa'
            log_result(mensaje)
            return True
        else:
            mensaje = r'Problemas de conexión'
            log_result(mensaje)
            return False
    except requests.exceptions.RequestException as e:
        mensaje_error = f"Error de conexión: {str(e)}"
        log_result(mensaje_error)
        return False

# Función para capturar la pantalla de la página web
def captura_pantalla_completa(url, archivo_salida):
    if verificar_conexion(url):  # Asegurarse de que la conexión es exitosa
        try:
            # Configurar opciones para el navegador
            opciones = Options()
            opciones.add_argument("--headless")  # Ejecutar en modo headless
            opciones.add_argument("--disable-gpu")  # Desactivar GPU (recomendado para headless)
            opciones.add_argument("--no-sandbox")  # Desactivar sandboxing (requerido en algunos entornos)
            opciones.add_argument("--remote-debugging-port=9222")  # Añadir puerto de depuración remota
            opciones.add_argument("--disable-dev-shm-usage")  # Evitar problemas con memoria compartida
            opciones.add_argument("--disable-software-rasterizer")  # Desactivar software rasterizer

            # Configurar el servicio del WebDriver con el gestor automático de drivers
            servicio = Service(ChromeDriverManager().install())

            # Inicializar el navegador
            navegador = webdriver.Chrome(service=servicio, options=opciones)

            # Navegar a la página web
            navegador.get(url)

            # Esperar un poco para asegurarnos de que la página se haya cargado completamente
            time.sleep(3)

            # Encuentra todos los botones "Consultar"
            botones_consultar = navegador.find_elements(By.XPATH, "//input[@type='submit' and @value='Consultar']")

            # Verifica si hay al menos tres botones "Consultar"
            if len(botones_consultar) >= 3:
                # Hacer clic en el tercer botón (índice 2, ya que los índices empiezan desde 0)
                botones_consultar[2].click()
            else:
                mensaje = "No se encontraron suficientes botones 'Consultar'."
                log_result(mensaje)
                print(mensaje)
                navegador.quit()
                return

            # Esperar un poco para asegurarnos de que la consulta se haya procesado y los resultados estén cargados
            time.sleep(5)

            # Hacer una captura de pantalla completa
            navegador.save_screenshot(archivo_salida)

            # Cerrar el navegador
            navegador.quit()
            status = 1
            mensaje = r'Captura con éxito'
            log_result(mensaje)

            send_email(status)

        except Exception as e:
            # Capturar cualquier excepción y registrar el error
            mensaje_error = f"Error durante el proceso: {str(e)}"
            log_result(mensaje_error)
            status = 0
            send_email(status)
    else:
        mensaje_error = f"Error de conexión"
        log_result(mensaje_error)

        status = 0
        send_email(status)

# Función para verificar si Outlook está abierto, y si no lo está, abrirlo
def iniciar_outlook_si_no_esta():
    is_outlook_open = False
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'].lower() == 'outlook.exe':
            is_outlook_open = True
            break
    
    if not is_outlook_open:
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            time.sleep(5)  # Esperar a que Outlook se inicie completamente
            log_result("Outlook ha sido abierto")
        except Exception as e:
            log_result(f"Error al abrir Outlook: {e}")
            return False
    return True

# Función para enviar el correo
def send_email(status):
    try:
        if status == 1:
            destinatarios = obtener_correos()
            asunto = obtener_asunto()
            cuerpo = obtener_cuerpo()

            # Verificar si Outlook está abierto
            if iniciar_outlook_si_no_esta():
                # Crear objeto Outlook
                outlook = win32com.client.Dispatch('Outlook.Application')
                mensaje = outlook.CreateItem(0)  # 0 = Correo

                # Configurar propiedades del mensaje
                mensaje.Subject = asunto
                mensaje.HTMLBody = cuerpo  # Usar HTMLBody para poder incluir la imagen

                # Adjuntar la imagen de forma incrustada en el cuerpo
                attachment = mensaje.Attachments.Add(filecaptura)

                # Configurar el Content-ID para la imagen
                attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "captura_image")

                # Usar Content-ID (CID) para incrustar la imagen en el HTML
                cid = "captura_image"  # El mismo que has usado antes
                imagen_incrustada = f'<img src="cid:{cid}">'

                # Insertar la imagen incrustada en el cuerpo del correo HTML
                # Verifica si el cuerpo contiene [imagen] para reemplazarlo
                if "[imagen]" in cuerpo:
                    cuerpo_incrustado = cuerpo.replace("[imagen]", imagen_incrustada)
                    mensaje.HTMLBody = cuerpo_incrustado
                else:
                    # Si no encuentra [imagen], agrega la imagen al principio del cuerpo
                    mensaje.HTMLBody = f"{imagen_incrustada}<br>{cuerpo}"

                # Agregar todos los destinatarios al mensaje
                mensaje.To = "; ".join(destinatarios)  # Los destinatarios deben estar separados por un punto y coma

                # Enviar mensaje
                mensaje.Send()

                mensaje = "envío exitoso"
                log_result(mensaje)

        else:
            # Datos del correo
            asunto = 'Notificación de error'
            cuerpo = 'Notificación de error'
            adjunto = obtener_nombre_log()

            # Crear objeto Outlook
            outlook = win32com.client.Dispatch('Outlook.Application')
            mensaje = outlook.CreateItem(0)  # 0 = Correo

            # Configurar propiedades del mensaje
            mensaje.Subject = asunto
            mensaje.HTMLBody = cuerpo  # Usar HTMLBody para poder incluir la imagen

            # Adjuntar el log de errores
            attachment = mensaje.Attachments.Add(adjunto)  

            # Agregar todos los destinatarios al mensaje
            mensaje.To = "wlimer-951@hotmail.com"  # Los destinatarios deben estar separados por un punto y coma

            # Enviar mensaje
            mensaje.Send()

            mensaje = "envío exitoso"
            log_result(mensaje)

    except Exception as e:
        # Registrar error
        mensaje_error = f"Error en envío de correo: {e}"
        log_result(mensaje_error)



# Función para obtener los correos
def obtener_correos():
    try:
        df = pd.read_csv(file_destinatarios, header=None)  # Lee el archivo sin cabecera
        correos = df[0].tolist()  # Convertir la columna a lista
        mensaje = 'lectura de correos exitosa'
        log_result(mensaje)
        return correos
    except Exception as e:
        mensaje_error = f"Error al leer los correos: {e}"
        log_result(mensaje_error)
        return []

# Función para obtener el asunto
def obtener_asunto():
    try:
        with open(file_asunto, 'r', encoding='utf-8') as file:
            asunto = file.read().strip()  # Leer el contenido y eliminar posibles saltos de línea
        mensaje = 'lectura de asunto exitosa'
        log_result(mensaje)
        return asunto
    except Exception as e:
        mensaje_error = f"Error al leer el asunto: {str(e)}"
        log_result(mensaje_error)
        return "Error al leer el asunto"

# Función para obtener el cuerpo del mensaje
def obtener_cuerpo():
    try:
        with open(file_cuerpo, 'r', encoding='utf-8') as file:
            cuerpo = file.read().strip()  # Leer el contenido y eliminar posibles saltos de línea
        mensaje = 'lectura de cuerpo de mensaje exitosa'
        log_result(mensaje)
        return cuerpo
    except Exception as e:
        mensaje_error = f"Error al leer el cuerpo: {str(e)}"
        log_result(mensaje_error)
        return "Error al leer el cuerpo"

# Llamada a la función de captura de pantalla después de verificar la conexión
captura_pantalla_completa(url, filecaptura)
