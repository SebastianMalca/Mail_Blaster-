# ✉️ Mail Blaster Institucional

**Mail Blaster Institucional** es una aplicación de escritorio moderna y eficiente diseñada para el envío masivo de correos electrónicos a través de servidores SMTP utilizando listas de destinatarios en formato Excel (`.xlsx`). La interfaz gráfica es intuitiva y está optimizada con un diseño oscuro elegante y estadísticas en tiempo real.

---

## 🚀 Características Principales

- **Interfaz Gráfica Premium**: Diseñada en modo oscuro utilizando `customtkinter` para una experiencia fluida y moderna.
- **Detección de Destinatarios**: Carga de contactos directamente desde hojas de Excel (`.xlsx`), con validación automática y depuración de correos electrónicos inválidos o vacíos.
- **Formatos de Correo Flexibles**:
  - **HTML**: Envío de boletines prediseñados en formato web.
  - **Imagen**: Envío de boletines visuales (JPEG/PNG) que la aplicación convierte automáticamente a código HTML responsivo (Base64) para una correcta visualización en cualquier dispositivo.
- **Envío Inteligente por Lotes**: Permite definir la cantidad de correos por lote y el intervalo de espera en segundos entre lotes para evitar bloqueos del servidor de correo o ser catalogado como SPAM.
- **Control Total y Estadísticas en Vivo**:
  - Panel indicador con cantidad de correos Totales, Enviados, Fallidos y Pendientes.
  - Barra de progreso gráfica y log de actividad detallado en tiempo real.
  - Opción de **Cancelar Envío** de manera segura en cualquier momento.
- **Prueba de Conexión**: Permite verificar las credenciales del servidor SMTP antes de iniciar el envío masivo.
- **Persistencia**: Guarda la configuración de red SMTP localmente (`smtp_settings.json`) exceptuando la contraseña por seguridad.

---

## 🛠️ Requisitos Previos

Para ejecutar la aplicación directamente desde el código fuente, necesitas tener instalado:
1. **Python 3.11 o superior** (Se recomienda Python 3.12).
2. Las dependencias especificadas en [requirements.txt](file:///c:/Users/OFICINAFL/Desktop/Automatizacion%20de%20correos/requirements.txt):
   - `customtkinter` (Para la interfaz gráfica)
   - `pandas` y `openpyxl` (Para la lectura y procesamiento de hojas de Excel)
   - `Pillow` (Para el procesamiento de imágenes)

---

## 🏁 Cómo Ejecutar la Aplicación

Tienes tres opciones para ejecutar el proyecto, ordenadas de la más sencilla a la más avanzada:

### Opción 1: Ejecución Automática (Recomendada para Windows)
Simplemente haz doble clic en el archivo [instalar_y_ejecutar.bat](file:///c:/Users/OFICINAFL/Desktop/Automatizacion%20de%20correos/instalar_y_ejecutar.bat).
Este script se encargará de:
1. Buscar Python instalado en tu sistema.
2. Instalar o actualizar de forma automática las dependencias necesarias.
3. Iniciar la aplicación ejecutando [main.py](file:///c:/Users/OFICINAFL/Desktop/Automatizacion%20de%20correos/main.py).

*Nota: Si no tienes Python instalado, el script te avisará. Puedes usar [instalar_python.bat](file:///c:/Users/OFICINAFL/Desktop/Automatizacion%20de%20correos/instalar_python.bat) si previamente descargaste el instalador oficial de Python en tu carpeta de Descargas.*

---

### Opción 2: Ejecución Manual por Consola (Línea de Comandos)
Si prefieres controlar el proceso manualmente o estás en otro sistema operativo:

1. Abre la terminal en el directorio del proyecto y ejecuta el siguiente comando para instalar las librerías necesarias:
   ```bash
   pip install -r requirements.txt
   ```
2. Ejecuta la aplicación iniciando el punto de entrada principal:
   ```bash
   python main.py
   ```

---

### Opción 3: Crear y Ejecutar como un Archivo Portable `.exe`
Si deseas compartir la aplicación con otros usuarios sin que tengan que instalar Python, puedes generar un ejecutable autocontenido:

1. Ejecuta el archivo [build_exe.bat](file:///c:/Users/OFICINAFL/Desktop/Automatizacion%20de%20correos/build_exe.bat) haciendo doble clic en él.
2. Este script instalará `pyinstaller` y empaquetará todo el código en un único archivo ejecutable.
3. Al finalizar, encontrarás el programa ejecutable listo para usar en la ruta:
   `dist/MailBlaster.exe`

---

## 📊 Formato del Archivo Excel (`.xlsx`)

Para que la aplicación procese correctamente la lista de contactos, el archivo Excel debe cumplir con lo siguiente:
- Debe tener una columna cuyo encabezado sea exactamente **`correo`** (se aceptan variaciones en mayúsculas/minúsculas como `Correo` o `CORREO`).
- Las filas de esta columna deben contener los correos electrónicos a los cuales se enviará el mensaje.
- Cualquier otra columna adicional será ignorada por completo, por lo que puedes mantener nombres, teléfonos u otros datos de control en el mismo archivo.

Ejemplo visual de la hoja de Excel:
| correo | nombre *(opcional)* |
| :--- | :--- |
| juan.perez@institucion.edu.pe | Juan Perez |
| maria.lopez@institucion.edu.pe | Maria Lopez |
| soporte@empresa.com | Soporte |

---

## 🔒 Configuración del Servidor SMTP

Para realizar los envíos, debes completar la configuración SMTP en el panel izquierdo de la aplicación:

1. **Servidor SMTP (Host)**: Dirección del servidor de correos (ej. `smtp.gmail.com` para Gmail, `smtp.office365.com` para Outlook/Office365).
2. **Puerto**: Por lo general, `587` si se usa STARTTLS, o `465` para SSL directo.
3. **Usuario**: Tu dirección de correo completa.
4. **Contraseña**:
   - **IMPORTANTE (Gmail/Outlook)**: Los proveedores de correo actuales no permiten usar tu contraseña de inicio de sesión habitual por motivos de seguridad. Deberás activar la **Verificación en Dos Pasos** en tu cuenta y generar una **Contraseña de Aplicación** específica para SMTP. Usa esa contraseña de 16 caracteres en la aplicación.
5. **TLS/SSL**: Activa el interruptor correspondiente según los requerimientos de tu proveedor de correo electrónico.
6. Presiona **🔌 Probar Conexión** para validar que todo funcione correctamente antes de iniciar un envío masivo.
