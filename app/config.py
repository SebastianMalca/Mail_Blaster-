"""
Módulo de configuración y constantes globales de la aplicación.
"""

APP_NAME = "Mail Blaster Institucional"
APP_VERSION = "1.0.0"

# Colores del tema
COLORS = {
    "primary": "#1565C0",
    "primary_dark": "#0D47A1",
    "primary_light": "#1976D2",
    "accent": "#00ACC1",
    "success": "#2E7D32",
    "warning": "#F57F17",
    "danger": "#C62828",
    "surface": "#1A1A2E",
    "surface_card": "#16213E",
    "surface_elevated": "#0F3460",
    "text_primary": "#E8EAF6",
    "text_secondary": "#90A4AE",
    "border": "#263238",
    "progress_bg": "#263238",
    "progress_fill": "#00ACC1",
}

# Valores por defecto de configuración SMTP
# Preconfigurado para Google Workspace / Gmail institucional UNMSM
SMTP_DEFAULTS = {
    "host": "smtp.gmail.com",
    "port": "587",
    "user": "",
    "password": "",
    "use_tls": True,   # STARTTLS activo
    "use_ssl": False,  # SSL/TLS directo desactivado
}

# Valores por defecto de envío
# Configuración conservadora para diagnóstico inicial
SEND_DEFAULTS = {
    "batch_size": 25,        # Reducido de 50 a 25 para diagnóstico
    "wait_seconds": 30,      # Aumentado de 10 a 30s para no saturar el servidor
    "subject": "Comunicado Institucional",  # Asunto por defecto
    "max_emails": 500,       # Límite máximo de correos por sesión (0 = sin límite)
}

# Configuración de archivo de guardado
SETTINGS_FILE = "smtp_settings.json"

# Umbral para doble confirmación de seguridad
LARGE_RECIPIENT_THRESHOLD = 500

# Asunto por defecto si el usuario no escribe ninguno
DEFAULT_SUBJECT = "Comunicado Institucional"

# SMTP bloqueado: host y puerto fijos para uso institucional UNMSM
# Cambiar a False para permitir editar host/puerto libremente
LOCK_SMTP_SERVER = True
