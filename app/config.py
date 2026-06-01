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
SMTP_DEFAULTS = {
    "host": "",
    "port": "587",
    "user": "",
    "password": "",
    "use_tls": True,
    "use_ssl": False,
}

# Valores por defecto de envío
SEND_DEFAULTS = {
    "batch_size": 50,
    "wait_seconds": 10,
    "subject": "",
}

# Configuración de archivo de guardado
SETTINGS_FILE = "smtp_settings.json"
