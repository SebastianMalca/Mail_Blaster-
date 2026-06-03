"""
Módulo de persistencia de progreso de envío entre sesiones.

Cuando el envío se detiene (por límite diario, cancelación, etc.) se guarda
un archivo JSON con el estado: qué Excel se usaba y cuántos correos se enviaron.
Al reabrir el programa y cargar el mismo Excel, se ofrece continuar desde
donde se quedó, saltando los destinatarios ya procesados.
"""

import json
import os
from datetime import datetime
from pathlib import Path

SESSION_FILE = "send_session.json"


def save_session(excel_path: str, sent_count: int, original_total: int) -> bool:
    """
    Guarda el progreso de envío de la sesión actual.

    Args:
        excel_path:     Ruta absoluta al archivo Excel cargado.
        sent_count:     Total de correos enviados exitosamente (acumulado
                        entre todas las sesiones para este archivo).
        original_total: Total de destinatarios válidos en el Excel.

    Returns:
        True si se guardó con éxito.
    """
    data = {
        "excel_path": str(excel_path),
        "excel_name": Path(excel_path).name,
        "original_total": original_total,
        "sent_count": sent_count,
        "pending": original_total - sent_count,
        "saved_at": datetime.now().isoformat(timespec="seconds"),
    }
    try:
        with open(SESSION_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


def load_session(excel_path: str) -> dict | None:
    """
    Carga la sesión guardada solo si corresponde al mismo archivo Excel.

    Args:
        excel_path: Ruta al Excel que se acaba de cargar.

    Returns:
        Dict con los datos de sesión, o None si no hay sesión válida
        para ese archivo.
    """
    path = Path(SESSION_FILE)
    if not path.exists():
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        # Comparar por ruta completa; si el usuario movió el archivo
        # también se acepta coincidencia por nombre.
        if (data.get("excel_path") == str(excel_path) or
                data.get("excel_name") == Path(excel_path).name):
            return data
        return None
    except Exception:
        return None


def load_any_session() -> dict | None:
    """
    Carga la sesión guardada sin filtrar por archivo Excel.
    Útil para mostrar un aviso al abrir la app si hay una sesión pendiente.
    """
    path = Path(SESSION_FILE)
    if not path.exists():
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def clear_session() -> None:
    """Elimina la sesión guardada (llamar cuando el envío finaliza al 100%)."""
    try:
        if Path(SESSION_FILE).exists():
            os.remove(SESSION_FILE)
    except Exception:
        pass


def session_exists() -> bool:
    """Indica si existe una sesión pendiente guardada."""
    return Path(SESSION_FILE).exists()
