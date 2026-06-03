"""
Módulo para manejo de archivos Excel y carga de destinatarios.
"""

import pandas as pd
from pathlib import Path
from typing import List, Tuple


class ExcelLoader:
    """Carga y valida un archivo Excel con columna 'correo'."""

    REQUIRED_COLUMN = "correo"

    def __init__(self, filepath: str):
        self.filepath = filepath
        self._emails: List[str] = []
        self._invalid: List[str] = []
        self._raw_count = 0

    @property
    def emails(self) -> List[str]:
        return self._emails

    @property
    def invalid_emails(self) -> List[str]:
        return self._invalid

    @property
    def raw_count(self) -> int:
        return self._raw_count

    def load(self) -> Tuple[bool, str]:
        """
        Carga el archivo Excel y extrae los correos válidos.
        Retorna (success, message).
        """
        path = Path(self.filepath)

        if not path.exists():
            return False, f"El archivo '{self.filepath}' no existe."

        if path.suffix.lower() != ".xlsx":
            return False, "Solo se aceptan archivos con extensión .xlsx."

        try:
            df = pd.read_excel(self.filepath, engine="openpyxl", dtype=str)
        except Exception as e:
            return False, f"Error al leer el archivo Excel: {e}"

        # Normalizar nombres de columna: lowercase, sin espacios y sin puntuación al final
        # Esto permite detectar columnas como "Correo:", "Correo :", "CORREO", etc.
        df.columns = [c.strip().lower().rstrip(":.,;") .strip() for c in df.columns]

        if self.REQUIRED_COLUMN not in df.columns:
            cols = ", ".join(df.columns.tolist())
            return (
                False,
                f"No se encontró la columna '{self.REQUIRED_COLUMN}'. "
                f"Columnas encontradas: {cols}",
            )

        raw = df[self.REQUIRED_COLUMN].dropna().astype(str).str.strip()
        self._raw_count = len(raw)

        valid = []
        invalid = []
        for email in raw:
            if self._is_valid_email(email):
                valid.append(email)
            else:
                invalid.append(email)

        self._emails = valid
        self._invalid = invalid

        msg = (
            f"Cargados {len(valid)} correos válidos de {self._raw_count} registros."
        )
        if invalid:
            msg += f" ({len(invalid)} correos inválidos omitidos)"

        return True, msg

    @staticmethod
    def _is_valid_email(email: str) -> bool:
        """Validación básica de formato de correo electrónico."""
        if not email or "@" not in email:
            return False
        parts = email.split("@")
        if len(parts) != 2:
            return False
        local, domain = parts
        if not local or not domain or "." not in domain:
            return False
        return True
