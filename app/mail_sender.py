"""
Módulo de lógica de envío de correos.
Maneja SMTP, construcción de mensajes y envío por lotes.
"""

import smtplib
import ssl
import time
import threading
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path
from typing import Callable, List, Optional
import base64


class SMTPConfig:
    """Configuración del servidor SMTP."""

    def __init__(
        self,
        host: str,
        port: int,
        user: str,
        password: str,
        use_tls: bool = True,
        use_ssl: bool = False,
    ):
        self.host = host
        self.port = port
        self.user = user
        self.password = password
        self.use_tls = use_tls
        self.use_ssl = use_ssl

    def validate(self) -> tuple[bool, str]:
        """Valida que los campos requeridos estén completos."""
        if not self.host.strip():
            return False, "El servidor SMTP es requerido."
        if not self.port or self.port <= 0:
            return False, "El puerto debe ser un número válido."
        if not self.user.strip():
            return False, "El usuario SMTP es requerido."
        if not self.password.strip():
            return False, "La contraseña SMTP es requerida."
        return True, ""


class MailContent:
    """Contenido del correo a enviar."""

    def __init__(self, subject: str, content_type: str, content_path: str):
        self.subject = subject
        self.content_type = content_type  # 'html' o 'image'
        self.content_path = content_path

    def validate(self) -> tuple[bool, str]:
        if not self.subject.strip():
            return False, "El asunto del correo es requerido."
        if not self.content_path:
            return False, "Debe seleccionar un archivo HTML o imagen."
        if not Path(self.content_path).exists():
            return False, f"El archivo '{self.content_path}' no existe."
        return True, ""


class SendStats:
    """Estadísticas de envío en tiempo real."""

    def __init__(self, total: int):
        self.total = total
        self.sent = 0
        self.failed = 0
        self.pending = total
        self._lock = threading.Lock()

    def record_sent(self):
        with self._lock:
            self.sent += 1
            self.pending -= 1

    def record_failed(self):
        with self._lock:
            self.failed += 1
            self.pending -= 1

    @property
    def progress_pct(self) -> float:
        if self.total == 0:
            return 0.0
        return (self.sent + self.failed) / self.total


class MailSender:
    """Motor de envío masivo de correos por lotes."""

    def __init__(
        self,
        smtp_config: SMTPConfig,
        mail_content: MailContent,
        recipients: List[str],
        batch_size: int,
        wait_seconds: float,
        on_progress: Optional[Callable] = None,
        on_log: Optional[Callable] = None,
        on_finished: Optional[Callable] = None,
    ):
        self.smtp_config = smtp_config
        self.mail_content = mail_content
        self.recipients = recipients
        self.batch_size = batch_size
        self.wait_seconds = wait_seconds
        self.on_progress = on_progress
        self.on_log = on_log
        self.on_finished = on_finished
        self._cancel_event = threading.Event()
        self.stats = SendStats(len(recipients))

    def cancel(self):
        """Solicita la cancelación del envío."""
        self._cancel_event.set()

    def _log(self, msg: str):
        if self.on_log:
            self.on_log(msg)

    def _notify_progress(self):
        if self.on_progress:
            self.on_progress(self.stats)

    def _build_message(self, recipient: str) -> MIMEMultipart:
        """Construye el objeto MIMEMultipart del mensaje."""
        msg = MIMEMultipart("related" if self.mail_content.content_type == "image" else "mixed")
        msg["Subject"] = self.mail_content.subject
        msg["From"] = self.smtp_config.user
        msg["To"] = recipient

        if self.mail_content.content_type == "html":
            with open(self.mail_content.content_path, "r", encoding="utf-8") as f:
                html_body = f.read()
            msg_alt = MIMEMultipart("alternative")
            msg_alt.attach(MIMEText("Este correo requiere un cliente compatible con HTML.", "plain", "utf-8"))
            msg_alt.attach(MIMEText(html_body, "html", "utf-8"))
            msg.attach(msg_alt)

        elif self.mail_content.content_type == "image":
            path = self.mail_content.content_path
            ext = Path(path).suffix.lower().lstrip(".")
            mime_type = "jpeg" if ext in ("jpg", "jpeg") else ext

            with open(path, "rb") as f:
                img_data = f.read()

            img_b64 = base64.b64encode(img_data).decode()
            html_body = (
                f'<html><body style="margin:0;padding:0;background:#fff;">'
                f'<img src="data:image/{mime_type};base64,{img_b64}" '
                f'style="max-width:100%;display:block;margin:auto;" alt="Comunicado"/>'
                f"</body></html>"
            )
            msg_alt = MIMEMultipart("alternative")
            msg_alt.attach(MIMEText("Este correo requiere un cliente compatible con HTML.", "plain", "utf-8"))
            msg_alt.attach(MIMEText(html_body, "html", "utf-8"))
            msg.attach(msg_alt)

        return msg

    def _connect_smtp(self) -> smtplib.SMTP:
        """Crea y devuelve una conexión SMTP autenticada."""
        cfg = self.smtp_config
        context = ssl.create_default_context()

        if cfg.use_ssl:
            server = smtplib.SMTP_SSL(cfg.host, cfg.port, context=context, timeout=30)
        else:
            server = smtplib.SMTP(cfg.host, cfg.port, timeout=30)
            if cfg.use_tls:
                server.ehlo()
                server.starttls(context=context)
                server.ehlo()

        server.login(cfg.user, cfg.password)
        return server

    def test_connection(self) -> tuple[bool, str]:
        """Prueba la conexión SMTP sin enviar correos."""
        try:
            server = self._connect_smtp()
            server.quit()
            return True, "Conexión exitosa al servidor SMTP."
        except smtplib.SMTPAuthenticationError:
            return False, "Error de autenticación: usuario o contraseña incorrectos."
        except smtplib.SMTPConnectError as e:
            return False, f"Error al conectar al servidor: {e}"
        except Exception as e:
            return False, f"Error inesperado: {e}"

    def run(self):
        """Ejecuta el envío masivo. Debe llamarse desde un hilo separado."""
        self._cancel_event.clear()
        self._log("Iniciando proceso de envío masivo...")
        self._log(f"Total de destinatarios: {self.stats.total}")
        self._log(f"Tamaño de lote: {self.batch_size} | Espera entre lotes: {self.wait_seconds}s")

        batches = [
            self.recipients[i : i + self.batch_size]
            for i in range(0, len(self.recipients), self.batch_size)
        ]
        total_batches = len(batches)

        for batch_num, batch in enumerate(batches, start=1):
            if self._cancel_event.is_set():
                self._log("⚠️ Envío cancelado por el usuario.")
                break

            self._log(f"--- Procesando lote {batch_num}/{total_batches} ({len(batch)} correos) ---")

            try:
                server = self._connect_smtp()
            except Exception as e:
                self._log(f"❌ Error al conectar al servidor SMTP: {e}")
                # Marcar todos del lote como fallidos
                for _ in batch:
                    self.stats.record_failed()
                self._notify_progress()
                continue

            for recipient in batch:
                if self._cancel_event.is_set():
                    server.quit()
                    self._log("⚠️ Envío cancelado por el usuario.")
                    # Marcar restantes del lote como pendientes -> fallidos
                    remaining = len(batch) - batch.index(recipient)
                    for _ in range(remaining):
                        self.stats.record_failed()
                    self._notify_progress()
                    if self.on_finished:
                        self.on_finished(self.stats, cancelled=True)
                    return

                try:
                    msg = self._build_message(recipient)
                    server.sendmail(self.smtp_config.user, recipient, msg.as_string())
                    self.stats.record_sent()
                    self._log(f"✅ Enviado a: {recipient}")
                except Exception as e:
                    self.stats.record_failed()
                    self._log(f"❌ Error enviando a {recipient}: {e}")

                self._notify_progress()

            try:
                server.quit()
            except Exception:
                pass

            # Esperar entre lotes (excepto el último)
            if batch_num < total_batches and not self._cancel_event.is_set():
                self._log(f"⏳ Esperando {self.wait_seconds}s antes del siguiente lote...")
                for _ in range(int(self.wait_seconds)):
                    if self._cancel_event.is_set():
                        break
                    time.sleep(1)

        self._log("=" * 50)
        self._log(
            f"Proceso finalizado. Enviados: {self.stats.sent} | "
            f"Fallidos: {self.stats.failed} | Total: {self.stats.total}"
        )

        if self.on_finished:
            self.on_finished(self.stats, cancelled=self._cancel_event.is_set())
