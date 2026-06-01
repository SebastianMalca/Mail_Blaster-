"""
Interfaz gráfica principal de Mail Blaster Institucional.
Construida con CustomTkinter.
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
import tkinter as tk
from PIL import Image, ImageTk
import threading
import os
import sys
from pathlib import Path
import webbrowser
import tempfile

from app.config import APP_NAME, APP_VERSION, COLORS, SMTP_DEFAULTS, SEND_DEFAULTS
from app.mail_sender import SMTPConfig, MailContent, MailSender, SendStats
from app.excel_loader import ExcelLoader
from app.settings_manager import save_settings, load_settings


# ───────────────────────────────────────────────
#  Configuración de apariencia
# ───────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


# ═══════════════════════════════════════════════════════════════════
#  COMPONENTES REUTILIZABLES
# ═══════════════════════════════════════════════════════════════════


class SectionLabel(ctk.CTkLabel):
    """Etiqueta de sección con línea decorativa."""

    def __init__(self, parent, text: str, **kwargs):
        super().__init__(
            parent,
            text=f"  {text}",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=COLORS["accent"],
            anchor="w",
            **kwargs,
        )


class StatCard(ctk.CTkFrame):
    """Tarjeta de estadística individual."""

    def __init__(self, parent, label: str, color: str, **kwargs):
        super().__init__(
            parent,
            corner_radius=12,
            fg_color=COLORS["surface_elevated"],
            border_width=1,
            border_color=color,
            **kwargs,
        )
        self._value_var = ctk.StringVar(value="0")
        self._color = color

        ctk.CTkLabel(
            self,
            text=label,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=COLORS["text_secondary"],
        ).pack(pady=(12, 2))

        ctk.CTkLabel(
            self,
            textvariable=self._value_var,
            font=ctk.CTkFont(family="Segoe UI", size=28, weight="bold"),
            text_color=color,
        ).pack(pady=(0, 12))

    def set_value(self, val: int):
        self._value_var.set(str(val))


class LogBox(ctk.CTkTextbox):
    """Cuadro de log con scroll automático."""

    def __init__(self, parent, **kwargs):
        super().__init__(
            parent,
            font=ctk.CTkFont(family="Consolas", size=11),
            fg_color="#0A0A1A",
            text_color="#A5D6A7",
            border_color=COLORS["border"],
            border_width=1,
            corner_radius=8,
            wrap="word",
            state="disabled",
            **kwargs,
        )

    def append(self, text: str):
        self.configure(state="normal")
        self.insert("end", text + "\n")
        self.see("end")
        self.configure(state="disabled")

    def clear(self):
        self.configure(state="normal")
        self.delete("1.0", "end")
        self.configure(state="disabled")


# ═══════════════════════════════════════════════════════════════════
#  VENTANA PRINCIPAL
# ═══════════════════════════════════════════════════════════════════


class MailBlasterApp(ctk.CTk):

    def __init__(self):
        super().__init__()

        self.title(f"{APP_NAME} v{APP_VERSION}")
        self.geometry("1200x820")
        self.minsize(1000, 700)
        self.configure(fg_color=COLORS["surface"])

        # Estado interno
        self._recipients: list[str] = []
        self._content_path: str = ""
        self._content_type: str = ""          # 'html' o 'image'
        self._sender: MailSender | None = None
        self._send_thread: threading.Thread | None = None
        self._preview_window: tk.Toplevel | None = None

        # Cargar configuración guardada
        self._saved_settings = load_settings()

        self._build_ui()
        self._load_saved_settings()

    # ─────────────────────────────────────────────
    #  CONSTRUCCIÓN DE UI
    # ─────────────────────────────────────────────

    def _build_ui(self):
        """Construye todos los elementos de la interfaz."""
        self._build_header()
        self._build_main_area()
        self._build_status_bar()

    def _build_header(self):
        """Barra superior con título y logo."""
        header = ctk.CTkFrame(
            self,
            height=64,
            corner_radius=0,
            fg_color=COLORS["surface_card"],
            border_width=0,
        )
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        # Título
        ctk.CTkLabel(
            header,
            text="✉  Mail Blaster Institucional",
            font=ctk.CTkFont(family="Segoe UI", size=20, weight="bold"),
            text_color=COLORS["text_primary"],
        ).pack(side="left", padx=24, pady=12)

        # Versión
        ctk.CTkLabel(
            header,
            text=f"v{APP_VERSION}",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=COLORS["text_secondary"],
        ).pack(side="right", padx=24)

    def _build_main_area(self):
        """Área principal dividida en panel izquierdo (config) y derecho (progreso)."""
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=16, pady=(12, 0))

        # ── Panel izquierdo (configuración) ──
        left = ctk.CTkScrollableFrame(
            main,
            width=480,
            corner_radius=12,
            fg_color=COLORS["surface_card"],
            border_width=1,
            border_color=COLORS["border"],
            scrollbar_button_color=COLORS["surface_elevated"],
        )
        left.pack(side="left", fill="y", padx=(0, 8))

        # ── Panel derecho (estadísticas + log) ──
        right = ctk.CTkFrame(
            main,
            corner_radius=12,
            fg_color=COLORS["surface_card"],
            border_width=1,
            border_color=COLORS["border"],
        )
        right.pack(side="left", fill="both", expand=True)

        self._build_left_panel(left)
        self._build_right_panel(right)

    def _build_left_panel(self, parent):
        """Panel de configuración (izquierda)."""
        pad = {"padx": 16, "pady": 4}

        # ── Encabezado del panel ──
        ctk.CTkLabel(
            parent,
            text="Configuración de Envío",
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            text_color=COLORS["text_primary"],
            anchor="w",
        ).pack(fill="x", padx=16, pady=(16, 8))

        # ════════════════════════════════════
        # SECCIÓN 1: SMTP
        # ════════════════════════════════════
        SectionLabel(parent, "⚙  Servidor SMTP").pack(fill="x", **pad)
        sep1 = ctk.CTkFrame(parent, height=1, fg_color=COLORS["border"])
        sep1.pack(fill="x", padx=16, pady=(0, 8))

        self._smtp_host = self._make_entry(parent, "Servidor SMTP (host)", "smtp.ejemplo.com")
        self._smtp_port = self._make_entry(parent, "Puerto", "587")
        self._smtp_user = self._make_entry(parent, "Usuario / Correo remitente", "usuario@ejemplo.com")
        self._smtp_pass = self._make_entry(parent, "Contraseña", "", show="●")

        # TLS / SSL switches
        tls_frame = ctk.CTkFrame(parent, fg_color="transparent")
        tls_frame.pack(fill="x", padx=16, pady=4)

        self._tls_var = ctk.BooleanVar(value=True)
        self._ssl_var = ctk.BooleanVar(value=False)

        ctk.CTkSwitch(
            tls_frame,
            text="Usar STARTTLS",
            variable=self._tls_var,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=COLORS["text_primary"],
            button_color=COLORS["accent"],
            progress_color=COLORS["primary"],
        ).pack(side="left", padx=(0, 16))

        ctk.CTkSwitch(
            tls_frame,
            text="Usar SSL/TLS directo",
            variable=self._ssl_var,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=COLORS["text_primary"],
            button_color=COLORS["accent"],
            progress_color=COLORS["primary"],
        ).pack(side="left")

        # Botón probar conexión
        ctk.CTkButton(
            parent,
            text="🔌  Probar Conexión",
            command=self._test_connection,
            height=34,
            corner_radius=8,
            fg_color=COLORS["surface_elevated"],
            hover_color=COLORS["primary"],
            text_color=COLORS["text_primary"],
            font=ctk.CTkFont(family="Segoe UI", size=12),
            border_width=1,
            border_color=COLORS["accent"],
        ).pack(fill="x", padx=16, pady=(6, 10))

        # ════════════════════════════════════
        # SECCIÓN 2: Destinatarios
        # ════════════════════════════════════
        SectionLabel(parent, "👥  Destinatarios (Excel .xlsx)").pack(fill="x", **pad)
        sep2 = ctk.CTkFrame(parent, height=1, fg_color=COLORS["border"])
        sep2.pack(fill="x", padx=16, pady=(0, 8))

        excel_frame = ctk.CTkFrame(parent, fg_color="transparent")
        excel_frame.pack(fill="x", padx=16, pady=4)

        self._excel_path_var = ctk.StringVar(value="Ningún archivo seleccionado")
        ctk.CTkLabel(
            excel_frame,
            textvariable=self._excel_path_var,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=COLORS["text_secondary"],
            anchor="w",
            wraplength=320,
        ).pack(side="left", fill="x", expand=True)

        ctk.CTkButton(
            excel_frame,
            text="📂 Abrir",
            command=self._load_excel,
            width=80,
            height=32,
            corner_radius=8,
            fg_color=COLORS["primary"],
            hover_color=COLORS["primary_dark"],
            font=ctk.CTkFont(family="Segoe UI", size=12),
        ).pack(side="right")

        self._recipients_label = ctk.CTkLabel(
            parent,
            text="0 destinatarios cargados",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=COLORS["text_secondary"],
            anchor="w",
        )
        self._recipients_label.pack(fill="x", padx=16, pady=(2, 10))

        # ════════════════════════════════════
        # SECCIÓN 3: Contenido del correo
        # ════════════════════════════════════
        SectionLabel(parent, "📄  Contenido del Correo").pack(fill="x", **pad)
        sep3 = ctk.CTkFrame(parent, height=1, fg_color=COLORS["border"])
        sep3.pack(fill="x", padx=16, pady=(0, 8))

        # Asunto
        ctk.CTkLabel(
            parent,
            text="Asunto:",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=COLORS["text_secondary"],
            anchor="w",
        ).pack(fill="x", padx=16, pady=(2, 0))
        self._subject_entry = ctk.CTkEntry(
            parent,
            placeholder_text="Asunto del correo masivo",
            height=36,
            corner_radius=8,
            border_color=COLORS["border"],
            font=ctk.CTkFont(family="Segoe UI", size=13),
        )
        self._subject_entry.pack(fill="x", padx=16, pady=(2, 8))

        # Botones de selección de archivo
        file_btn_frame = ctk.CTkFrame(parent, fg_color="transparent")
        file_btn_frame.pack(fill="x", padx=16, pady=4)

        ctk.CTkButton(
            file_btn_frame,
            text="🌐 Seleccionar HTML",
            command=self._load_html,
            height=36,
            corner_radius=8,
            fg_color=COLORS["primary"],
            hover_color=COLORS["primary_dark"],
            font=ctk.CTkFont(family="Segoe UI", size=12),
        ).pack(side="left", expand=True, fill="x", padx=(0, 4))

        ctk.CTkButton(
            file_btn_frame,
            text="🖼 Seleccionar Imagen",
            command=self._load_image,
            height=36,
            corner_radius=8,
            fg_color=COLORS["surface_elevated"],
            hover_color=COLORS["primary"],
            border_width=1,
            border_color=COLORS["primary"],
            font=ctk.CTkFont(family="Segoe UI", size=12),
        ).pack(side="right", expand=True, fill="x", padx=(4, 0))

        # Indicador de archivo seleccionado
        self._content_path_var = ctk.StringVar(value="Ningún archivo de contenido seleccionado")
        ctk.CTkLabel(
            parent,
            textvariable=self._content_path_var,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=COLORS["text_secondary"],
            anchor="w",
            wraplength=430,
        ).pack(fill="x", padx=16, pady=(4, 2))

        # Botón de vista previa
        self._preview_btn = ctk.CTkButton(
            parent,
            text="👁  Ver Vista Previa",
            command=self._show_preview,
            height=32,
            corner_radius=8,
            state="disabled",
            fg_color=COLORS["surface_elevated"],
            hover_color=COLORS["accent"],
            text_color=COLORS["text_primary"],
            font=ctk.CTkFont(family="Segoe UI", size=12),
            border_width=1,
            border_color=COLORS["accent"],
        )
        self._preview_btn.pack(fill="x", padx=16, pady=(2, 10))

        # ════════════════════════════════════
        # SECCIÓN 4: Configuración de lotes
        # ════════════════════════════════════
        SectionLabel(parent, "⚡  Configuración de Lotes").pack(fill="x", **pad)
        sep4 = ctk.CTkFrame(parent, height=1, fg_color=COLORS["border"])
        sep4.pack(fill="x", padx=16, pady=(0, 8))

        batch_row = ctk.CTkFrame(parent, fg_color="transparent")
        batch_row.pack(fill="x", padx=16, pady=4)

        # Tamaño del lote
        lote_frame = ctk.CTkFrame(batch_row, fg_color="transparent")
        lote_frame.pack(side="left", expand=True, fill="x", padx=(0, 8))
        ctk.CTkLabel(
            lote_frame,
            text="Correos por lote:",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=COLORS["text_secondary"],
            anchor="w",
        ).pack(anchor="w")
        self._batch_size_entry = ctk.CTkEntry(
            lote_frame,
            placeholder_text="50",
            height=34,
            corner_radius=8,
            border_color=COLORS["border"],
            font=ctk.CTkFont(family="Segoe UI", size=13),
        )
        self._batch_size_entry.pack(fill="x")
        self._batch_size_entry.insert(0, str(SEND_DEFAULTS["batch_size"]))

        # Tiempo de espera
        wait_frame = ctk.CTkFrame(batch_row, fg_color="transparent")
        wait_frame.pack(side="right", expand=True, fill="x")
        ctk.CTkLabel(
            wait_frame,
            text="Espera entre lotes (segundos):",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=COLORS["text_secondary"],
            anchor="w",
        ).pack(anchor="w")
        self._wait_entry = ctk.CTkEntry(
            wait_frame,
            placeholder_text="10",
            height=34,
            corner_radius=8,
            border_color=COLORS["border"],
            font=ctk.CTkFont(family="Segoe UI", size=13),
        )
        self._wait_entry.pack(fill="x")
        self._wait_entry.insert(0, str(SEND_DEFAULTS["wait_seconds"]))

        # ════════════════════════════════════
        # BOTONES DE ACCIÓN
        # ════════════════════════════════════
        btn_sep = ctk.CTkFrame(parent, height=1, fg_color=COLORS["border"])
        btn_sep.pack(fill="x", padx=16, pady=16)

        self._start_btn = ctk.CTkButton(
            parent,
            text="▶   INICIAR ENVÍO",
            command=self._start_sending,
            height=48,
            corner_radius=10,
            fg_color=COLORS["success"],
            hover_color="#1B5E20",
            text_color="white",
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
        )
        self._start_btn.pack(fill="x", padx=16, pady=(0, 8))

        self._cancel_btn = ctk.CTkButton(
            parent,
            text="⏹   CANCELAR ENVÍO",
            command=self._cancel_sending,
            height=44,
            corner_radius=10,
            fg_color=COLORS["danger"],
            hover_color="#7F0000",
            text_color="white",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            state="disabled",
        )
        self._cancel_btn.pack(fill="x", padx=16, pady=(0, 16))

        # Botón guardar configuración
        ctk.CTkButton(
            parent,
            text="💾  Guardar Configuración SMTP",
            command=self._save_settings,
            height=34,
            corner_radius=8,
            fg_color=COLORS["surface_elevated"],
            hover_color=COLORS["primary"],
            text_color=COLORS["text_secondary"],
            font=ctk.CTkFont(family="Segoe UI", size=11),
        ).pack(fill="x", padx=16, pady=(0, 16))

    def _build_right_panel(self, parent):
        """Panel de progreso y log (derecha)."""
        # ── Encabezado ──
        header_row = ctk.CTkFrame(parent, fg_color="transparent")
        header_row.pack(fill="x", padx=16, pady=(16, 8))

        ctk.CTkLabel(
            header_row,
            text="Panel de Progreso",
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            text_color=COLORS["text_primary"],
            anchor="w",
        ).pack(side="left")

        self._clear_log_btn = ctk.CTkButton(
            header_row,
            text="🗑 Limpiar Log",
            command=lambda: self._log_box.clear(),
            width=110,
            height=28,
            corner_radius=6,
            fg_color=COLORS["surface_elevated"],
            hover_color=COLORS["border"],
            text_color=COLORS["text_secondary"],
            font=ctk.CTkFont(family="Segoe UI", size=11),
        )
        self._clear_log_btn.pack(side="right")

        # ── Tarjetas de estadísticas ──
        stats_frame = ctk.CTkFrame(parent, fg_color="transparent")
        stats_frame.pack(fill="x", padx=16, pady=(0, 12))

        self._card_total = StatCard(stats_frame, "TOTAL", COLORS["text_primary"])
        self._card_total.pack(side="left", expand=True, fill="x", padx=(0, 6))

        self._card_sent = StatCard(stats_frame, "ENVIADOS", COLORS["success"])
        self._card_sent.pack(side="left", expand=True, fill="x", padx=3)

        self._card_failed = StatCard(stats_frame, "FALLIDOS", COLORS["danger"])
        self._card_failed.pack(side="left", expand=True, fill="x", padx=3)

        self._card_pending = StatCard(stats_frame, "PENDIENTES", COLORS["warning"])
        self._card_pending.pack(side="right", expand=True, fill="x", padx=(6, 0))

        # ── Barra de progreso ──
        prog_frame = ctk.CTkFrame(parent, fg_color="transparent")
        prog_frame.pack(fill="x", padx=16, pady=(0, 4))

        self._progress_pct_var = ctk.StringVar(value="0%")
        pct_row = ctk.CTkFrame(prog_frame, fg_color="transparent")
        pct_row.pack(fill="x")

        ctk.CTkLabel(
            pct_row,
            text="Progreso:",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=COLORS["text_secondary"],
            anchor="w",
        ).pack(side="left")

        ctk.CTkLabel(
            pct_row,
            textvariable=self._progress_pct_var,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=COLORS["accent"],
            anchor="e",
        ).pack(side="right")

        self._progress_bar = ctk.CTkProgressBar(
            prog_frame,
            height=20,
            corner_radius=10,
            progress_color=COLORS["accent"],
            fg_color=COLORS["progress_bg"],
        )
        self._progress_bar.pack(fill="x", pady=(4, 0))
        self._progress_bar.set(0)

        # ── Estado actual ──
        self._status_var = ctk.StringVar(value="Listo para enviar.")
        ctk.CTkLabel(
            parent,
            textvariable=self._status_var,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=COLORS["text_secondary"],
            anchor="w",
        ).pack(fill="x", padx=16, pady=(4, 8))

        # ── Log de actividad ──
        SectionLabel(parent, "📋  Log de Actividad").pack(fill="x", padx=16, pady=(4, 4))
        self._log_box = LogBox(parent)
        self._log_box.pack(fill="both", expand=True, padx=16, pady=(0, 16))

    def _build_status_bar(self):
        """Barra de estado inferior."""
        bar = ctk.CTkFrame(
            self,
            height=28,
            corner_radius=0,
            fg_color=COLORS["surface_card"],
        )
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)

        self._statusbar_var = ctk.StringVar(value="Listo.")
        ctk.CTkLabel(
            bar,
            textvariable=self._statusbar_var,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=COLORS["text_secondary"],
            anchor="w",
        ).pack(side="left", padx=12)

        ctk.CTkLabel(
            bar,
            text=APP_NAME,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=COLORS["text_secondary"],
            anchor="e",
        ).pack(side="right", padx=12)

    # ─────────────────────────────────────────────
    #  HELPERS: WIDGETS
    # ─────────────────────────────────────────────

    def _make_entry(self, parent, label: str, placeholder: str = "", show: str = "") -> ctk.CTkEntry:
        """Crea una entrada con etiqueta."""
        ctk.CTkLabel(
            parent,
            text=f"{label}:",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=COLORS["text_secondary"],
            anchor="w",
        ).pack(fill="x", padx=16, pady=(4, 0))
        entry = ctk.CTkEntry(
            parent,
            placeholder_text=placeholder,
            height=36,
            corner_radius=8,
            border_color=COLORS["border"],
            font=ctk.CTkFont(family="Segoe UI", size=13),
            show=show,
        )
        entry.pack(fill="x", padx=16, pady=(2, 4))
        return entry

    # ─────────────────────────────────────────────
    #  CARGA Y VALIDACIÓN
    # ─────────────────────────────────────────────

    def _load_saved_settings(self):
        """Carga la configuración guardada en los campos de la UI."""
        s = self._saved_settings
        self._smtp_host.insert(0, s.get("host", ""))
        self._smtp_port.delete(0, "end")
        self._smtp_port.insert(0, str(s.get("port", "587")))
        self._smtp_user.insert(0, s.get("user", ""))
        self._tls_var.set(s.get("use_tls", True))
        self._ssl_var.set(s.get("use_ssl", False))

    def _save_settings(self):
        """Guarda la configuración SMTP actual."""
        data = {
            "host": self._smtp_host.get().strip(),
            "port": self._smtp_port.get().strip(),
            "user": self._smtp_user.get().strip(),
            "use_tls": self._tls_var.get(),
            "use_ssl": self._ssl_var.get(),
        }
        if save_settings(data):
            self._statusbar_var.set("✅ Configuración guardada.")
            messagebox.showinfo("Configuración", "Configuración SMTP guardada correctamente.\n(La contraseña no se guarda por seguridad)")
        else:
            messagebox.showerror("Error", "No se pudo guardar la configuración.")

    def _load_excel(self):
        """Abre diálogo para cargar archivo Excel."""
        path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if not path:
            return

        loader = ExcelLoader(path)
        ok, msg = loader.load()

        if not ok:
            messagebox.showerror("Error al cargar Excel", msg)
            return

        self._recipients = loader.emails
        short_name = Path(path).name
        self._excel_path_var.set(short_name)
        self._recipients_label.configure(
            text=f"✅ {len(self._recipients):,} destinatarios válidos cargados",
            text_color=COLORS["success"],
        )
        self._card_total.set_value(len(self._recipients))
        self._card_pending.set_value(len(self._recipients))
        self._log_box.append(f"📂 Excel cargado: {short_name}")
        self._log_box.append(f"   → {msg}")
        self._statusbar_var.set(f"Excel: {short_name} | {len(self._recipients):,} destinatarios")

        if loader.invalid_emails:
            self._log_box.append(f"   ⚠ Correos inválidos omitidos ({len(loader.invalid_emails)}):")
            for inv in loader.invalid_emails[:10]:
                self._log_box.append(f"     - {inv}")
            if len(loader.invalid_emails) > 10:
                self._log_box.append(f"     ... y {len(loader.invalid_emails) - 10} más.")

    def _load_html(self):
        """Abre diálogo para seleccionar archivo HTML."""
        path = filedialog.askopenfilename(
            title="Seleccionar archivo HTML",
            filetypes=[("HTML files", "*.html *.htm"), ("All files", "*.*")],
        )
        if not path:
            return
        self._content_path = path
        self._content_type = "html"
        self._content_path_var.set(f"🌐 HTML: {Path(path).name}")
        self._preview_btn.configure(state="normal")
        self._log_box.append(f"🌐 HTML seleccionado: {Path(path).name}")

    def _load_image(self):
        """Abre diálogo para seleccionar imagen."""
        path = filedialog.askopenfilename(
            title="Seleccionar imagen",
            filetypes=[("Image files", "*.jpg *.jpeg *.png"), ("All files", "*.*")],
        )
        if not path:
            return
        self._content_path = path
        self._content_type = "image"
        self._content_path_var.set(f"🖼 Imagen: {Path(path).name}")
        self._preview_btn.configure(state="normal")
        self._log_box.append(f"🖼 Imagen seleccionada: {Path(path).name}")

    # ─────────────────────────────────────────────
    #  VISTA PREVIA
    # ─────────────────────────────────────────────

    def _show_preview(self):
        """Muestra una ventana de vista previa del contenido seleccionado."""
        if not self._content_path:
            messagebox.showwarning("Vista previa", "No hay contenido seleccionado.")
            return

        # Cerrar ventana de preview anterior si existe
        if self._preview_window and self._preview_window.winfo_exists():
            self._preview_window.destroy()

        win = ctk.CTkToplevel(self)
        win.title("Vista Previa del Contenido")
        win.geometry("800x640")
        win.configure(fg_color=COLORS["surface"])
        win.grab_set()
        self._preview_window = win

        ctk.CTkLabel(
            win,
            text="Vista Previa del Correo",
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
            text_color=COLORS["text_primary"],
        ).pack(pady=(16, 4))

        ctk.CTkLabel(
            win,
            text=Path(self._content_path).name,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=COLORS["text_secondary"],
        ).pack(pady=(0, 12))

        if self._content_type == "html":
            self._preview_html(win)
        elif self._content_type == "image":
            self._preview_image(win)

        ctk.CTkButton(
            win,
            text="Cerrar",
            command=win.destroy,
            width=120,
            height=36,
            corner_radius=8,
            fg_color=COLORS["surface_elevated"],
            hover_color=COLORS["border"],
        ).pack(pady=12)

    def _preview_html(self, win: ctk.CTkToplevel):
        """Muestra el contenido HTML en el sistema (abre el navegador)."""
        try:
            webbrowser.open(f"file:///{self._content_path.replace(os.sep, '/')}")
        except Exception as e:
            pass

        # También mostramos el código fuente en la ventana
        try:
            with open(self._content_path, "r", encoding="utf-8") as f:
                html_src = f.read()
        except Exception as e:
            html_src = f"Error al leer el archivo: {e}"

        frame = ctk.CTkFrame(win, fg_color=COLORS["surface_card"], corner_radius=8)
        frame.pack(fill="both", expand=True, padx=16, pady=4)

        ctk.CTkLabel(
            frame,
            text="El archivo HTML se abrió en tu navegador predeterminado.\nVista del código fuente:",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=COLORS["accent"],
        ).pack(pady=(8, 4))

        txt = ctk.CTkTextbox(
            frame,
            font=ctk.CTkFont(family="Consolas", size=10),
            fg_color="#0A0A1A",
            text_color="#CE93D8",
            corner_radius=6,
        )
        txt.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        txt.insert("1.0", html_src[:8000] + ("\n... (archivo truncado en vista previa)" if len(html_src) > 8000 else ""))
        txt.configure(state="disabled")

    def _preview_image(self, win: ctk.CTkToplevel):
        """Muestra la imagen en la ventana de vista previa."""
        try:
            frame = ctk.CTkFrame(win, fg_color=COLORS["surface_card"], corner_radius=8)
            frame.pack(fill="both", expand=True, padx=16, pady=4)

            img = Image.open(self._content_path)
            # Escalar para caber en 760x500
            img.thumbnail((760, 500), Image.LANCZOS)
            ctk_img = ctk.CTkImage(light_image=img, dark_image=img, size=img.size)

            lbl = ctk.CTkLabel(frame, image=ctk_img, text="")
            lbl.pack(expand=True, pady=16)

            ctk.CTkLabel(
                frame,
                text=f"Tamaño: {img.size[0]}×{img.size[1]} px  |  {Path(self._content_path).stat().st_size // 1024} KB",
                font=ctk.CTkFont(family="Segoe UI", size=11),
                text_color=COLORS["text_secondary"],
            ).pack(pady=(0, 8))

        except Exception as e:
            ctk.CTkLabel(
                win,
                text=f"Error al mostrar la imagen:\n{e}",
                text_color=COLORS["danger"],
            ).pack(pady=20)

    # ─────────────────────────────────────────────
    #  PRUEBA DE CONEXIÓN
    # ─────────────────────────────────────────────

    def _test_connection(self):
        """Prueba la conexión SMTP en un hilo separado."""
        cfg = self._get_smtp_config()
        if not cfg:
            return

        self._statusbar_var.set("🔌 Probando conexión SMTP...")
        self._log_box.append("🔌 Probando conexión SMTP...")

        def do_test():
            dummy_sender = MailSender(
                smtp_config=cfg,
                mail_content=MailContent("", "html", ""),
                recipients=[],
                batch_size=1,
                wait_seconds=0,
            )
            ok, msg = dummy_sender.test_connection()
            self.after(0, lambda: self._on_test_result(ok, msg))

        threading.Thread(target=do_test, daemon=True).start()

    def _on_test_result(self, ok: bool, msg: str):
        if ok:
            self._log_box.append(f"✅ {msg}")
            self._statusbar_var.set(f"✅ {msg}")
            messagebox.showinfo("Conexión SMTP", msg)
        else:
            self._log_box.append(f"❌ {msg}")
            self._statusbar_var.set(f"❌ {msg}")
            messagebox.showerror("Error de Conexión", msg)

    # ─────────────────────────────────────────────
    #  VALIDACIÓN Y CONSTRUCCIÓN DE OBJETOS
    # ─────────────────────────────────────────────

    def _get_smtp_config(self) -> SMTPConfig | None:
        """Construye y valida la configuración SMTP desde la UI."""
        try:
            port = int(self._smtp_port.get().strip())
        except ValueError:
            messagebox.showerror("Error de configuración", "El puerto debe ser un número entero.")
            return None

        cfg = SMTPConfig(
            host=self._smtp_host.get().strip(),
            port=port,
            user=self._smtp_user.get().strip(),
            password=self._smtp_pass.get(),
            use_tls=self._tls_var.get(),
            use_ssl=self._ssl_var.get(),
        )
        ok, msg = cfg.validate()
        if not ok:
            messagebox.showerror("Error de configuración", msg)
            return None
        return cfg

    def _get_mail_content(self) -> MailContent | None:
        """Construye y valida el contenido del correo."""
        subject = self._subject_entry.get().strip()
        content = MailContent(
            subject=subject,
            content_type=self._content_type,
            content_path=self._content_path,
        )
        ok, msg = content.validate()
        if not ok:
            messagebox.showerror("Error de contenido", msg)
            return None
        return content

    def _get_batch_config(self) -> tuple[int, float] | None:
        """Obtiene y valida la configuración de lotes."""
        try:
            batch_size = int(self._batch_size_entry.get().strip())
            wait_seconds = float(self._wait_entry.get().strip())
            if batch_size <= 0 or wait_seconds < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror(
                "Error de configuración",
                "El tamaño del lote debe ser un entero positivo\ny la espera un número >= 0.",
            )
            return None
        return batch_size, wait_seconds

    # ─────────────────────────────────────────────
    #  ENVÍO
    # ─────────────────────────────────────────────

    def _start_sending(self):
        """Valida todo y lanza el proceso de envío."""
        if not self._recipients:
            messagebox.showwarning("Sin destinatarios", "Por favor carga un archivo Excel con destinatarios.")
            return

        cfg = self._get_smtp_config()
        if not cfg:
            return

        content = self._get_mail_content()
        if not content:
            return

        batch_cfg = self._get_batch_config()
        if not batch_cfg:
            return

        batch_size, wait_seconds = batch_cfg

        # Confirmación
        confirm = messagebox.askyesno(
            "Confirmar Envío Masivo",
            f"¿Confirmar el envío masivo?\n\n"
            f"• Destinatarios: {len(self._recipients):,}\n"
            f"• Asunto: {content.subject}\n"
            f"• Contenido: {Path(content.content_path).name}\n"
            f"• Servidor: {cfg.host}:{cfg.port}\n"
            f"• Lote: {batch_size} correos | Espera: {wait_seconds}s\n\n"
            f"Esta acción enviará correos reales.",
        )
        if not confirm:
            return

        # Bloquear UI
        self._set_sending_mode(True)
        self._log_box.clear()
        self._reset_stats(len(self._recipients))

        # Crear sender
        self._sender = MailSender(
            smtp_config=cfg,
            mail_content=content,
            recipients=self._recipients,
            batch_size=batch_size,
            wait_seconds=wait_seconds,
            on_progress=self._on_progress,
            on_log=self._on_log,
            on_finished=self._on_finished,
        )

        self._send_thread = threading.Thread(target=self._sender.run, daemon=True)
        self._send_thread.start()

    def _cancel_sending(self):
        """Solicita la cancelación del envío."""
        if self._sender:
            result = messagebox.askyesno(
                "Cancelar Envío",
                "¿Está seguro de que desea cancelar el envío masivo?\n"
                "Los correos del lote actual se interrumpirán.",
            )
            if result:
                self._sender.cancel()
                self._status_var.set("⚠️ Cancelando envío...")
                self._statusbar_var.set("⚠️ Cancelando envío...")

    # ─────────────────────────────────────────────
    #  CALLBACKS (llamados desde el hilo de envío)
    # ─────────────────────────────────────────────

    def _on_progress(self, stats: SendStats):
        """Actualiza la UI con el progreso actual (thread-safe con after)."""
        self.after(0, lambda: self._update_progress_ui(stats))

    def _on_log(self, msg: str):
        """Agrega una línea al log (thread-safe)."""
        self.after(0, lambda: self._log_box.append(msg))

    def _on_finished(self, stats: SendStats, cancelled: bool = False):
        """Callback cuando el envío finaliza (thread-safe)."""
        self.after(0, lambda: self._handle_finished(stats, cancelled))

    def _update_progress_ui(self, stats: SendStats):
        """Actualiza todos los elementos de progreso en la UI."""
        self._card_sent.set_value(stats.sent)
        self._card_failed.set_value(stats.failed)
        self._card_pending.set_value(stats.pending)
        self._progress_bar.set(stats.progress_pct)
        pct = int(stats.progress_pct * 100)
        self._progress_pct_var.set(f"{pct}%")
        self._status_var.set(
            f"Enviados: {stats.sent} | Fallidos: {stats.failed} | Pendientes: {stats.pending}"
        )

    def _handle_finished(self, stats: SendStats, cancelled: bool):
        """Maneja el fin del proceso de envío."""
        self._set_sending_mode(False)
        self._update_progress_ui(stats)

        if cancelled:
            self._status_var.set(f"⚠️ Envío cancelado. Enviados: {stats.sent} | Fallidos: {stats.failed}")
            self._statusbar_var.set("Envío cancelado por el usuario.")
            messagebox.showwarning(
                "Envío Cancelado",
                f"El envío fue cancelado.\n\nEnviados: {stats.sent}\nFallidos: {stats.failed}\nPendientes: {stats.pending}",
            )
        else:
            self._status_var.set(f"✅ Envío completado. Enviados: {stats.sent} | Fallidos: {stats.failed}")
            self._statusbar_var.set(f"Envío completado. {stats.sent:,} enviados / {stats.failed} fallidos.")
            messagebox.showinfo(
                "Envío Completado",
                f"El proceso de envío ha finalizado.\n\n"
                f"✅ Enviados: {stats.sent:,}\n"
                f"❌ Fallidos: {stats.failed}\n"
                f"📊 Total procesado: {stats.sent + stats.failed:,}",
            )

    # ─────────────────────────────────────────────
    #  HELPERS DE UI
    # ─────────────────────────────────────────────

    def _set_sending_mode(self, sending: bool):
        """Habilita/deshabilita controles según si se está enviando."""
        self._start_btn.configure(state="disabled" if sending else "normal")
        self._cancel_btn.configure(state="normal" if sending else "disabled")
        if sending:
            self._status_var.set("📤 Enviando correos...")
            self._statusbar_var.set("Envío en progreso...")

    def _reset_stats(self, total: int):
        """Reinicia las tarjetas de estadísticas."""
        self._card_total.set_value(total)
        self._card_sent.set_value(0)
        self._card_failed.set_value(0)
        self._card_pending.set_value(total)
        self._progress_bar.set(0)
        self._progress_pct_var.set("0%")
