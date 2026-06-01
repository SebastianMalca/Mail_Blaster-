"""
Punto de entrada principal de Mail Blaster Institucional.
"""

import sys
import os

# Asegurar que el directorio del proyecto esté en el path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app.gui import MailBlasterApp


def main():
    app = MailBlasterApp()
    app.mainloop()


if __name__ == "__main__":
    main()
