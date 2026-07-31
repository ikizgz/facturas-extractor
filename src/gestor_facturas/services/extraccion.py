"""Módulo de extracción de texto y tablas de archivos PDF usando pdfplumber."""

from __future__ import annotations

import logging
from pathlib import Path

import pdfplumber

# Logger específico para este módulo
logger = logging.getLogger(__name__)


def extraer_texto_pdf(pdf_path: str | Path) -> str:
    """Extrae todo el texto plano de un archivo PDF respetando el layout

    utilizando pdfplumber, evitando excepciones ciegas y usando logging.
    """
    path = Path(pdf_path)
    texto_completo: list[str] = []

    try:
        with pdfplumber.open(path) as pdf:
            for pagina in pdf.pages:
                texto_pagina = pagina.extract_text(layout=True)
                if texto_pagina:
                    texto_completo.append(texto_pagina)
        return "\n".join(texto_completo)

    except (FileNotFoundError, PermissionError) as e:
        logger.error(f"Error de acceso al archivo PDF {path.name}: {e}")
        return ""
    except (
        Exception
    ):  # Captura defensiva externa si pdfplumber lanza un fallo interno de parseo
        logger.exception(
            "Error inesperado al extraer texto con pdfplumber de %s", path.name
        )
        return ""
