"""Módulo de reconocimiento óptico de caracteres (OCR) usando pytesseract y pdf2image."""

from __future__ import annotations

import logging
from pathlib import Path

import pytesseract
from pdf2image import convert_from_path

# Logger específico para este módulo
logger = logging.getLogger(__name__)


def realizar_ocr_pdf(pdf_path: str | Path, idioma: str = "spa") -> str:
    """Convierte las páginas de un PDF en imágenes y aplica OCR con Tesseract

    para extraer el texto de documentos escaneados.
    """
    path = Path(pdf_path)
    texto_completo: list[str] = []

    try:
        logger.info(f"Iniciando proceso OCR para el archivo: {path.name}")

        # Convertir PDF a lista de imágenes (requiere poppler instalado en el sistema)
        imagenes = convert_from_path(path)

        for i, imagen in enumerate(imagenes, start=1):
            texto_pagina = pytesseract.image_to_string(imagen, lang=idioma)
            if texto_pagina.strip():
                texto_completo.append(texto_pagina)
            else:
                logger.warning(
                    f"El OCR no ha detectado texto en la página {i} de {path.name}"
                )

        return "\n".join(texto_completo)

    except (FileNotFoundError, PermissionError) as e:
        logger.error(f"Error de acceso al archivo durante el OCR {path.name}: {e}")
        return ""
    except Exception:
        logger.exception("Error crítico al procesar el archivo %s", path.name)
        return ""
