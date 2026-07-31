"""Módulo core que define el pipeline de procesamiento de facturas."""

from __future__ import annotations

import logging
from pathlib import Path

from gestor_facturas.core.config import (
    LONGITUD_MINIMA_TEXTO,  # Umbral mínimo de caracteres para considerar un texto nativo como válido y útil
)
from gestor_facturas.services.extraccion import extraer_texto_pdf
from gestor_facturas.services.ocr import realizar_ocr_pdf

# Logger específico para el pipeline
logger = logging.getLogger(__name__)


def procesar_pdf(pdf_path: str | Path) -> str:
    """Ejecuta el pipeline de extracción de texto para un PDF dado.

    1. Intenta extracción nativa con pdfplumber.
    2. Comprueba si el texto está vacío o es inferior al umbral mínimo (texto corrupto o insuficiente).
    3. Si no es válido, recurre automáticamente a OCR.
    """
    path = Path(pdf_path)
    logger.info(f"Iniciando pipeline de procesamiento para: {path.name}")

    # Paso 1: Intentar extracción nativa
    texto = extraer_texto_pdf(path)
    texto_limpio = texto.strip()

    # Paso 2: Validación inteligente (vacío o texto insuficiente / pobre)
    if not texto_limpio:
        logger.warning(f"El PDF {path.name} no contiene texto nativo. Activando OCR...")
        texto = realizar_ocr_pdf(path)
    elif len(texto_limpio) < LONGITUD_MINIMA_TEXTO:
        logger.warning(
            f"El texto extraído de forma nativa en {path.name} es muy escaso "
            f"({len(texto_limpio)} caracteres). Posible PDF escaneado con mala capa de texto. Activando OCR..."
        )
        texto = realizar_ocr_pdf(path)
    else:
        logger.info(
            f"Texto nativo extraído y validado correctamente para {path.name} "
            f"({len(texto_limpio)} caracteres)."
        )

    return texto
