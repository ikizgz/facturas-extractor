"""Módulo core que define el pipeline de procesamiento y orquestación de facturas."""

from __future__ import annotations

import logging
from pathlib import Path

from gestor_facturas.core.config import LONGITUD_MINIMA_TEXTO
from gestor_facturas.providers import PROVEEDORES_REGISTRADOS
from gestor_facturas.services.extraccion import extraer_texto_pdf
from gestor_facturas.services.ocr import realizar_ocr_pdf

# Logger específico para el pipeline
logger = logging.getLogger(__name__)


def procesar_pdf(pdf_path: str | Path) -> dict | None:
    """Ejecuta el pipeline completo para un PDF dado:

    1. Extrae el texto (nativo o mediante OCR si es necesario).
    2. Identifica el proveedor adecuado mediante el registro de parsers.
    3. Extrae y devuelve los datos normalizados listos para el Excel.
    """
    path = Path(pdf_path)
    logger.info(f"Iniciando pipeline de procesamiento para: {path.name}")

    # Paso 1: Intentar extracción nativa con pdfplumber
    texto = extraer_texto_pdf(path)
    texto_limpio = texto.strip()

    # Paso 2: Validación inteligente (vacío o texto insuficiente / pobre)
    if not texto_limpio:
        logger.warning(
            f"El PDF {path.name} no contiene texto nativo. Activando OCR..."
        )
        texto = realizar_ocr_pdf(path)
    elif len(texto_limpio) < LONGITUD_MINIMA_TEXTO:
        logger.warning(
            f"El texto extraído de forma nativa en {path.name} es muy escaso "
            f"({len(texto_limpio)} caracteres). Mínimo requerido: {LONGITUD_MINIMA_TEXTO}. Activando OCR..."
        )
        texto = realizar_ocr_pdf(path)
    else:
        logger.info(
            f"Texto nativo extraído y validado correctamente para {path.name} "
            f"({len(texto_limpio)} caracteres)."
        )

    if not texto.strip():
        logger.error(f"No se pudo extraer ningún texto útil del archivo {path.name}")
        return None

    # Paso 3: Buscar el proveedor adecuado recorriendo el registro
    proveedor_encontrado = None
    for parser in PROVEEDORES_REGISTRADOS:
        if parser.puede_procesar(texto):
            proveedor_encontrado = parser
            break

    if not proveedor_encontrado:
        logger.warning(
            f"No se encontró ningun parser registrado capaz de procesar el archivo: {path.name}"
        )
        # Opcional: podrías devolver un diccionario genérico o por defecto si lo deseas
        return None

    # Paso 4: Extraer datos estructurados con el parser específico
    logger.info(f"Proveedor identificado para {path.name}: {proveedor_encontrado.__class__.__name__}")
    datos_factura = proveedor_encontrado.extraer_datos(texto, path)
    
    return datos_factura
