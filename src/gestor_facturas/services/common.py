"""Funciones auxiliares y comunes para el procesamiento y normalización de datos."""

from __future__ import annotations

import logging

logger = logging.getLogger(__name__)


def _normalizar_numero(valor_str: str | None) -> float:
    """Convierte una cadena numérica con formato europeo (ej.

    '8,16' o '1.234,56') a un float de Python.
    """
    if not valor_str:
        return 0.0
    try:
        # Limpiar símbolos de moneda, espacios y adaptar separadores de miles/decimales
        limpio = (
            str(valor_str)
            .replace("€", "")
            .replace(" ", "")
            .replace(".", "")
            .replace(",", ".")
        )
        return float(limpio)
    except (ValueError, TypeError):
        logger.warning(f"No se pudo normalizar el valor numérico: '{valor_str}'")
        return 0.0


def _parsear_fecha(texto_fecha: str) -> str:
    """Convierte texto de fecha en español (ej: '24 julio 2026') a formato estándar YYYY-MM-DD."""
    meses = {
        "enero": "01",
        "febrero": "02",
        "marzo": "03",
        "abril": "04",
        "mayo": "05",
        "junio": "06",
        "julio": "07",
        "agosto": "08",
        "septiembre": "09",
        "octubre": "10",
        "noviembre": "11",
        "diciembre": "12",
    }
    if not texto_fecha:
        return ""

    try:
        limpio = texto_fecha.lower().replace("de", "").strip()
        partes = limpio.split()
        if len(partes) >= 3:
            dia = partes[0].zfill(2)
            mes_txt = partes[1]
            anio = partes[2]
            mes = meses[mes_txt]  # KeyError si el mes no está en el diccionario
            return f"{anio}-{mes}-{dia}"
    except (KeyError, IndexError, AttributeError):
        logger.warning(f"No se pudo parsear la fecha: '{texto_fecha}'")

    return texto_fecha  # Devuelve el texto original si falla la conversión
