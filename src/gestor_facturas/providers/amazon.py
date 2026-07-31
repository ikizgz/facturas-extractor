"""Parser específico para facturas de Amazon y sus diferentes proveedores/emisores."""

from __future__ import annotations

import logging
import re
from pathlib import Path

from gestor_facturas.providers.base import ProveedorBase

# Logger específico para el parser de Amazon
logger = logging.getLogger(__name__)

# Diccionario de entidades conocidas de Amazon por su NIF/IVA
KNOWN_AMAZON_ENTITIES = {
    "N0186600C": "Amazon Business EU S.à.r.l",
    "W0264006H": "Amazon Business EU S.à.r.l, Sucursal en España",
    "DE814584193": "Amazon EU S.à r.l., Niederlassung Deutschland",
    "FR12487773327": "Amazon EU S.à r.l., Succursale Française",
    "IT08973230967": "Amazon EU S.à r.l., Succursale Italiana",
    "W0184081H": "Amazon EU S.à r.l., Sucursal en España",
    "PL5262907815": "Amazon EU S.à r.l., Sucursal en Polonia",
}


def _normalizar_numero(valor_str: str | None) -> float:
    """Convierte una cadena numérica con formato europeo (ej.

    '8,16' o '1.234,56') a float.
    """
    if not valor_str:
        return 0.0
    try:
        # Limpiar símbolos de moneda y espacios
        limpio = (
            valor_str.replace("€", "")
            .replace(" ", "")
            .replace(".", "")
            .replace(",", ".")
        )
        return float(limpio)
    except ValueError:
        return 0.0


class AmazonParser(ProveedorBase):
    """Implementación del parser para facturas del ecosistema Amazon."""

    def puede_procesar(self, texto_pdf: str) -> bool:
        """Determina si la factura pertenece a Amazon analizando palabras clave."""
        t = texto_pdf.upper()
        return "AMAZON" in t or "AMAZON BUSINESS" in t

    def extraer_datos(self, texto_pdf: str, archivo_pdf: str | Path) -> dict:
        """Extrae y normaliza los datos de una factura de Amazon."""
        path = Path(archivo_pdf)
        raw_text = " ".join(texto_pdf.split())
        logger.info(f"Procesando factura de Amazon con parser específico: {path.name}")

        # 1. Extracción de fecha (compatible con formatos tipo "20 julio 2026" o "19 de julio de 2026")
        fecha_str = ""
        m_fecha = re.search(
            r"Fecha de la factura.*?Fecha de la entrega[^\d]*(\d{1,2}\s+(?:de\s+)?[a-zA-Z]+\s+(?:de\s+)?\d{4})",
            raw_text,
            re.IGNORECASE,
        )
        if not m_fecha:
            m_fecha = re.search(
                r"Fecha de la factura[^\d]*(\d{1,2}\s+(?:de\s+)?[a-zA-Z]+\s+(?:de\s+)?\d{4})",
                raw_text,
                re.IGNORECASE,
            )

        if m_fecha:
            fecha_str = m_fecha.group(1).strip()
        else:
            # Fallback buscando cualquier fecha legible en el texto
            todas_fechas = re.findall(
                r"(\d{1,2}\s+(?:de\s+)?[a-zA-Z]+\s+(?:de\s+)?\d{4})", texto_pdf
            )
            if todas_fechas:
                fecha_str = todas_fechas[0]

        # 2. Número de factura
        mnum = re.search(r"Número de la factura\s*([A-Z0-9]+)", raw_text, re.IGNORECASE)
        numero_factura = mnum.group(1) if mnum else path.stem

        # 3. Vendedor y NIF/CIF exactos
        cif = ""
        proveedor = ""

        vendido_match = re.search(
            r"Vendido por\s+([A-Za-z0-9àèìòùÀÈÌÒÙáéíóúÁÉÍÓÚñÑäëïöüÄËÏÖÜ\.\s,\-_]+?)(?=\s+IVA\s+[A-Z0-9]+|\s+Fecha de la factura|$)",
            raw_text,
            re.IGNORECASE,
        )
        if vendido_match:
            proveedor = vendido_match.group(1).strip()

        iva_vend_match = re.search(
            r"Vendido por\s+.*?IVA\s*([A-Z0-9]+)", raw_text, re.IGNORECASE
        )
        if iva_vend_match:
            cif = iva_vend_match.group(1).strip()

        if cif in KNOWN_AMAZON_ENTITIES:
            proveedor = KNOWN_AMAZON_ENTITIES[cif]
        elif not proveedor:
            proveedor = "Amazon Business EU S.à.r.l, Sucursal en España"
            cif = "W0264006H"

        # 4. Importes e Impuestos (Base, Tasa IVA, Importe IVA, Total)
        importe_base = 0.0
        tasa_iva = 0.0
        importe_iva = 0.0
        importe_total = 0.0

        # Buscar en la tabla de desglose de IVA (ej: "21.0% | 4,95 € | 1,04 €" o "0% | 8,16 € | 0,00 €")
        tabla_iva_match = re.search(
            r"IVA\s*(\d{1,2}(?:[.,]\d{1,2})?)\s*%\s*(-?\d+[.,]\d{2})\s*€?\s*(-?\d+[.,]\d{2})\s*€?",
            raw_text,
            re.IGNORECASE,
        )
        if tabla_iva_match:
            tasa_iva = _normalizar_numero(tabla_iva_match.group(1))
            importe_base = _normalizar_numero(tabla_iva_match.group(2))
            importe_iva = _normalizar_numero(tabla_iva_match.group(3))

        # Buscar importe total de la factura
        tot_match = re.search(
            r"Total de la factura\s*(-?\d+[.,]\d{2})\s*€", raw_text, re.IGNORECASE
        )
        if not tot_match:
            tot_match = re.search(
                r"Total\s*(-?\d+[.,]\d{2})\s*€", raw_text, re.IGNORECASE
            )

        if tot_match:
            importe_total = _normalizar_numero(tot_match.group(1))
        else:
            todos_totales = re.findall(
                r"Total[^\d]*(-?\d+[.,]\d{2})\s*€", raw_text, re.IGNORECASE
            )
            if todos_totales:
                importe_total = _normalizar_numero(todos_totales[-1])

        # Coherencia matemática si faltan algunos campos
        if tasa_iva == 0.0 and importe_base == 0.0 and importe_total > 0:
            importe_base = importe_total
            importe_iva = 0.0

        if importe_total == 0.0 and importe_base > 0:
            importe_total = round(importe_base + importe_iva, 2)

        if importe_iva == 0.0 and importe_base > 0 and importe_total > importe_base:
            importe_iva = round(importe_total - importe_base, 2)

        # Devolver diccionario con las cabeceras exactas solicitadas para el Excel
        return {
            "fecha": fecha_str,
            "numero_factura": numero_factura,
            "proveedor": proveedor,
            "nif_cif": cif,
            "importe_base": importe_base,
            "tasa_iva": tasa_iva,
            "importe_iva": importe_iva,
            "importe_total": importe_total,
        }
