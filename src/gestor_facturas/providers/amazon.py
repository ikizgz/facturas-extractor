"""Parser específico para la extracción de datos en facturas y notas de crédito de Amazon."""

from __future__ import annotations

import re
from pathlib import Path

from gestor_facturas.providers.base import ProveedorBase
from gestor_facturas.services.common import _normalizar_numero, _parsear_fecha


class AmazonParser(ProveedorBase):
    """Parser para procesar documentos de Amazon (Facturas y Notas de Crédito)."""

    def puede_procesar(self, raw_text: str) -> bool:
        """Determina si el texto pertenece a una factura o nota de crédito de Amazon."""
        texto = raw_text.lower()

        # Identificamos que sea de Amazon y que corresponda a un documento fiscal válido
        es_amazon = "amazon" in texto
        es_documento_valido = "factura" in texto or "nota de crédito" in texto

        return bool(es_amazon and es_documento_valido)

    def extraer_datos(self, raw_text: str, pdf_path: str | Path) -> dict:
        """Extrae los campos normalizados de una factura o nota de crédito de Amazon."""

        # 1. Número de Factura / Nota de Crédito
        # Buscamos primero si es una nota de crédito, si no, buscamos factura estándar
        num_match = re.search(
            r"Número de la nota de crédito\s*([A-Z0-9]+)", raw_text, re.IGNORECASE
        )
        if not num_match:
            num_match = re.search(
                r"Número de la factura\s*([A-Z0-9]+)", raw_text, re.IGNORECASE
            )

        numero_factura = num_match.group(1) if num_match else "DESCONOCIDO"

        # 2. Fecha (Buscamos fecha de factura o de nota de crédito)
        fecha_match = re.search(
            r"Fecha de la nota de crédito[:\s]*(\d{1,2}\s+(?:de\s+)?[a-zA-Záéíóú]+\s+\d{4})",
            raw_text,
            re.IGNORECASE,
        )
        if not fecha_match:
            fecha_match = re.search(
                r"Fecha de la factura[/\s]*(?:Fecha de la entrega)?[:\s]*(\d{1,2}\s+(?:de\s+)?[a-zA-Záéíóú]+\s+\d{4})",
                raw_text,
                re.IGNORECASE,
            )

        fecha = _parsear_fecha(fecha_match.group(1)) if fecha_match else ""

        # 3. Proveedor y NIF / CIF del emisor
        # Intentamos extraer el vendedor real tras "Vendido por"
        proveedor = "Amazon"
        vendedor_match = re.search(r"Vendido por\s*([^\n\r]+)", raw_text, re.IGNORECASE)
        if vendedor_match:
            # Limpiamos posibles espacios o saltos de línea sobrantes
            vendedor_candidato = vendedor_match.group(1).strip()
            if vendedor_candidato:
                proveedor = vendedor_candidato

        # Buscamos patrones de NIF/IVA de proveedor habituales
        nif_match = re.search(
            r"(?:IVA|NIF|CIF)[:\s]*([A-Z]{2}[A-Z0-9]{8,12})", raw_text, re.IGNORECASE
        )
        nif_cif = nif_match.group(1) if nif_match else "DESCONOCIDO"

        # 4. Importes e Impuestos (Soportando valores negativos para abonos)
        importe_base = 0.0
        tasa_iva = 0.0
        importe_iva = 0.0
        importe_total = 0.0

        # Búsqueda de la línea de desglose de IVA (ej: "0% -8,39 € 0,00 €" o "21.0% 4,95 € 1,04 €")
        tabla_iva_match = re.search(
            r"(\d{1,2}(?:[.,]\d{1,2})?)\s*%\s*(-?\d+[.,]\d{2})\s*€?\s*(-?\d+[.,]\d{2})\s*€?",
            raw_text,
            re.IGNORECASE,
        )

        if tabla_iva_match:
            t_iva_str = tabla_iva_match.group(1).replace(",", ".")
            tasa_iva = float(t_iva_str)
            importe_base = _normalizar_numero(tabla_iva_match.group(2))
            importe_iva = _normalizar_numero(tabla_iva_match.group(3))

        # Buscar importe total (o total pendiente / reembolso en notas de crédito)
        tot_match = re.search(
            r"Total pendiente\s*(-?\d+[.,]\d{2})\s*€", raw_text, re.IGNORECASE
        )
        if not tot_match:
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

        # Coherencia matemática si faltan campos
        if tasa_iva == 0.0 and importe_base == 0.0 and importe_total != 0:
            importe_base = importe_total
            importe_iva = 0.0

        if importe_total == 0.0 and importe_base != 0:
            importe_total = round(importe_base + importe_iva, 2)

        if (
            importe_iva == 0.0
            and importe_base != 0
            and abs(importe_total) > abs(importe_base)
        ):
            importe_iva = round(importe_total - importe_base, 2)

        # 5. Estructura exacta requerida para el Excel unificado
        return {
            "fecha": fecha,
            "numero_factura": numero_factura,
            "proveedor": proveedor,
            "nif_cif": nif_cif,
            "importe_base": importe_base,
            "tasa_iva": tasa_iva,
            "importe_iva": importe_iva,
            "importe_total": importe_total,
        }
