# -*- coding: utf-8 -*-
# providers/salvadorescoda.py
from __future__ import annotations

import re
from typing import List

from .base import ProviderParser
from .common import Row, norm_num, parse_date_text


class SalvadorEscodaParser(ProviderParser):
    name = "SALVADOR ESCODA S.A."

    def detect(self, text: str) -> bool:
        # Detección ultra-segura
        return "SALVADOR ESCODA" in text.upper() or "A-08710006" in text.upper()

    def parse(self, text: str, path) -> List[Row]:
        # Normalización conservando los espacios clave
        raw_text = " ".join(text.split())

        # 1. Número de factura: Busca "FACTURA" y coge el número de 7 dígitos siguiente
        mnum = re.search(r"FACTURA\s+(\d{7})", raw_text)
        number = mnum.group(1) if mnum else path.stem

        fecha = parse_date_text(text)

        # 2. Extracción de Totales (El truco está aquí)
        # Buscamos la secuencia completa en el raw_text que incluye los importes.
        # Captura: BASE IMPONIBLE, luego el IVA, luego el TOTAL.
        # Usamos un regex que se salta toda la basura entre medio.
        totals_pattern = r"BASE IMPONIBLE.*?([0-9.,]+)\s+21,00\s+% IVA.*?([0-9.,]+)\s+EUR\s+([0-9.,]+)"

        m_totals = re.search(totals_pattern, raw_text, re.IGNORECASE)

        if m_totals:
            base_imp = norm_num(m_totals.group(1))
            iva_total = norm_num(m_totals.group(2))
            total_factura = norm_num(m_totals.group(3))
        else:
            # Fallback si el formato es distinto
            base_imp = None
            iva_total = None
            total_factura = None

        rows: List[Row] = []
        rows.append(
            {
                "fecha_factura": fecha,
                "numero_factura": number,
                "empresa": "SALVADOR ESCODA S.A.",
                "CIF": "A08710006",
                "importe_base": base_imp,
                "%IVA": 0.21,
                "IVA": iva_total,
                "importe_total": total_factura,
                "Notas": "Compra material Salvador Escoda",
            }
        )

        return rows
