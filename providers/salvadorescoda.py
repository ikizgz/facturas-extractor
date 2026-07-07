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
        t = text.upper()
        return any(
            x in t for x in ["SALVADOR ESCODA", "ESCODA", "A08710006", "A-08710006"]
        )

    def parse(self, text: str, path) -> List[Row]:
        raw_text = " ".join(text.split())

        # 1. Extracción de datos básicos
        mnum = re.search(r"FACTURA\s*(\d{7})", raw_text)
        number = mnum.group(1) if mnum else path.stem
        fecha = parse_date_text(text)

        # 2. Captura del % IVA
        iva_pct_match = re.search(
            r"([0-9]+[,.][0-9]{1,2})\s*%\s*IVA", raw_text, re.IGNORECASE
        )
        # 1. Extraemos el valor usando norm_num
        valor_extraido = norm_num(iva_pct_match.group(1)) if iva_pct_match else None
        # 2. Asignamos con una guarda (si es None, usamos 0.21)
        iva_pct = (valor_extraido / 100) if valor_extraido is not None else 0.21

        # 3. Captura de importes (Base, IVA, Total)
        totals_match = re.search(
            r"BASE\s+IMPONIBLE.*?([0-9.,]+)\s+([0-9.,]+)\s+EUR\s+([0-9.,]+)",
            raw_text,
            re.IGNORECASE,
        )

        if totals_match:
            base_imp = norm_num(totals_match.group(1))
            iva_val = norm_num(totals_match.group(2))
            total_factura = norm_num(totals_match.group(3))
        else:
            base_imp, iva_val, total_factura = None, None, None

        # 4. Lógica de notas (solo si hay algo especial)
        nota = ""
        # Si el OCR nos ha dado una lectura pobre o sospechosa, lo marcamos
        if len(text) < 500:
            nota = "OCR: Texto corto o baja calidad"

        rows: List[Row] = []
        rows.append(
            {
                "fecha_factura": fecha,
                "numero_factura": number,
                "empresa": "SALVADOR ESCODA S.A.",
                "CIF": "A08710006",
                # Aseguramos que son float para que Excel los trate como moneda
                "importe_base": float(base_imp) if base_imp else 0.0,
                "%IVA": float(iva_pct),
                "IVA": float(iva_val) if iva_val else 0.0,
                "importe_total": float(total_factura) if total_factura else 0.0,
                "Notas": nota,
            }
        )

        return rows
