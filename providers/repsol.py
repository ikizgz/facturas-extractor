# -*- coding: utf-8 -*-
from __future__ import annotations

import re
from typing import List

from .base import ProviderParser
from .common import Row, norm_cif, norm_num

REPSOL_NUM_RE = re.compile(r"N[ºo]\s*Factura\s*[:#]?\s*([0-9/]+)", re.IGNORECASE)
REPSOL_DATE_RE = re.compile(r"Fecha\s*[:#]?\s*(\d{2}/\d{2}/\d{4})", re.IGNORECASE)
REPSOL_BASE_RE = re.compile(
    r"Importe\s+del\s+producto\s*\(\s*Base\s+Imponible\s*\)\s*([\d.,]+)", re.IGNORECASE
)
REPSOL_IVA_RE = re.compile(
    r"IVA\s*\d{1,2}[.,]\d{2}%\s*de\s*[\d.,]+\s*€\s*([\d.,]+)", re.IGNORECASE
)
REPSOL_TOTAL_RE = re.compile(
    r"TOTAL\s+FACTURA\s+EUROS[^\d]*([\d.,]+)\s*€", re.IGNORECASE
)


class RepsolParser(ProviderParser):
    name = "REPSOL"

    def detect(self, text: str) -> bool:
        up = text.upper()
        return ("REPSOL SOLUCIONES ENERGETICAS" in up) or ("TOTAL FACTURA EUROS" in up)

    def parse(self, text: str, path) -> List[Row]:
        mnum = REPSOL_NUM_RE.search(text)
        number = mnum.group(1) if mnum else None
        mdate = REPSOL_DATE_RE.search(text)
        fecha = mdate.group(1) if mdate else None
        base = iva = total = None
        mb = REPSOL_BASE_RE.search(text)
        mi = REPSOL_IVA_RE.search(text)
        mt = REPSOL_TOTAL_RE.search(text)
        if mb:
            base = norm_num(mb.group(1))
        if mi:
            iva = norm_num(mi.group(1))
        if mt:
            total = norm_num(mt.group(1))
        if base and iva and not total:
            total = round(base + iva, 2)
        pct = round(iva / base, 6) if base and iva and base > 0 else None
        return [
            {
                "fecha_factura": fecha,
                "numero_factura": number or path.stem,
                "empresa": "REPSOL SOLUCIONES ENERGETICAS, S.A.",
                "CIF": norm_cif("A80298839"),
                "importe_base": base,
                "%IVA": pct,
                "IVA": iva,
                "importe_total": total,
                "Notas": "",
            }
        ]
