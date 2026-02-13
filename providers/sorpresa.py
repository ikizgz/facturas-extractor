# -*- coding: utf-8 -*-

from __future__ import annotations

import re
from typing import List

from .base import ProviderParser
from .common import Row, norm_cif, norm_num


class SorpresaParser(ProviderParser):
    name = "SORPRESA"

    def detect(self, text: str) -> bool:
        up = text.upper()
        return ("SORPRESA HOGAR" in up) or ("XIAOJIE WANG" in up)

    def parse(self, text: str, path) -> List[Row]:
        up = text.upper()
        mnum = re.search(r"N\s*\*\s*FAC\s*[:#]?\s*(\d{3,12})", up)
        number = mnum.group(1) if mnum else None
        m = re.search(
            r"TOTAL\s*:\s*(\d{1,2}(?:[.,]\d{1,2})?)%\s*:\s*([0-9][0-9.,]*)\s*(\d{1,2}(?:[.,]\d{1,2})?)%\s*:\s*([0-9][0-9.,]*)\s*([0-9][0-9.,]*)",
            up,
        )
        base = iva = total = pct = None
        if m:
            pct = float(m.group(1).replace(",", ".")) / 100.0
            base = norm_num(m.group(2))
            iva = norm_num(m.group(4))
            total = norm_num(m.group(5))
        if base and iva and not total:
            total = round(base + iva, 2)
        return [
            {
                "fecha_factura": None,
                "numero_factura": number or path.stem,
                "empresa": "SORPRESA HOGAR",
                "CIF": norm_cif("X6526242S"),
                "importe_base": base,
                "%IVA": pct,
                "IVA": iva,
                "importe_total": total,
                "Notas": "",
            }
        ]
