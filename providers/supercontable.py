# -*- coding: utf-8 -*-

from __future__ import annotations

import re
from typing import List

from .base import ProviderParser
from .common import Row, norm_num, parse_date_text


class SupercontableParser(ProviderParser):
    name = "SUPERCONTABLE"

    def detect(self, text: str) -> bool:
        up = text.upper()
        return ("SUPERCONTABLE" in up) or ("RCR PROYECTOS DE SOFTWARE" in up)

    def parse(self, text: str, path) -> List[Row]:
        up = text.upper()
        mnum = re.search(r"FACTURA\s+(PO\d+/\d+)", up)
        number = mnum.group(1) if mnum else None
        fecha = parse_date_text(text)
        m = re.search(
            r"\b(\d{1,3}[\d.,]*)\s+(\d{1,3}[\d.,]*)\s+21\s*%\s+(\d{1,3}[\d.,]*)\s+(\d{1,3}[\d.,]*)\s*EUR",
            up,
        )
        base = iva = total = None
        if m:
            base = norm_num(m.group(2))
            iva = norm_num(m.group(3))
            total = norm_num(m.group(4))
        if base and iva and not total:
            total = round(base + iva, 2)
        return [
            {
                "fecha_factura": fecha,
                "numero_factura": number or path.stem,
                "empresa": "RCR PROYECTOS DE SOFTWARE, S.L.U.",
                "CIF": None,
                "importe_base": base,
                "%IVA": 0.21 if base and iva else None,
                "IVA": iva,
                "importe_total": total,
                "Notas": "",
            }
        ]
