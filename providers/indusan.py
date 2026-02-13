# -*- coding: utf-8 -*-
from __future__ import annotations

import re
from typing import List

from .base import ProviderParser
from .common import Row, norm_cif, norm_num, parse_date_text


class IndusanParser(ProviderParser):
    name = "INDUSAN"

    def detect(self, text: str) -> bool:
        up = text.upper()
        return ("INDUSTRIAS REUNIDAS SANITARIAS" in up) or ("INDUSAN" in up)

    def parse(self, text: str, path) -> List[Row]:
        up = text.upper()
        mnum = re.search(r"FACTURA[^]*?(\d{3,})", up)
        number = mnum.group(1) if mnum else None
        fecha = parse_date_text(text)
        mbase = re.search(r"BASE\s+IMPONIBLE[^]*?([0-9][0-9.,]*)", up)
        miva = re.search(r"IVA\s*%\s*21[^]*?([0-9][0-9.,]*)", up)
        mtotal = re.search(r"TOTAL\s+FACTURA[^]*?([0-9][0-9.,]*)", up)
        base = norm_num(mbase.group(1)) if mbase else None
        iva = norm_num(miva.group(1)) if miva else None
        total = norm_num(mtotal.group(1)) if mtotal else None
        if base and iva and not total:
            total = round(base + iva, 2)
        return [
            {
                "fecha_factura": fecha,
                "numero_factura": number or path.stem,
                "empresa": "INDUSTRIAS REUNIDAS SANITARIAS S.L.",
                "CIF": norm_cif("B50040005"),
                "importe_base": base,
                "%IVA": 0.21 if base and iva else None,
                "IVA": iva,
                "importe_total": total,
                "Notas": "",
            }
        ]
