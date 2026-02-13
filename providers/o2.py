# -*- coding: utf-8 -*-
from __future__ import annotations

import re
from typing import List

from .base import ProviderParser
from .common import Row, norm_cif, norm_num, parse_date_text


class O2Parser(ProviderParser):
    name = "O2"

    def detect(self, text: str) -> bool:
        up = text.upper()
        return ("TELEFÓNICA DE ESPAÑA" in up) or ("FACTURA NÚM" in up) or ("O2" in up)

    def parse(self, text: str, path) -> List[Row]:
        up = text.upper()
        mnum = re.search(r"FACTURA\s+N[ÚU]M\s*[:#]?\s*([A-Z0-9]+)", up)
        number = mnum.group(1) if mnum else None
        fecha = parse_date_text(text)
        mbase = re.search(r"BASE\s+IMPONIBLE\s*([0-9][0-9.,]*)\s*€", up)
        miva = re.search(
            r"IVA\s*\(\s*21\.?00\s*%\s*\)\s*sobre\s*[0-9][0-9.,]*\s*€\s*([0-9][0-9.,]*)\s*€",
            up,
        )
        mtotal = re.search(r"TOTAL\s+FACTURA\s*([0-9][0-9.,]*)\s*€", up)
        base = norm_num(mbase.group(1)) if mbase else None
        iva = norm_num(miva.group(1)) if miva else None
        total = norm_num(mtotal.group(1)) if mtotal else None
        pct = 0.21 if base and iva else None
        if base and iva and not total:
            total = round(base + iva, 2)
        return [
            {
                "fecha_factura": fecha,
                "numero_factura": number or path.stem,
                "empresa": "TELEFÓNICA DE ESPAÑA, S.A.U.",
                "CIF": norm_cif("A82018474"),
                "importe_base": base,
                "%IVA": pct,
                "IVA": iva,
                "importe_total": total,
                "Notas": "",
            }
        ]
