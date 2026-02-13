# -*- coding: utf-8 -*-
from __future__ import annotations
import re
from typing import List
from .base import ProviderParser
from .common import Row, norm_num, parse_date_text, norm_cif

class MercadaizParser(ProviderParser):
    name = "MERCADAIZ"
    def detect(self, text: str) -> bool:
        up = text.upper()
        return ("VIUDA DE LONDAIZ" in up) or ("GASOLEOS MERCADAIZ" in up)
    def parse(self, text: str, path) -> List[Row]:
        up = text.upper()
        mnum = re.search(r"FA\s*[-/]\s*(\d{3,})", up)
        number = ("FA-" + mnum.group(1)) if mnum else None
        fecha = parse_date_text(text)
        mbase = re.search(r"BASE\s+IMPONIBLE\s*([0-9][0-9.,]*)", up)
        miva  = re.search(r"TOTAL\s+I\.?V\.?A\.?\s*([0-9][0-9.,]*)", up)
        mtotal= re.search(r"TOTAL\s+FACTURA\s*([0-9][0-9.,]*)", up)
        base = norm_num(mbase.group(1)) if mbase else None
        iva  = norm_num(miva.group(1))  if miva  else None
        total= norm_num(mtotal.group(1))if mtotal else None
        pct = (round(iva/base,6) if base and iva and base>0 else None)
        if base and iva and not total:
            total = round(base+iva,2)
        return [{
            "fecha_factura": fecha,
            "numero_factura": number or path.stem,
            "empresa": "VIUDA DE LONDAIZ Y SOBRINOS DE L. MERCADER, S.A.",
            "CIF": norm_cif("A20004008"),
            "importe_base": base,
            "%IVA": pct,
            "IVA": iva,
            "importe_total": total,
            "Notas": "",
        }]
