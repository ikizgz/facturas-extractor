#-*- coding: utf-8 -*-
# providers/nuevo_proveedor.

from __future__ import annotations
import re
from typing import List
from .base import ProviderParser
from .common import Row, norm_num, parse_date_text, norm_cif

class MiProveedorParser(ProviderParser):
    name = "MI_PROVEEDOR"

    def detect(self, text: str) -> bool:
        up = text.upper()
        return "MI PROVEEDOR S.A." in up or "MI LOGO" in up

    def parse(self, text: str, path) -> List[Row]:
        up = text.upper()
        # 1) Nº de factura por etiqueta: "Factura núm:"
        mnum = re.search(r"FACTURA\\s+N[ÚU]M\\s*[:#]?\\s*([A-Z0-9\\-/]+)", up)
        number = mnum.group(1) if mnum else None

        # 2) Fecha (usa helper que soporta dd/mm/yyyy y 'dd de mes de yyyy')
        fecha = parse_date_text(text)

        # 3) Importes por labels
        mbase = re.search(r"BASE\\s+IMPONIBLE\\s*([0-9][0-9.,]*)", up)
        miva  = re.search(r"TOTAL\\s+IVA\\s*([0-9][0-9.,]*)", up)
        mtotal= re.search(r"TOTAL\\s+FACTURA\\s*([0-9][0-9.,]*)", up)

        base = norm_num(mbase.group(1)) if mbase else None
        iva  = norm_num(miva.group(1))  if miva  else None
        total= norm_num(mtotal.group(1))if mtotal else None

        if base and iva and not total:
            total = round(base + iva, 2)

        pct = round(iva/base, 6) if base and iva and base > 0 else None

        return [{
            "fecha_factura": fecha,
            "numero_factura": number or path.stem,
            "empresa": "MI PROVEEDOR S.A.",
            "CIF": norm_cif("A00000000"),
            "importe_base": base,
            "%IVA": pct,
            "IVA": iva,
            "importe_total": total,
            "Notas": "",
        }]
