# providers/alcampo.py
from __future__ import annotations

import re
from typing import List

from .base import ProviderParser
from .common import Row, norm_cif, norm_num, parse_date_text


class AlcampoParser(ProviderParser):
    name = "ALCAMPO"

    def detect(self, text: str) -> bool:
        up = text.upper()
        return (
            ("ALCAMPO S.A" in up)
            or ("FAT ALCAMPO" in up)
            or ("HIPERMERCADO UTEBO" in up)
        )

    def parse(self, text: str, path) -> List[Row]:
        up = text.upper()
        # Nº de factura: "Factura N*: 250500100877" o "Factura N%: 250600102127"
        mnum = re.search(r"FACTURA\s+N[\*%]?:\s*(\d{6,})", up)
        number = mnum.group(1) if mnum else None

        # Fecha: "Utebo, a 21 de Junio de 2025"
        fecha = parse_date_text(text)

        # Totales (pie)
        mbase = re.search(r"TOTAL\s+BASE\s+IMPONIBLE\s*([0-9][0-9.,]*)", up)
        miva = re.search(r"TOTAL\s+IMPUESTO\s*([0-9][0-9.,]*)", up)
        mtotal = re.search(r"TOTAL\s+FACTURA\s*([0-9][0-9.,]*)", up)
        base = norm_num(mbase.group(1)) if mbase else None
        iva = norm_num(miva.group(1)) if miva else None
        total = norm_num(mtotal.group(1)) if mtotal else None

        # Fallback: bloque principal
        if base is None:
            mbase2 = re.search(r"BASE\s+IMP\.?\s*([0-9][0-9.,]*)\s*€", up)
            if mbase2:
                base = norm_num(mbase2.group(1))
        if iva is None:
            miva2 = re.search(r"IMPUESTO\s*([0-9][0-9.,]*)\s*€", up)
            if miva2:
                iva = norm_num(miva2.group(1))
        if total is None:
            mtot2 = re.search(r"IMP\.\s*L[IÍ]QUIDO\.?\s*([0-9][0-9.,]*)\s*€", up)
            if mtot2:
                total = norm_num(mtot2.group(1))

        pct = None
        if base and iva and base > 0:
            pct = round(iva / base, 6)
        if total is None and base is not None and iva is not None:
            total = round(base + iva, 2)

        return [
            {
                "fecha_factura": fecha,
                "numero_factura": number or path.stem,
                "empresa": "ALCAMPO S.A.",
                "CIF": norm_cif("A28581882"),
                "importe_base": base,
                "%IVA": pct,
                "IVA": iva,
                "importe_total": total,
                "Notas": "",
            }
        ]
