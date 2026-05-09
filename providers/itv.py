# providers/itv.py
from __future__ import annotations

import re
from typing import List, Optional

from .base import ProviderParser
from .common import Row, norm_cif, norm_num, parse_date_text


class ItvParser(ProviderParser):
    name = "ARAGONESA DE SERVICIOS ITV"

    def detect(self, text: str) -> bool:
        up = text.upper()
        return ("ARAGONESA DE SERVICIOS ITV" in up) or ("SERVICIOS ITV, S.A." in up)

    def _find(self, up: str, pat: str) -> Optional[float]:
        m = re.search(pat, up)
        return norm_num(m.group(1)) if m else None

    def parse(self, text: str, path) -> List[Row]:
        up = text.upper()
        # Nº factura: "FACTURA N* 000001743/50072024F" (OCR variantes: "N*2")
        mnum = re.search(r"FACTURA\s+N\*?\d*\s*([0-9]{6,}/[0-9A-Z]+)", up)
        number = mnum.group(1) if mnum else None

        # Fecha: dd/mm/yyyy
        fecha = parse_date_text(text)

        # Base / tasa / total (IVA derivado)
        base = self._find(up, r"BASE\s+IMPONIBLE\s*[:\s]*([0-9][0-9.,]*)")
        tasa = None
        for pat in [
            r"TASA\s+TR[ÁA]FICO\s*[:\s]*([0-9][0-9.,]*)",
            r"TASA\s+T[RÁA]FICO\s*[:\s]*([0-9][0-9.,]*)",
        ]:
            tasa = self._find(up, pat)
            if tasa is not None:
                break
        total = self._find(up, r"TOTAL\s+FACTURA\s*[:\s]*([0-9][0-9.,]*)")

        # IVA = total - base - tasa
        iva = None
        if base is not None and total is not None:
            iva = round(total - (base or 0.0) - (tasa or 0.0), 2)
            if iva < 0:
                iva = None

        # % IVA sobre la línea de base
        pct = None
        if base and iva and base > 0:
            pct = round(iva / base, 6)

        # Notas comunes (dos IVAs): "VARIOS IVAS + TOTAL <importe>"
        notas = (
            f"VARIOS IVAS + TOTAL {total:.2f}" if total is not None else "VARIOS IVAS"
        )

        rows: List[Row] = []
        # Línea 1: servicio (base + IVA)
        rows.append(
            {
                "fecha_factura": fecha,
                "numero_factura": number or path.stem,
                "empresa": "ARAGONESA DE SERVICIOS ITV, S.A.",
                "CIF": norm_cif("A18096511"),
                "importe_base": base,
                "%IVA": pct,
                "IVA": iva,
                "importe_total": (
                    round((base or 0.0) + (iva or 0.0), 2)
                    if (base is not None and iva is not None)
                    else None
                ),
                "Notas": notas,
            }
        )
        # Línea 2: tasa tráfico (IVA 0%)
        rows.append(
            {
                "fecha_factura": fecha,
                "numero_factura": number or path.stem,
                "empresa": "ARAGONESA DE SERVICIOS ITV, S.A.",
                "CIF": norm_cif("A18096511"),
                "importe_base": tasa,
                "%IVA": 0.0 if tasa is not None else None,
                "IVA": 0.0 if tasa is not None else None,
                "importe_total": tasa,
                "Notas": notas,
            }
        )
        return rows
