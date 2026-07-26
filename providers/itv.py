#!/usr/bin/env python3
# providers/itv.py

from __future__ import annotations

import re

from .base import ProviderParser
from .common import Row, norm_num, parse_date_text


class ItvParser(ProviderParser):
    name = "ARAGONESA DE SERVICIOS ITV"

    def detect(self, text: str) -> bool:
        return "ARAGONESA DE SERVICIOS ITV" in text.upper()

    def parse(self, text: str, path) -> list[Row]:
        # 1. Limpieza selectiva
        clean_text = re.sub(
            r"[^a-zA-Z0-9/.,\s-]", " ", text
        )  # Añadido '-' por si hay negativos
        raw_text = " ".join(clean_text.split())

        # 2. Número de factura (quirúrgico)
        mnum = re.search(r"(\d{9}/[A-Z0-9]{8}F)", raw_text)
        number = mnum.group(1) if mnum else path.stem
        fecha = parse_date_text(text)

        # 3. Extracción de valores (añadido -? para admitir importes negativos)
        def get_val(keyword):
            pattern = rf"{keyword}.*?(-?[0-9]+[,.][0-9]{{2}})"
            m = re.search(pattern, raw_text, re.IGNORECASE)
            return norm_num(m.group(1)) if m else None

        base_imp = get_val("BASE IMPONIBLE")
        tasa = get_val(
            "TASA"
        )  # Simplificado para usar la misma lógica robusta de get_val
        iva_cuota = get_val(r"IVA[^\d]*21")
        total_factura = get_val("TOTAL FACTURA")

        rows: list[Row] = []

        # 4. Fila 1: Servicio ITV
        if base_imp is not None:
            iva_val = iva_cuota if iva_cuota is not None else round(base_imp * 0.21, 2)
            rows.append(
                {
                    "fecha_factura": fecha,
                    "numero_factura": number,
                    "empresa": "ARAGONESA DE SERVICIOS ITV, S.A.",
                    "CIF": "A18096511",
                    "importe_base": base_imp,
                    "%IVA": 0.21,
                    "IVA": iva_val,
                    "importe_total": round(base_imp + iva_val, 2),
                    "Notas": f"Servicio ITV | Importe total fra: {total_factura}",
                }
            )

        # 5. Fila 2: Tasa (Solo si detecta valor distinto de 0)
        if tasa and tasa != 0:
            rows.append(
                {
                    "fecha_factura": fecha,
                    "numero_factura": number,
                    "empresa": "ARAGONESA DE SERVICIOS ITV, S.A.",
                    "CIF": "A18096511",
                    "importe_base": tasa,
                    "%IVA": 0.0,
                    "IVA": 0.0,
                    "importe_total": tasa,
                    "Notas": f"Tasa Tráfico | Importe total fra: {total_factura}",
                }
            )

        return rows
