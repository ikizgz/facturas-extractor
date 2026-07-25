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
        clean_text = re.sub(r"[^a-zA-Z0-9/.,\s]", " ", text)
        raw_text = " ".join(clean_text.split())

        # 2. Número de factura (quirúrgico)
        mnum = re.search(r"(\d{9}/[A-Z0-9]{8}F)", raw_text)
        number = mnum.group(1) if mnum else path.stem
        fecha = parse_date_text(text)

        # 3. Extracción de valores
        def get_val(keyword):
            # El .*? permite saltar el "1" intruso o cualquier basura
            pattern = rf"{keyword}.*?([0-9]+[,.][0-9]{{2}})"
            m = re.search(pattern, raw_text, re.IGNORECASE)
            return norm_num(m.group(1)) if m else None

        base_imp = get_val("BASE IMPONIBLE")
        # Captura de la tasa ignorando cualquier carácter intermedio
        tasa = get_val(r"TASA.*?\s+([0-9]+[,.][0-9]{2})")
        iva_cuota = get_val(r"IVA[^\d]*21")
        # Captura del total ignorando cualquier carácter intermedio
        total_factura = get_val(r"TOTAL FACTURA.*?\s+([0-9]+[,.][0-9]{2})")

        rows: list[Row] = []

        # 4. Fila 1: Servicio ITV
        if base_imp:
            iva_val = iva_cuota or round(base_imp * 0.21, 2)
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

        # 5. Fila 2: Tasa (Solo si detecta valor > 0)
        if tasa and tasa > 0:
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
