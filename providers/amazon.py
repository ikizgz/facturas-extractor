#!/usr/init/env python3
# providers/amazon.py

from __future__ import annotations

import re

from .base import ProviderParser
from .common import Row, norm_num, parse_date_text

KNOWN_AMAZON_ENTITIES = {
    "N0186600C": "Amazon Business EU S.à.r.l",
    "W0264006H": "Amazon Business EU S.à.r.l, Sucursal en España",
    "DE814584193": "Amazon EU S.à r.l., Niederlassung Deutschland",
    "FR12487773327": "Amazon EU S.à r.l., Succursale Française",
    "IT08973230967": "Amazon EU S.à r.l., Succursale Italiana",
    "W0184081H": "Amazon EU S.à r.l., Sucursal en España",
    "PL5262907815": "Amazon EU S.à r.l., Sucursal en Polonia",
}


class AmazonParser(ProviderParser):
    name = "AMAZON"

    def detect(self, text: str) -> bool:
        t = text.upper()
        return "AMAZON" in t or "AMAZON BUSINESS" in t

    def parse(self, text: str, path) -> list[Row]:
        raw_text = " ".join(text.split())

        # 1. Extracción robusta de la fecha de factura/entrega (capturando formatos como "20 julio 2026" o "19 de julio de 2026")
        fecha = None
        m_fecha = re.search(
            r"Fecha de la factura.*?Fecha de la entrega[^\d]*(\d{1,2}\s+(?:de\s+)?[a-zA-Z]+\s+(?:de\s+)?\d{4})",
            raw_text,
            re.IGNORECASE,
        )
        if not m_fecha:
            m_fecha = re.search(
                r"Fecha de la factura[^\d]*(\d{1,2}\s+(?:de\s+)?[a-zA-Z]+\s+(?:de\s+)?\d{4})",
                raw_text,
                re.IGNORECASE,
            )

        if m_fecha:
            fecha = parse_date_text(m_fecha.group(1))

        # Fallback si no se captura mediante la etiqueta
        if not fecha:
            todas_fechas = re.findall(
                r"(\d{1,2}\s+(?:de\s+)?[a-zA-Z]+\s+(?:de\s+)?\d{4})", text
            )
            if len(todas_fechas) >= 2:
                fecha = parse_date_text(todas_fechas[1])
            elif len(todas_fechas) == 1:
                fecha = parse_date_text(todas_fechas[0])
            else:
                fecha = parse_date_text(text)

        # 2. Número de factura
        mnum = re.search(r"Número de la factura\s*([A-Z0-9]+)", raw_text, re.IGNORECASE)
        number = mnum.group(1) if mnum else path.stem

        # 3. Vendedor y CIF exactos desde "Vendido por"
        cif = ""
        empresa = ""

        vendido_match = re.search(
            r"Vendido por\s+([A-Za-z0-9àèìòùÀÈÌÒÙáéíóúÁÉÍÓÚñÑäëïöüÄËÏÖÜ\.\s,\-_]+?)(?=\s+IVA\s+[A-Z0-9]+|\s+Fecha de la factura|$)",
            raw_text,
            re.IGNORECASE,
        )
        if vendido_match:
            empresa = vendido_match.group(1).strip()

        iva_vend_match = re.search(
            r"Vendido por\s+.*?IVA\s*([A-Z0-9]+)", raw_text, re.IGNORECASE
        )
        if iva_vend_match:
            cif = iva_vend_match.group(1).strip()

        if cif in KNOWN_AMAZON_ENTITIES:
            empresa = KNOWN_AMAZON_ENTITIES[cif]
        elif not empresa:
            empresa = "Amazon Business EU S.à.r.l, Sucursal en España"
            cif = "W0264006H"

        base_imp = None
        iva_val = None
        total_factura = None
        iva_pct = 0.21

        # 4. Extracción precisa de la tabla de impuestos
        tabla_iva_match = re.search(
            r"IVA\s*(\d{1,2}(?:[.,]\d{1,2})?)\s*%\s*(-?\d+[.,]\d{2})\s*€?\s*(-?\d+[.,]\d{2})\s*€?",
            raw_text,
            re.IGNORECASE,
        )
        if tabla_iva_match:
            pct_val = norm_num(tabla_iva_match.group(1))
            if pct_val is not None:
                iva_pct = pct_val / 100.0
            base_imp = norm_num(tabla_iva_match.group(2))
            iva_val = norm_num(tabla_iva_match.group(3))

        tot_match = re.search(
            r"Total de la factura\s*(-?\d+[.,]\d{2})\s*€", raw_text, re.IGNORECASE
        )
        if not tot_match:
            tot_match = re.search(
                r"Total\s*(-?\d+[.,]\d{2})\s*€", raw_text, re.IGNORECASE
            )

        if tot_match:
            total_factura = norm_num(tot_match.group(1))
        else:
            todos_totales = re.findall(
                r"Total[^\d]*(-?\d+[.,]\d{2})\s*€", raw_text, re.IGNORECASE
            )
            if todos_totales:
                total_factura = norm_num(todos_totales[-1])

        if iva_pct == 0.0:
            if base_imp is None and total_factura is not None:
                base_imp = total_factura
            iva_val = 0.0

        if total_factura is None and base_imp is not None and iva_val is not None:
            total_factura = round(base_imp + iva_val, 2)

        if iva_val is None and base_imp is not None and total_factura is not None:
            iva_val = round(total_factura - base_imp, 2)

        # 5. Notas y control de calidad
        nota = ""
        if iva_pct == 0.0:
            nota = "Exenta / Inversión sujeto pasivo"

        if len(text) < 300 or base_imp is None or total_factura is None:
            nota = (
                "Exenta / Inversión sujeto pasivo - " if iva_pct == 0.0 else ""
            ) + "OCR: Revisar importes Amazon"

        # 6. Construcción de filas
        rows: list[Row] = []
        rows.append(
            {
                "fecha_factura": fecha,
                "numero_factura": number,
                "empresa": empresa,
                "CIF": cif,
                "importe_base": float(base_imp) if base_imp is not None else 0.0,
                "%IVA": float(iva_pct),
                "IVA": float(iva_val) if iva_val is not None else 0.0,
                "importe_total": float(total_factura)
                if total_factura is not None
                else 0.0,
                "Notas": nota.strip(),
            }
        )

        return rows
