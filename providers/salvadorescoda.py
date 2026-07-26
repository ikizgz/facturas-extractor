#!/usr/bin/env python3
# providers/salvadorescoda.py

from __future__ import annotations

import re

from .base import ProviderParser
from .common import Row, norm_num, parse_date_text


class SalvadorEscodaParser(ProviderParser):
    name = "SALVADOR ESCODA S.A."

    def detect(self, text: str) -> bool:
        t = text.upper()
        return any(
            x in t for x in ["SALVADOR ESCODA", "ESCODA", "A08710006", "A-08710006"]
        )

    def parse(self, text: str, path) -> list[Row]:
        raw_text = " ".join(text.split())

        # 1. Número de factura (ampliado hasta 10 dígitos por si acaso) y fecha
        mnum = re.search(r"FACTURA\D*(\d{6,10})", raw_text, re.IGNORECASE)
        number = mnum.group(1) if mnum else path.stem
        fecha = parse_date_text(text)

        # 2. Captura exacta del % IVA justo antes de la etiqueta % IVA
        iva_pct_match = re.search(
            r"([0-9]+(?:[,.][0-9]{1,2})?)\s*%\s*(?:IVA|SR\.EQ|\b)",
            raw_text,
            re.IGNORECASE,
        )
        if not iva_pct_match:
            iva_pct_match = re.search(r"\b(4|10|21)(?:[,.00])?\s*%", raw_text)

        valor_extraido = norm_num(iva_pct_match.group(1)) if iva_pct_match else 21.0
        if valor_extraido is None or valor_extraido <= 0:
            valor_extraido = 21.0

        iva_pct = valor_extraido / 100.0

        base_imp = None
        iva_val = None
        total_factura = None

        # 3. Extracción del bloque de importes (permitiendo signo negativo opcional -?)
        bloque_match = re.search(
            r"BASE\s+IMPONIBLE.*?(?=FORMA\s+DE\s+PAGO|$)", raw_text, re.IGNORECASE
        )

        if bloque_match:
            segmento = bloque_match.group(0)

            # Añadimos -? opcional para capturar importes negativos en abonos/rectificativas
            nums_encontrados = re.findall(
                r"-?\b\d{1,3}(?:[.,]\d{3})*[.,]\d{2}\b", segmento
            )
            nums_limpios = [
                norm_num(n) for n in nums_encontrados if norm_num(n) is not None
            ]

            # Descartamos el valor del porcentaje si se coló como número entero/decimal plano
            nums_limpios = [n for n in nums_limpios if n != valor_extraido]

            if len(nums_limpios) >= 3:
                base_imp = (
                    nums_limpios[1] if len(nums_limpios) >= 4 else nums_limpios[0]
                )
                iva_val = nums_limpios[
                    -2
                ]  # Cuota de IVA (puede ser negativa en abonos)
                total_factura = nums_limpios[-1]  # Total factura
            elif len(nums_limpios) == 2:
                base_imp = nums_limpios[0]
                total_factura = nums_limpios[-1]

        # 4. Red de seguridad estricta para capturar el Total si el bloque fallase (añadido -?)
        if total_factura is None:
            total_match = re.search(
                r"EUR\s*(-?\d{1,3}(?:[.,]\d{3})*[.,]\d{2})", raw_text, re.IGNORECASE
            )
            if total_match:
                total_factura = norm_num(total_match.group(1))

        # 5. Notas de control de calidad
        nota = ""
        if (
            len(text) < 400
            or base_imp is None
            or total_factura is None
            or iva_val is None
        ):
            nota = "OCR: Revisar importes impresos"

        rows: list[Row] = []
        rows.append(
            {
                "fecha_factura": fecha,
                "numero_factura": number,
                "empresa": "SALVADOR ESCODA S.A.",
                "CIF": "A08710006",
                "importe_base": float(base_imp) if base_imp is not None else 0.0,
                "%IVA": float(iva_pct),
                "IVA": float(iva_val) if iva_val is not None else 0.0,
                "importe_total": float(total_factura)
                if total_factura is not None
                else 0.0,
                "Notas": nota,
            }
        )

        return rows
