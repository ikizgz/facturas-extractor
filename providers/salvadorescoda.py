# -*- coding: utf-8 -*-
# providers/salvadorescoda.py

from __future__ import annotations

import re
from typing import List

from .base import ProviderParser
from .common import Row, norm_num, parse_date_text


class SalvadorEscodaParser(ProviderParser):
    name = "SALVADOR ESCODA S.A."

    def detect(self, text: str) -> bool:
        t = text.upper()
        return any(
            x in t for x in ["SALVADOR ESCODA", "ESCODA", "A08710006", "A-08710006"]
        )

    def parse(self, text: str, path) -> List[Row]:
        raw_text = " ".join(text.split())

        # 1. Número de factura y fecha
        mnum = re.search(r"FACTURA\D*(\d{6,7})", raw_text, re.IGNORECASE)
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

        # 3. Extracción milimétrica basada en la estructura que nos has indicado:
        # Buscamos el fragmento exacto que va desde "BASE IMPONIBLE" hasta "FORMA DE PAGO"
        bloque_match = re.search(
            r"BASE\s+IMPONIBLE.*?(?=FORMA\s+DE\s+PAGO|$)", raw_text, re.IGNORECASE
        )

        if bloque_match:
            segmento = bloque_match.group(0)

            # Extraemos todos los números con formato decimal en este segmento ordenados tal cual aparecen
            nums_encontrados = re.findall(
                r"\b\d{1,3}(?:[.,]\d{3})*[.,]\d{2}\b", segmento
            )
            nums_limpios = [
                norm_num(n) for n in nums_encontrados if norm_num(n) is not None
            ]

            # Según tu desglose exacto:
            # En la línea de valores aparecen típicamente:
            # 1. Importe inicial (ej: 101,38 o 55,90)
            # 2. Base imponible repetida o equivalente (ej: 101,38 o 55,90)
            # 3. Cuota de IVA (ej: 21,29 o 11,74)
            # 4. Total factura (detrás de EUR)

            # Descartamos el valor del porcentaje si se coló como número entero/decimal plano
            nums_limpios = [n for n in nums_limpios if n != valor_extraido]

            if len(nums_limpios) >= 3:
                # El segundo o primer número relevante suele ser la Base Imponible real
                base_imp = (
                    nums_limpios[1] if len(nums_limpios) >= 4 else nums_limpios[0]
                )
                iva_val = nums_limpios[
                    -2
                ]  # El penúltimo número antes del total es la cuota de IVA impresa
                total_factura = nums_limpios[-1]  # El último número es siempre el total
            elif len(nums_limpios) == 2:
                base_imp = nums_limpios[0]
                total_factura = nums_limpios[-1]

        # 4. Red de seguridad estricta para capturar el Total si el bloque fallase
        if total_factura is None:
            total_match = re.search(
                r"EUR\s*(\d{1,3}(?:[.,]\d{3})*[.,]\d{2})", raw_text, re.IGNORECASE
            )
            if total_match:
                total_factura = norm_num(total_match.group(1))

        # 5. Notas de control de calidad
        nota = ""
        if len(text) < 400 or not base_imp or not total_factura or not iva_val:
            nota = "OCR: Revisar importes impresos"

        rows: List[Row] = []
        rows.append(
            {
                "fecha_factura": fecha,
                "numero_factura": number,
                "empresa": "SALVADOR ESCODA S.A.",
                "CIF": "A08710006",
                "importe_base": float(base_imp) if base_imp else 0.0,
                "%IVA": float(iva_pct),
                "IVA": float(iva_val) if iva_val else 0.0,
                "importe_total": float(total_factura) if total_factura else 0.0,
                "Notas": nota,
            }
        )

        return rows
