#!/usr/bin/env python3
# providers/leroymerlin.py

from __future__ import annotations

import re

from .base import ProviderParser
from .common import Row, norm_num, parse_date_text


class LeroyMerlinParser(ProviderParser):
    name = "LEROY MERLIN ESPANA S.L.U."

    def detect(self, text: str) -> bool:
        """Detecta si la factura corresponde a Leroy Merlin mediante palabras clave o NIF."""
        t = text.upper()
        return any(
            x in t for x in ["LEROY-MERLIN", "LEROY MERLIN", "B84818442", "B-84818442"]
        )

    def parse(self, text: str, path) -> list[Row]:
        # Normalizamos los espacios del texto bruto extraído
        raw_text = " ".join(text.split())

        # 1. Extracción del número de factura y fecha de emisión
        mnum = re.search(r"FACTURA\s*(\d{3}-\d{4}-\d{6,7})", raw_text, re.IGNORECASE)
        number = mnum.group(1) if mnum else path.stem
        fecha = parse_date_text(text)

        base_imp = None
        iva_val = None
        total_factura = None
        iva_pct = 0.21  # Por defecto

        # 2. Extracción del porcentaje de IVA aplicado (ej: 21.00)
        tasa_match = re.search(
            r"(\d{1,2}[.,]\d{2})\s+\d{1,2}[.,]\d{2}\s+-?\d+", raw_text
        )
        if tasa_match:
            pct_val = norm_num(tasa_match.group(1))
            if pct_val and pct_val > 0:
                iva_pct = pct_val / 100.0

        # 3. Extracción milimétrica de la línea final de totales (a partir de la etiqueta 'EUR')
        eur_matches = list(
            re.finditer(
                r"\bEUR\b\s*(-?\d+(?:[.,]\d{3})*[.,]\d{2})", raw_text, re.IGNORECASE
            )
        )

        if eur_matches:
            # Nos posicionamos en la última coincidencia de 'EUR' (cierres de totales)
            ultima_coincidencia = eur_matches[-1].start()
            segmento_final = raw_text[ultima_coincidencia:]

            # Capturamos todos los importes decimales (con soporte para negativos en abonos)
            nums_finales = re.findall(
                r"-?\b\d{1,3}(?:[.,]\d{3})*[.,]\d{2}\b", segmento_final
            )
            limpios = [norm_num(n) for n in nums_finales if norm_num(n) is not None]

            # Orden exacto tras el EUR: [Base Imponible, Cuota IVA, Total TIl]
            if len(limpios) >= 3:
                base_imp = limpios[0]
                iva_val = limpios[1]  # Valor real impreso en factura
                total_factura = limpios[2]
            elif len(limpios) == 2:
                base_imp = limpios[0]
                total_factura = limpios[1]

        # 4. Red de seguridad general por si el bloque 'EUR' variase
        if base_imp is None or total_factura is None:
            todos_nums = re.findall(r"-?\b\d{1,3}(?:[.,]\d{3})*[.,]\d{2}\b", raw_text)
            nums_validos = [norm_num(n) for n in todos_nums if norm_num(n) is not None]
            if len(nums_validos) >= 3:
                base_imp, iva_val, total_factura = nums_validos[-3:]

        # 5. Red de seguridad extrema opcional solo si el IVA no se pudo leer directamente
        if iva_val is None and base_imp is not None and total_factura is not None:
            iva_val = round(total_factura - base_imp, 2)

        # 6. Control de calidad y notas para la revisión
        nota = ""
        if len(text) < 300 or base_imp is None or total_factura is None:
            nota = "OCR: Revisar importes Leroy Merlin"

        # 7. Construcción de la fila de resultados estructurada
        rows: list[Row] = []
        rows.append(
            {
                "fecha_factura": fecha,
                "numero_factura": number,
                "empresa": "LEROY MERLIN ESPANA S.L.U.",
                "CIF": "B84818442",
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
