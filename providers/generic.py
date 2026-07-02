# -*- coding: utf-8 -*-
from __future__ import annotations

import re
from typing import List, Optional

from .base import ProviderParser
from .common import Row, norm_num, parse_date_text

LABELS = {
    "base": [r"BASE\s+IMPONIBLE", r"IMPORTE\s+BASE", r"\bBI\b", r"NETO", r"SUBTOTAL"],
    "iva": [r"CUOTA\s*IVA", r"IMPORTE\s*IVA", r"\bIVA\b", r"TOTAL\s*IVA"],
    "total": [r"TOTAL\s*(?:FACTURA|A\s*PAGAR|EUR|€)?\b", r"\bTOTAL\b"],
}


class GenericParser(ProviderParser):
    name = "GENERIC"

    def detect(self, text: str) -> bool:
        return True

    def parse(self, text: str, path) -> List[Row]:
        # 1. Dividir por páginas
        pages = re.split(r"--- PAGE \d+ ---", text)
        pages = [p.strip() for p in pages if p.strip()]

        # 2. Datos identificativos (siempre de la primera página)
        first_page = pages[0] if pages else text
        fecha = parse_date_text(first_page)

        # 3. Búsqueda inteligente de importes:
        # Recorremos las páginas de la última a la primera hasta encontrar los datos
        base, iva, tot = None, None, None

        for p in reversed(pages):
            if base is None:
                base = self._find_value_by_label(p, LABELS["base"], "base")
            if iva is None:
                iva = self._find_value_by_label(p, LABELS["iva"], "iva")
            if tot is None:
                tot = self._find_value_by_label(p, LABELS["total"], "total")

            # Si ya tenemos los 3, podemos parar de buscar
            if all([base, tot]):
                break

        return [
            {
                "fecha_factura": fecha,
                "numero_factura": path.stem,
                "empresa": "PROVEEDOR DESCONOCIDO",
                "CIF": "...",  # (Tu lógica de CIF anterior)
                "importe_base": base,
                "%IVA": None,
                "IVA": iva,
                "importe_total": tot,
                "Notas": "Genérico - Revisado pág. atrás",
            }
        ]

    def _find_value_by_label(
        self, text: str, patterns: List[str], role: str
    ) -> Optional[float]:
        # 1. Buscamos la línea que contenga la etiqueta
        for pat in patterns:
            lab_re = re.compile(pat, re.IGNORECASE)
            # Vamos a buscar de abajo hacia arriba (última página)
            lines = text.splitlines()
            for line in reversed(lines):
                if lab_re.search(line):
                    # Usamos una regex más estricta para el dinero:
                    # Busca un número que tenga al menos una coma o punto, seguido de opcional €
                    m = re.search(
                        r"(\d+[\.,]\d{2})\s*(?:€|EUR)?", line.replace(",", ".")
                    )
                    if m:
                        # Filtro de seguridad: si el número es sospechosamente pequeño
                        # y estamos buscando un total, lo ignoramos
                        val = norm_num(m.group(1))
                        # Ignoramos si es None o si es un total demasiado pequeño
                        if val is None or (role == "total" and val < 1.0):
                            continue
                        return val
        return None
