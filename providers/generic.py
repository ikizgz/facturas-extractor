# -*- coding: utf-8 -*-
from __future__ import annotations

import re
from typing import List, Optional

from .base import ProviderParser
from .common import NUM_MONEY_RE, NUM_PCT_RE, VAT_ROW_RE, Row, norm_num, to_decimal_pct

LABELS = {
    "base": [
        r"BASE\s+IMPONIBLE",
        r"IMPORTE\s+BASE",
        r"\bBI\b",
        r"NETO",
        r"SUBTOTAL",
        r"TOTAL\s+BASE\s+IMPONIBLE",
    ],
    "iva": [
        r"CUOTA\s*IVA",
        r"IMPORTE\s*IVA",
        r"\bIVA\b",
        r"TOTAL\s*IVA\b",
        r"TOTAL\s+IMPUESTO",
    ],
    "total": [
        r"TOTAL\s*(?:FACTURA|A\s*PAGAR|EUR|€)?\b",
        r"\bTOTAL\b",
        r"TOTAL\s+DE\s+LA\s+FACTURA",
    ],
}


class GenericParser(ProviderParser):
    name = "GENERIC"

    def detect(self, text: str) -> bool:
        return True

    def _pick_money_candidate(
        self, cands, role, base_hint=None, iva_hint=None, pct_hint=None
    ):
        def score(x):
            v, has_euro, has_dec = x
            s = 0
            if has_euro:
                s += 3
            if has_dec:
                s += 1
            if role == "iva" and base_hint:
                if v <= (base_hint or 0) * 0.35:
                    s += 3
            if role == "total":
                target = None
                if base_hint is not None and pct_hint is not None:
                    target = base_hint * (1.0 + pct_hint)
                elif base_hint is not None and iva_hint is not None:
                    target = base_hint + iva_hint
                if target:
                    s += max(0, 3 - abs(v - target) / max(1.0, target) * 10)
            return s

        cands.sort(key=score, reverse=True)
        return cands[0][0] if cands else None

    def _find_value_by_label_smart(
        self, lines, patterns, role, base_hint=None, iva_hint=None, pct_hint=None
    ) -> Optional[float]:
        for pat in patterns:
            lab_re = re.compile(pat, re.IGNORECASE)
            for i, ln in enumerate(lines):
                if not lab_re.search(ln):
                    continue
                window = [ln] + lines[i + 1 : i + 5]
                cands = []
                for w in window:
                    euro = "€" in w
                    for m in NUM_MONEY_RE.finditer(w):
                        v = norm_num(m.group(1))
                        if v is None:
                            continue
                        has_dec = "," in m.group(1) or "." in m.group(1)
                        cands.append((v, euro, has_dec))
                val = self._pick_money_candidate(
                    cands,
                    role,
                    base_hint=base_hint,
                    iva_hint=iva_hint,
                    pct_hint=pct_hint,
                )
                if val is not None:
                    return val
        return None

    def parse(self, text: str, path) -> List[Row]:
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        base_tbl, iva_tbl = None, None
        bases, cuotas = [], []
        for m in VAT_ROW_RE.finditer(text):
            b = norm_num(m.group(2))
            c = norm_num(m.group(3))
            if b is not None and c is not None:
                bases.append(b)
                cuotas.append(c)
        if bases and cuotas:
            base_tbl, iva_tbl = round(sum(bases), 2), round(sum(cuotas), 2)
        pct = None
        m_pct = NUM_PCT_RE.search(text)
        if m_pct:
            pct = to_decimal_pct(m_pct.group(1))
        base = (
            base_tbl
            if base_tbl is not None
            else self._find_value_by_label_smart(lines, LABELS["base"], role="base")
        )
        iva = (
            iva_tbl
            if iva_tbl is not None
            else self._find_value_by_label_smart(
                lines, LABELS["iva"], role="iva", base_hint=base, pct_hint=pct
            )
        )
        tot = self._find_value_by_label_smart(
            lines,
            LABELS["total"],
            role="total",
            base_hint=base,
            iva_hint=iva,
            pct_hint=pct,
        )
        if iva is not None and float(iva).is_integer() and int(iva) in (4, 10, 21):
            if pct is None:
                pct = int(iva) / 100.0
            iva = None
        if pct is None and base and iva and base > 0:
            pct = round(iva / base, 6)
        if tot is None and base is not None and iva is not None:
            tot = round(base + iva, 2)
        if base and tot and tot < base:
            if pct:
                tot = round(base * (1.0 + pct), 2)
            elif iva:
                tot = round(base + iva, 2)
        return [
            {
                "fecha_factura": None,
                "numero_factura": path.stem,
                "empresa": None,
                "CIF": None,
                "importe_base": base,
                "%IVA": pct,
                "IVA": iva,
                "importe_total": tot,
                "Notas": "",
            }
        ]
