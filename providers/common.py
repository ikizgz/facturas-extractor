# -*- coding: utf-8 -*-
from __future__ import annotations

import re
import unicodedata
from datetime import datetime
from typing import Dict, List, Optional, Union

ADDRESS_TOKENS = [
    r"\bCALLE\b",
    r"\bCL\b",
    r"\bC/\b",
    r"\bAVDA?\b",
    r"\bCRTA\b",
    r"\bKM\b",
    r"\bTEL\b",
    r"\bFAX\b",
    r"\bCP\b",
    r"\bZARAGOZA\b",
    r"\bESPAÑA\b",
]
BLOCKLIST_TOKENS = {
    "PROPIETARIO",
    "FRA.CONTADO",
    "CONTADO",
    "CARRETERA",
    "ZARAGOZA",
    "GARRAPINILLOS",
    "REFERENCIA",
    "FRACONTADO",
    "FORMA DE PAGO",
    "ORIGINAL",
    "MANDAREMOS EL RECIBO A TU CUENTA",
    "IBERCAJA BANCO",
    "DE UN VISTAZO",
    "TOTAL A PAGAR",
    "PÁGINA",
    "ADQUIRIENTE",
    "TITULAR",
}
MONTHS_ES = {
    "ENERO": 1,
    "FEBRERO": 2,
    "MARZO": 3,
    "ABRIL": 4,
    "MAYO": 5,
    "JUNIO": 6,
    "JULIO": 7,
    "AGOSTO": 8,
    "SEPTIEMBRE": 9,
    "SETIEMBRE": 9,
    "OCTUBRE": 10,
    "NOVIEMBRE": 11,
    "DICIEMBRE": 12,
}
NUM_MONEY_RE = re.compile(r"([€]?\d[\d.,]*)\s*(?!%)")
NUM_PCT_RE = re.compile(r"(\d{1,2}(?:[.,]\d{1,2})?)\s*%")
VAT_ROW_RE = re.compile(
    r"(\d{1,2}(?:[.,]\d{1,2})?)\s*%[\s\S]*?(\d[\d.,]*)[\s\S]*?(\d[\d.,]*)"
)
VAT_ES_RE = re.compile(
    r"^(ES)?([A-HJNPQRSUVW]\d{7}[0-9A-J]|\d{8}[A-Z]|[XYZ]\d{7}[A-Z])$"
)
VAT_EU_RE = re.compile(r"^[A-Z]{2}[A-Z0-9\-.]{8,14}$")
Row = Dict[str, Optional[Union[str, float]]]


def strip_accents_punct(s: Optional[str]) -> str:
    s = s or ""
    s = "".join(
        c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn"
    )
    s = re.sub(r"[^A-Za-z0-9 ÁÉÍÓÚáéíóú&.,\-]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def norm_cif(s: Optional[str]) -> str:
    if not s:
        return ""
    return re.sub(r"[^A-Za-z0-9]", "", str(s)).upper()


def plausible_vat(cid: Optional[str]) -> bool:
    if not cid:
        return False
    cid_n = norm_cif(cid)
    return bool(VAT_ES_RE.match(cid_n) or VAT_EU_RE.match(cid_n))


def norm_num(s: Optional[str]) -> Optional[float]:
    if s is None:
        return None
    st = str(s).strip()
    if not st:
        return None
    for sym in ("€", "EUR", " ", " "):
        st = st.replace(sym, "")
    st = st.replace("%", "")
    if "," in st and "." in st:
        st = st.replace(".", "").replace(",", ".")
    elif "," in st:
        st = st.replace(",", ".")
    try:
        return float(st)
    except Exception:
        return None


def to_decimal_pct(s: Optional[Union[str, float]]) -> Optional[float]:
    if s is None:
        return None
    if isinstance(s, float):
        return round(s / 100.0, 6) if s > 1.0 else round(s, 6)
    st = str(s).strip().replace(" ", "").replace(" ", "")
    if st.endswith("%"):
        st = st[:-1]
    st = st.replace(",", ".")
    try:
        val = float(st)
    except ValueError:
        return None
    return round(val / 100.0, 6) if val > 1.0 else round(val, 6)


def parse_date_text(text: Optional[str]) -> Optional[str]:
    txt = text or ""
    m = re.search(
        r"Fecha\s+Factura\s*[:#]?\s*(\d{1,2}/\d{1,2}/\d{4})", txt, re.IGNORECASE
    )
    if m:
        dd, mm, yyyy = m.group(1).split("/")
        try:
            mi = int(mm)
            if 1 <= mi <= 12:
                return datetime.strptime(m.group(1), "%d/%m/%Y").date().isoformat()
        except Exception:
            pass
    candidates: List[datetime] = []
    for pat, fmt in [
        (r"(\d{1,2}/\d{1,2}/\d{4})", "%d/%m/%Y"),
        (r"(\d{1,2}-\d{1,2}-\d{4})", "%d-%m-%Y"),
    ]:
        for m2 in re.finditer(pat, txt):
            dd, mm, yyyy = re.split(r"/|-", m2.group(1))
            try:
                mi = int(mm)
                yi = int(yyyy)
                if 1 <= mi <= 12 and yi >= 2018:
                    candidates.append(datetime.strptime(m2.group(1), fmt))
            except Exception:
                continue
    if candidates:
        best = sorted(candidates, key=lambda d: (d.year, d.month, d.day), reverse=True)[
            0
        ]
        return best.date().isoformat()
    m3 = re.search(
        r"(\d{1,2})\s+de\s+([A-Za-zÁÉÍÓÚáéíóú]+)\s+de\s+(\d{4})", txt, re.IGNORECASE
    )
    if m3:
        day = int(m3.group(1))
        mon = MONTHS_ES.get(strip_accents_punct(m3.group(2)).upper())
        year = int(m3.group(3))
        if mon and year >= 2018:
            try:
                return datetime(year, mon, day).date().isoformat()
            except Exception:
                pass
    return None
