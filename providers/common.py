#!/usr/bin/env python3
# common.py

from __future__ import annotations

import logging
import re
import unicodedata
from datetime import datetime

logger = logging.getLogger(__name__)

# --- NUEVO: Lista de CIFs/NIFs propios para no confundirlos con los del proveedor ---
MY_CIFS: set[str] = {
    "J99198285",  # ACR S.C.
    # Añade aquí otros NIFs/CIFs de tus empresas si tienes más
}

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
    r"\bUTEBO\b",  # Nuevo
    r"\bGARRAPINILLOS\b",  # Nuevo
]

BLOCKLIST_TOKENS = {
    "PROPIETARIO",
    "FRA.CONTADO",
    "CONTADO",
    "CARRETERA",
    "ZARAGOZA",
    "UTEBO",  # Nuevo
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

# --- MEJORADO: Añadidos meses abreviados (3 letras) ---
MONTHS_ES = {
    "ENERO": 1,
    "ENE": 1,
    "FEBRERO": 2,
    "FEB": 2,
    "MARZO": 3,
    "MAR": 3,
    "ABRIL": 4,
    "ABR": 4,
    "MAYO": 5,
    "MAY": 5,
    "JUNIO": 6,
    "JUN": 6,
    "JULIO": 7,
    "JUL": 7,
    "AGOSTO": 8,
    "AGO": 8,
    "SEPTIEMBRE": 9,
    "SETIEMBRE": 9,
    "SEP": 9,
    "SET": 9,
    "OCTUBRE": 10,
    "OCT": 10,
    "NOVIEMBRE": 11,
    "NOV": 11,
    "DICIEMBRE": 12,
    "DIC": 12,
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
Row = dict[str, str | float | None]


def strip_accents_punct(s: str | None) -> str:
    s = s or ""
    s = "".join(
        c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn"
    )
    s = re.sub(r"[^A-Za-z0-9 ÁÉÍÓÚáéíóú&.,\-]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def norm_cif(s: str | None) -> str:
    if not s:
        return ""
    return re.sub(r"[^A-Za-z0-9]", "", str(s)).upper()


def plausible_vat(cid: str | None) -> bool:
    if not cid:
        return False
    cid_n = norm_cif(cid)

    # --- MEJORADO: Ignorar los CIFs propios ---
    if cid_n in MY_CIFS:
        return False

    return bool(VAT_ES_RE.match(cid_n) or VAT_EU_RE.match(cid_n))


def norm_num(s: str | None) -> float | None:
    if s is None:
        return None
    st = str(s).strip()
    if not st:
        return None
    for sym in ("€", "EUR", " ", " "):
        st = st.replace(sym, "")
    st = st.replace("%", "")
    if "," in st and "." in st:
        st = st.replace(".", "").replace(",", ".")
    elif "," in st:
        st = st.replace(",", ".")
    try:
        return float(st)
    except (ValueError, TypeError):
        return None


def to_decimal_pct(s: str | float | None) -> float | None:
    if s is None:
        return None
    if isinstance(s, float):
        return round(s / 100.0, 6) if s > 1.0 else round(s, 6)
    st = str(s).strip().replace(" ", "").replace(" ", "")
    st = st.removesuffix("%")
    st = st.replace(",", ".")
    try:
        val = float(st)
    except (ValueError, TypeError):
        return None
    return round(val / 100.0, 6) if val > 1.0 else round(val, 6)


def parse_date_text(text: str | None) -> str | None:
    txt = text or ""
    m = re.search(
        r"Fecha\s+Factura\s*[:#]?\s*(\d{1,2}/\d{1,2}/\d{4})", txt, re.IGNORECASE
    )
    if m:
        dd, mm, yyyy = m.group(1).split("/")
        try:
            mi = int(mm)
            if 1 <= mi <= 12:
                return datetime.strptime(m.group(1), "%d/%m/%Y").date().isoformat()  # noqa: DTZ007
        except (ValueError, TypeError) as e:
            logger.warning("Error al procesar la fecha de la factura: %s", e)

    # SI NO ENCONTRÓ LA FECHA ANTERIOR
    candidates: list[datetime] = []

    # 1. Buscar formatos estándar dd/mm/yyyy o dd-mm-yyyy
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
                    candidates.append(datetime.strptime(m2.group(1), fmt))  # noqa: DTZ007
            except (ValueError, TypeError) as e:
                logger.debug("Error analizando formato estándar de fecha: %s", e)
                continue

    # --- NUEVO: Buscar formatos con mes abreviado (Ej: 24-ABR-26) ---
    for m_abrev in re.finditer(r"(\d{1,2})[-/]([A-Za-z]{3,})[-/](\d{2,4})", txt):
        dd = int(m_abrev.group(1))
        mes_str = strip_accents_punct(m_abrev.group(2)).upper()
        yy = int(m_abrev.group(3))

        mi = MONTHS_ES.get(mes_str)
        if mi:
            yi = yy if yy >= 2000 else 2000 + yy
            if yi >= 2018:
                try:
                    candidates.append(datetime(yi, mi, dd))  # noqa: DTZ001
                except (ValueError, TypeError) as e:
                    logger.debug("Error procesando fecha con mes abreviado: %s", e)

    if candidates:
        best = max(candidates, key=lambda d: (d.year, d.month, d.day))
        return best.date().isoformat()

    # 3. Buscar formato con o sin "de" (Ej: "20 julio 2026" o "19 de julio de 2026")
    m3 = re.search(
        r"(\d{1,2})\s+(?:de\s+)?([A-Za-zÁÉÍÓÚáéíóú]+)\s+(?:de\s+)?(\d{4})",
        txt,
        re.IGNORECASE,
    )
    if m3:
        day = int(m3.group(1))
        mon = MONTHS_ES.get(strip_accents_punct(m3.group(2)).upper())
        year = int(m3.group(3))
        if mon and year >= 2018:
            try:
                return datetime(year, mon, day).date().isoformat()  # noqa: DTZ001
            except (ValueError, TypeError) as e:
                logger.debug("Error procesando fecha literal extendida: %s", e)

    return None
