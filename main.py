# -*- coding: utf-8 -*-
"""
main.py (limpio y optimizado, 2026-01)

- Secuencial por PDFs + proceso hijo con timeout (multiplataforma, sin resource/os.nice)
- OCR zonal con OSD/PSM más rápido por defecto (dpi=135, sleep_ms=0)
- Detección robusta de importes (labels ampliadas: Leroy Merlin, Alcampo, Amazon, Repsol, Norauto, etc.)
- Heurística proveedor/CIF reforzada + overrides por proveedor conocido
- Soporte de marcas con múltiples VAT (Amazon, etc.): prioriza VAT detectado y añade notas en ausencia o discrepancia
- Limpieza de razón social para evitar arrastrar dirección/registro
- Excel con formatos y defensas ante None
"""

from __future__ import annotations

import argparse
import gc
import json
import logging
import re
import time
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Union

import pandas as pd
from PyPDF2 import PdfReader

# -------------------- CONFIG RÁPIDA --------------------
DEFAULT_DPI = 135
DEFAULT_SLEEP_MS = 0
DEFAULT_THROTTLE_EVERY = 6
DEFAULT_THROTTLE_MS = 800
DEFAULT_CHILD_TIMEOUT_S = 60

# Cliente (para excluirlo de proveedor)
CLIENT_IDS = {"J99198285", "ESJ99198285"}
CLIENT_NAME_KEYWORDS = {
    "ARAGONESA DE CONDUCCIONES Y REFRIGERACION",
    "ARAGONESA DE CONDUCCIONES",
    "ACR SC",
    "ACR S.C.",
}

# Pistas de proveedor
KNOWN_VENDOR_HINTS = [
    "LEROY MERLIN",
    "SALVADOR ESCODA",
    "ALCAMPO",
    "REPSOL",
    "NORAUTO",
    "NOROTO",
    "REMLE",
    "AMAZON",
    "CREATORS OF TOMORROW",
    "WIESEMANN",
    "INDUSTRIAS REUNIDAS SANITARIAS",
    "ARAGONESA DE SERVICIOS ITV",
    "GASOLEOS MERCADAIZ",
    "XIAOJIE WANG",
    "SORPRESA HOGAR",
]
COMPANY_HINTS = {"REGISTRO MERCANTIL", "R.M.", "INSCRITA EN"}

# Overrides de proveedor (nombre/CIF canónicos cuando el texto contiene estas marcas)
# Nota: en marcas multi-VAT (p.ej. AMAZON) cif_fix = None para no imponer VAT.
VENDOR_OVERRIDES: List[Tuple[re.Pattern, str, Optional[str]]] = [
    (re.compile(r"LEROY\s+MERLIN", re.I), "LEROY MERLIN ESPAÑA, S.L.U.", "B84818442"),
    (re.compile(r"\bALCAMPO\b", re.I), "ALCAMPO S.A.", "A28581882"),
    (
        re.compile(r"REPSOL\s+SOLUCIONES\s+ENERGETIC", re.I),
        "REPSOL SOLUCIONES ENERGETICAS, S.A.",
        "A80298839",
    ),
    (re.compile(r"NORAUTO|NOROTO", re.I), "NOROTO S.A.U.", "A78119773"),
    (re.compile(r"SALVADOR\s+ESCODA", re.I), "SALVADOR ESCODA S.A.", "A08710006"),
    (re.compile(r"REMLE", re.I), "REMLE S.A.", "A08388811"),
    (re.compile(r"AMAZON", re.I), "AMAZON (ver detalle en factura)", None),
    (
        re.compile(r"SHENZHEN\s+GAOCHENG", re.I),
        "SHENZHEN GAOCHENG ELECTRONIC COMMERCE CO., LTD.",
        "ESN0050580J",
    ),
    (
        re.compile(r"CREATORS\s+OF\s+TOMORROW|WIESEMANN", re.I),
        "CREATORS OF TOMORROW GMBH",
        "N27643081",
    ),
    (
        re.compile(r"ARAGONESA\s+DE\s+SERVICIOS\s+ITV", re.I),
        "ARAGONESA DE SERVICIOS ITV, S.A.",
        None,
    ),
    (
        re.compile(r"INDUSTRIAS\s+REUNIDAS\s+SANITARIAS|INDUSAN", re.I),
        "INDUSTRIAS REUNIDAS SANITARIAS S.L.",
        "B50040005",
    ),
    (
        re.compile(r"GASOLEOS\s+MERCADAIZ|VIUDA\s+DE\s+LONDAIZ", re.I),
        "VIUDA DE LONDAIZ Y SOBRINOS DE L. MERCADER, S.A.",
        "A20004008",
    ),
]

# VATs conocidos (no obligatorios). Claves en mayúsculas.
KNOWN_VATS_BY_BRAND = {
    "AMAZON": {
        "ESN0186600C",
        "ESW0264006H",
        "DE814584193",
        "FR12487773327",
        "IT08973230967",
        "ESW0184081H",
        "PL5262907815",
    },
    # Añade aquí otras marcas multi-VAT si lo necesitas
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
LEGAL_SUFFIXES = [
    r"\bS\.?A\.?U?\.?\b",
    r"\bS\.?L\.?U?\.?\b",
    r"\bS\.?C\.?\b",
    r"\bCOOP\b",
]

# Etiquetas ampliadas (regex) incl. variantes comunes de proveedores
LABELS: Dict[str, List[str]] = {
    # Base imponible
    "base": [
        r"BASE\s+IMPONIBLE",
        r"IMPORTE\s+BASE",
        r"\bBI\b",
        r"NETO",
        r"SUBTOTAL",
        r"IMPORTE\s+DEL\s+PRODUCTO\s*\(\s*BASE\s+IMPONIBLE\s*\)",  # Repsol
        r"TOTAL\s*SI\b",
        r"TOTAL\s*SI\s*\(EUR\)",  # Leroy Merlin
        r"NET\s+AMOUNT",  # Wiesemann/Creators
        r"PRECIO\s+TOTAL\s*\(\s*IVA\s*EXCLUIDO\s*\)",  # Amazon
        r"TOTAL\s+BASE\s+IMPONIBLE",
    ],
    # IVA / cuota
    "iva": [
        r"CUOTA\s*IVA",
        r"IMPORTE\s*IVA",
        r"\bIVA\b",
        r"TOTAL\s*IVA\b",
        r"TOTAL\s+IMPUESTO",  # Alcampo
        r"VAT\b",
        r"IVA/IGIC/IPSI",
        r"IVANGICAPSI",
    ],
    # Porcentaje IVA
    "pct": [
        r"TASA\s*IVA\b",
        r"%\s*IVA\b",
        r"IVA\s*\d{1,2}(?:[.,]\d{1,2})?\s*%",
        r"VAT\s*\d{1,2}(?:[.,]\d{1,2})?\s*%",
    ],
    # Total factura
    "total": [
        r"TOTAL\s*(?:FACTURA|A\s*PAGAR|EUR|€)?\b",
        r"\bTOTAL\b",
        r"TOTAL\s*TII\b",
        r"TOTAL\s*TIl\b",  # Leroy Merlin (I/l)
        r"TOTAL\s+DE\s+LA\s+FACTURA",
        r"TOTAL\s+AMOUNT",
    ],
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
]


# -------------------- Utils --------------------
def strip_accents_punct(s: str) -> str:
    s = s or ""
    s = "".join(
        c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn"
    )
    s = re.sub(r"[^A-Za-z0-9 ÁÉÍÓÚáéíóú&.,\-]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def norm_cif(s: Optional[str]) -> Optional[str]:
    if not s:
        return s
    return re.sub(r"[^A-Za-z0-9]", "", str(s)).upper()


def norm_num(s: Optional[str]) -> Optional[float]:
    if s is None:
        return None
    st = str(s).strip()
    if not st:
        return None
    for sym in ("€", "EUR", " ", "\u00a0"):
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
    st = str(s).strip().replace(" ", "").replace("\u00a0", "")
    if st.endswith("%"):
        st = st[:-1]
    st = st.replace(",", ".")
    try:
        val = float(st)
    except ValueError:
        return None
    return round(val / 100.0, 6) if val > 1.0 else round(val, 6)


# -------------------- Texto PDF --------------------
def read_pdf_text_native(path: Path) -> str:
    try:
        reader = PdfReader(str(path))
        return "\n".join([(p.extract_text() or "") for p in reader.pages]) or ""
    except Exception as e:
        logging.debug(f"No se pudo leer texto nativo: {path.name} ({e})")
        return ""


def read_pdf_text_ocr(
    path: Path, dpi: int, sleep_ms: int, poppler_path: str, tesseract_exe: str
) -> str:
    try:
        import pytesseract
        from pdf2image import convert_from_path, pdfinfo_from_path

        if tesseract_exe:
            pytesseract.pytesseract.tesseract_cmd = tesseract_exe
        info = pdfinfo_from_path(str(path), userpw="", poppler_path=poppler_path)
        max_pages = int(info.get("Pages", 1))
        collected: List[str] = []
        for page in range(1, max_pages + 1):
            images = convert_from_path(
                str(path),
                dpi=dpi,
                first_page=page,
                last_page=page,
                fmt="jpeg",
                thread_count=1,
                poppler_path=poppler_path,
            )
            img = images[0]
            W, H = img.size
            # OSD
            try:
                osd = pytesseract.image_to_osd(img)
                if "Rotate: 90" in osd:
                    img = img.rotate(-90, expand=True)
                    W, H = img.size
                elif "Rotate: 270" in osd:
                    img = img.rotate(90, expand=True)
                    W, H = img.size
            except Exception:
                pass
            # Zonas (texto + números)
            cfg_text = "--psm 6"
            cfg_nums = "--psm 7 -c tessedit_char_whitelist=0123456789.,%€"
            header = img.crop((0, 0, W, int(0.20 * H)))
            footer = img.crop((0, int(0.80 * H), W, H))
            left = img.crop((0, 0, int(0.18 * W), H)).rotate(90, expand=True)
            totals = img.crop((int(0.45 * W), int(0.60 * H), W, H))
            collected.append(
                pytesseract.image_to_string(header, lang="spa+eng", config=cfg_text)
            )
            collected.append(
                pytesseract.image_to_string(footer, lang="spa+eng", config=cfg_text)
            )
            collected.append(
                pytesseract.image_to_string(left, lang="spa+eng", config=cfg_text)
            )
            collected.append(
                pytesseract.image_to_string(totals, lang="spa+eng", config=cfg_nums)
            )
            for im in (header, footer, left, totals):
                try:
                    im.close()
                except Exception:
                    pass
            try:
                img.close()
            except Exception:
                pass
            del images
            gc.collect()
            if sleep_ms > 0:
                time.sleep(sleep_ms / 1000.0)
        return "\n".join(collected)
    except Exception as e:
        logging.error(f"OCR falló en {path.name}: {e}")
        return ""


# -------------------- CIF/Empresa --------------------
ES_VAT_PATTERNS = [
    r"\b(?:ES\s*[-.]?)?([A-HJNPQRSUVW]\s*-?\s*\d{7}\s*-?\s*[0-9A-J])\b",
    r"\b(?:ES\s*[-.]?)?(\d{8}\s*-?\s*[A-Z])\b",
    r"\b(?:ES\s*[-.]?)?([XYZ]\s*-?\s*\d{7}\s*-?\s*[A-Z])\b",
]
EU_VAT_GENERIC = r"\b([A-Z]{2}[A-Z0-9\-.]{8,14})\b"


def _clean_company_line(s: str) -> str:
    s = re.sub(r"\bN[IF]{1,2}\b\s*[:#]?\s*[A-Z0-9\-\.]+", "", s, flags=re.IGNORECASE)
    m = re.search(r"[,\.]", s)
    if m:
        s = s[: m.start()]
    for tok in ADDRESS_TOKENS:
        s = re.sub(tok + r".*$", "", s, flags=re.IGNORECASE)
    return s.strip()


def find_vat_candidates(text: str) -> List[str]:
    cands = []
    for pat in ES_VAT_PATTERNS:
        for m in re.finditer(pat, text, flags=re.IGNORECASE):
            cid = norm_cif(m.group(1))
            if not cid:
                continue
            if cid in CLIENT_IDS or cid.replace("ES", "") in CLIENT_IDS:
                continue
            cands.append(cid)
    for m in re.finditer(EU_VAT_GENERIC, text):
        cid = norm_cif(m.group(1))
        if cid and cid not in CLIENT_IDS:
            cands.append(cid)
    # únicos manteniendo orden
    out: List[str] = []
    for c in cands:
        if c not in out:
            out.append(c)
    return out


def detect_brand(text_up: str) -> Optional[str]:
    for brand in KNOWN_VATS_BY_BRAND.keys():
        if brand in text_up:
            return brand
    return None


def find_company(text: str) -> Tuple[Optional[str], Optional[str]]:
    up = text.upper()

    # Overrides de nombre: no imponemos VAT si cif_fix es None
    company_override, cif_override = None, None
    for pat, name, cif_fix in VENDOR_OVERRIDES:
        if pat.search(up):
            company_override = name
            cif_override = cif_fix
            break

    # Candidatos VAT del texto (excluyendo cliente)
    vats = find_vat_candidates(text)

    # Detectar marca y priorizar VAT de su catálogo si hay varios
    brand = detect_brand(up)
    supplier_vat = None
    if brand and vats:
        allowed = KNOWN_VATS_BY_BRAND.get(brand, set())
        preferred = [v for v in vats if v in allowed]
        if preferred:
            supplier_vat = preferred[0]

    # Si aún no hay VAT, aplica la heurística general (CIF empresa A-HJNPQRSUVW...)
    if supplier_vat is None and vats:
        vats.sort(
            key=lambda v: 2 if re.match(r"^[A-HJNPQRSUVW]", v) else 1, reverse=True
        )
        supplier_vat = vats[0]

    # Elegir línea de razón social limpia
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    def is_client_line(up_str: str) -> bool:
        if any(k in up_str for k in CLIENT_NAME_KEYWORDS):
            return True
        if any(cid in up_str for cid in CLIENT_IDS):
            return True
        return False

    def good_company_line(s: str) -> bool:
        us = s.upper()
        if is_client_line(us):
            return False
        if any(re.search(suf, us, re.IGNORECASE) for suf in LEGAL_SUFFIXES):
            return True
        if any(h in us for h in KNOWN_VENDOR_HINTS):
            return True
        return False

    company = None
    if supplier_vat:
        idx = up.find(supplier_vat)
        window = text[max(0, idx - 400) : idx + 400] if idx >= 0 else text
        for raw in window.splitlines():
            cand = strip_accents_punct(raw)
            if good_company_line(cand) and len(cand) >= 4:
                company = _clean_company_line(cand)
                break
    if not company:
        for raw in lines:
            cand = strip_accents_punct(raw)
            if good_company_line(cand) and len(cand) >= 4:
                company = _clean_company_line(cand)
                break

    # Aplicar override de nombre (sin imponer VAT si None)
    if company_override:
        company = company_override
    if cif_override is not None:
        supplier_vat = cif_override

    return company, supplier_vat


# -------------------- Fecha / Nº factura --------------------
INVOICE_PATTERNS = [
    r"\bFACTURA(?:\s+SIMPLIFICADA)?\s*(?:N[ºo\.-:]?\s*)?([A-Za-z0-9\-_/\.]+)",
    r"\bN[ºo\.-:]?\s*([A-Za-z0-9\-_/\.]+)\s*(?:FACTURA)\b",
    r"\bNUM(?:ERO|\.|:)?\s*([A-Za-z0-9\-_/\.]+)",
    r"\bFA\s*-\s*([A-Za-z0-9\-]+)",
]
BAD_TOKENS = {"FECHA", "NIF", "IF", "TEL", "TELEFONO"}


def parse_date_text(text: str) -> Optional[str]:
    for pat, fmt in [
        (r"(\d{2}/\d{2}/\d{4})", "%d/%m/%Y"),
        (r"(\d{2}-\d{2}-\d{4})", "%d-%m-%Y"),
        (r"(\d{1,2}/\d{1,2}/\d{2})", "%d/%m/%y"),
    ]:
        m = re.search(pat, text)
        if m:
            try:
                return datetime.strptime(m.group(1), fmt).date().isoformat()
            except Exception:
                pass
    m = re.search(
        r"(\d{1,2})\s+de\s+([A-Za-zÁÉÍÓÚáéíóú]+)\s+de\s+(\d{4})", text, re.IGNORECASE
    )
    if m:
        day = int(m.group(1))
        mon = MONTHS_ES.get(strip_accents_punct(m.group(2)).upper())
        year = int(m.group(3))
        if mon:
            try:
                return datetime(year, mon, day).date().isoformat()
            except Exception:
                pass
    m = re.search(r"Fecha\s+Factura\s*[:#]?\s*(\d{2}/\d{2}/\d{4})", text, re.IGNORECASE)
    if m:
        try:
            return datetime.strptime(m.group(1), "%d/%m/%Y").date().isoformat()
        except Exception:
            pass
    return None


def find_invoice_number(text: str) -> Optional[str]:
    for pat in INVOICE_PATTERNS:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            cand = m.group(1).strip().upper()
            if cand in BAD_TOKENS:
                continue
            return m.group(1).strip()
    return None


# -------------------- Importes --------------------
NUM_RE = re.compile(r"([€]?[0-9][0-9.,]*)")


def _find_value_by_label_patterns(
    lines: List[str], label_patterns: List[str], max_down: int = 4
) -> Optional[float]:
    for pat in label_patterns:
        lab_re = re.compile(pat, flags=re.IGNORECASE)
        for i, ln in enumerate(lines):
            if lab_re.search(ln):
                m = NUM_RE.search(ln)
                if m:
                    v = norm_num(m.group(1))
                    if v is not None:
                        return v
                for j in range(1, max_down + 1):
                    if i + j < len(lines):
                        m2 = NUM_RE.search(lines[i + j])
                        if m2:
                            v = norm_num(m2.group(1))
                            if v is not None:
                                return v
    return None


VAT_ROW_RE = re.compile(
    r"(\d{1,2}(?:[.,]\d{1,2})?)\s*%[^\n\r]*?([0-9][0-9.,]*)[^\n\r]*?([0-9][0-9.,]*)"
)


def parse_vat_rows_sum(
    text: str,
) -> Tuple[Optional[float], Optional[float], Optional[str]]:
    bases, cuotas, rates = [], [], []
    for m in VAT_ROW_RE.finditer(text):
        rate = to_decimal_pct(m.group(1))
        base = norm_num(m.group(2))
        cuota = norm_num(m.group(3))
        if base is not None and cuota is not None:
            bases.append(base)
            cuotas.append(cuota)
            if rate is not None:
                rates.append(rate)
    if bases and cuotas:
        nota = None
        if len(set(rates)) > 1:
            nota = "IVAs: " + "+".join(
                sorted({f"{int(r * 100)}%" for r in rates if r is not None})
            )
        return round(sum(bases), 2), round(sum(cuotas), 2), nota
    return None, None, None


def parse_amounts_generic(text: str) -> Dict[str, Optional[float]]:
    out: Dict[str, Optional[float]] = {
        "importe_base": None,
        "IVA": None,
        "importe_total": None,
        "%IVA": None,
    }
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    base_tbl, iva_tbl, nota_tbl = parse_vat_rows_sum(text)
    base = _find_value_by_label_patterns(lines, LABELS["base"], 4)
    iva = _find_value_by_label_patterns(lines, LABELS["iva"], 4)
    tot = _find_value_by_label_patterns(lines, LABELS["total"], 4)

    # Porcentaje (de etiqueta directa: "IVA 21%" / "VAT 21%")
    pct = None
    for pat in LABELS.get("pct", []):
        m = re.search(pat, text, flags=re.IGNORECASE)
        if m:
            mm = re.search(r"(\d{1,2}(?:[.,]\d{1,2})?)\s*%", m.group(0))
            if mm:
                pct = to_decimal_pct(mm.group(1))
                break

    # Consolidación con tabla
    if base_tbl is not None and (base is None or (base and base_tbl >= base * 0.90)):
        base = base_tbl
    if iva_tbl is not None and (iva is None or (iva and iva_tbl >= iva * 0.90)):
        iva = iva_tbl
    if tot is None and base is not None and iva is not None:
        tot = round(float(base) + float(iva), 2)

    out["importe_base"], out["IVA"], out["importe_total"], out["%IVA"] = (
        base,
        iva,
        tot,
        pct,
    )
    return out


Row = Dict[str, Optional[Union[str, float]]]


# -------------------- Hijo: función top-level (Windows spawn compatible) --------------------
def child_worker(
    q, pdf_path: str, dpi: int, sleep_ms: int, poppler_path: str, tesseract_exe: str
):
    try:
        res = extract_records_from_pdf_inner(
            pdf_path, dpi, sleep_ms, poppler_path, tesseract_exe
        )
        q.put(json.dumps(res))
    except Exception as e:
        p = Path(pdf_path)
        q.put(json.dumps([{"numero_factura": p.name, "Notas": f"Error: {e}"}]))


# -------------------- Extracción principal --------------------
def extract_records_from_pdf_inner(
    pdf_path: str, dpi: int, sleep_ms: int, poppler_path: str, tesseract_exe: str
) -> List[Row]:
    path = Path(pdf_path)
    text = read_pdf_text_native(path)
    ocr_used = False
    if len(text) < 80:
        # Solo OCR zonal cuando no hay texto nativo
        text = read_pdf_text_ocr(
            path,
            dpi=dpi,
            sleep_ms=sleep_ms,
            poppler_path=poppler_path,
            tesseract_exe=tesseract_exe,
        )
        ocr_used = True

    company, cif = find_company(text)
    fecha = parse_date_text(text)
    number = find_invoice_number(text)
    am = parse_amounts_generic(text)
    base, iva, total, pct = (
        am.get("importe_base"),
        am.get("IVA"),
        am.get("importe_total"),
        am.get("%IVA"),
    )

    notas_parts = []
    if ocr_used:
        notas_parts.append("OCR")
    if pct is None and base and iva:
        _, _, nota_tbl = parse_vat_rows_sum(text)
        if nota_tbl:
            notas_parts.append(nota_tbl)

    # Nota genérica para marcas multi-VAT sin VAT visible
    text_up = text.upper()
    brand = detect_brand(text_up)
    if brand and brand in KNOWN_VATS_BY_BRAND and not cif:
        notas_parts.append(f"VAT {brand.title()} no visible en documento (no forzado)")
    # Validación suave si VAT no coincide con catálogo
    if (
        brand
        and brand in KNOWN_VATS_BY_BRAND
        and cif
        and cif not in KNOWN_VATS_BY_BRAND[brand]
    ):
        notas_parts.append(f"VAT {cif} no coincide con catálogo {brand} (revisar)")

    notas = "; ".join([n for n in notas_parts if n])

    if total is None and base is not None and iva is not None:
        total = round(float(base) + float(iva), 2)

    return [
        {
            "fecha_factura": fecha,
            "numero_factura": number if number else path.stem,
            "empresa": strip_accents_punct(company) if company else None,
            "CIF": norm_cif(cif),
            "importe_base": base,
            "%IVA": pct,
            "IVA": iva,
            "importe_total": total,
            "Notas": notas,
        }
    ]


# -------------------- Aislamiento + Timeout --------------------
def run_child_extract(
    pdf_path: Path,
    dpi: int,
    sleep_ms: int,
    poppler_path: str,
    tesseract_exe: str,
    timeout_s: int,
) -> List[Row]:
    import multiprocessing as mp
    from multiprocessing import Process, Queue

    q: Queue = mp.Queue(maxsize=1)
    p = Process(
        target=child_worker,
        args=(q, str(pdf_path), dpi, sleep_ms, poppler_path, tesseract_exe),
        daemon=True,
    )
    p.start()
    p.join(timeout_s)
    if p.is_alive():
        try:
            p.terminate()
            p.join(5)
        except Exception:
            pass
        return [{"numero_factura": pdf_path.name, "Notas": f"Timeout {timeout_s}s"}]
    try:
        data = q.get_nowait()
        return json.loads(data)
    except Exception:
        return [{"numero_factura": pdf_path.name, "Notas": "Sin datos del hijo"}]


# -------------------- Excel --------------------
def format_excel(path_out: Path) -> None:
    try:
        from openpyxl import load_workbook
        from openpyxl.worksheet.worksheet import Worksheet
    except Exception:
        return
    try:
        wb = load_workbook(str(path_out))
    except Exception:
        return
    if not wb.sheetnames:
        return
    ws = wb.active
    if not isinstance(ws, Worksheet):
        wb.save(str(path_out))
        return
    try:
        header_row = (
            list(ws[1]) if (isinstance(ws.max_row, int) and ws.max_row >= 1) else []
        )
        headers: Dict[str, int] = {
            (str(cell.value).strip() if cell.value is not None else f"COL_{idx}"): idx
            for idx, cell in enumerate(header_row, start=1)
        }
    except Exception:
        wb.save(str(path_out))
        return
    col_fecha = headers.get("fecha_factura")
    col_base = headers.get("importe_base")
    col_iva = headers.get("IVA")
    col_total = headers.get("importe_total")
    col_pct = headers.get("%IVA")
    fmt_fecha = "dd/mm/yyyy"
    fmt_moneda = '"€"#,##0.00'
    fmt_pct = "0.00%"
    max_row: int = ws.max_row if isinstance(ws.max_row, int) else 1
    if max_row < 2:
        wb.save(str(path_out))
        return
    if isinstance(col_fecha, int):
        for r in range(2, max_row + 1):
            ws.cell(row=r, column=col_fecha).number_format = fmt_fecha
    for col in (col_base, col_iva, col_total):
        if isinstance(col, int):
            for r in range(2, max_row + 1):
                ws.cell(row=r, column=col).number_format = fmt_moneda
    if isinstance(col_pct, int):
        for r in range(2, max_row + 1):
            ws.cell(row=r, column=col_pct).number_format = fmt_pct
    try:
        wb.save(str(path_out))
    except Exception:
        pass


# -------------------- MAIN --------------------
def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extractor de facturas PDF → Excel (rápido, OCR zonal, timeout)"
    )
    parser.add_argument(
        "--input", "-i", type=str, required=True, help="Carpeta con los PDFs"
    )
    parser.add_argument(
        "--output",
        "-o",
        type=str,
        default=None,
        help="Excel de salida (por defecto: facturas_datos_extraidos.xlsx en la carpeta de entrada)",
    )
    parser.add_argument(
        "--ocr",
        choices=["on", "off"],
        default="on",
        help="OCR si el PDF no tiene texto",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=DEFAULT_DPI,
        help=f"DPI para OCR (por defecto {DEFAULT_DPI})",
    )
    parser.add_argument(
        "--poppler",
        type=str,
        default="",
        help="Ruta a binarios de Poppler si no están en PATH",
    )
    parser.add_argument(
        "--tesseract",
        type=str,
        default="",
        help="Ruta completa a tesseract si no está en PATH",
    )
    parser.add_argument(
        "--log", type=str, default="INFO", choices=["DEBUG", "INFO", "WARNING", "ERROR"]
    )
    parser.add_argument(
        "--sleep-ms",
        type=int,
        default=DEFAULT_SLEEP_MS,
        help=f"Pausa (ms) entre páginas OCR (por defecto {DEFAULT_SLEEP_MS})",
    )
    parser.add_argument(
        "--throttle-every",
        type=int,
        default=DEFAULT_THROTTLE_EVERY,
        help=f"Cada N PDFs aplicar pausa (por defecto {DEFAULT_THROTTLE_EVERY})",
    )
    parser.add_argument(
        "--throttle-ms",
        type=int,
        default=DEFAULT_THROTTLE_MS,
        help=f"Pausa (ms) cuando se cumple throttle-every (por defecto {DEFAULT_THROTTLE_MS})",
    )
    parser.add_argument(
        "--child-timeout-s",
        type=int,
        default=DEFAULT_CHILD_TIMEOUT_S,
        help=f"Timeout por PDF en segundos (por defecto {DEFAULT_CHILD_TIMEOUT_S})",
    )
    args = parser.parse_args()

    logging.basicConfig(
        level=getattr(logging, args.log), format="%(levelname)s: %(message)s"
    )
    input_dir = Path(args.input).expanduser().resolve()
    if not input_dir.exists() or not input_dir.is_dir():
        raise SystemExit(
            f"La carpeta de entrada no existe o no es una carpeta: {input_dir}"
        )
    output_xlsx = (
        Path(args.output).expanduser().resolve()
        if args.output
        else (input_dir / "facturas_datos_extraidos.xlsx")
    )

    pdfs = sorted([p for p in input_dir.glob("**/*.pdf")])
    if not pdfs:
        raise SystemExit(f"No se han encontrado PDFs en: {input_dir}")
    logging.info(f"Procesando {len(pdfs)} PDFs desde {input_dir}")

    rows: List[Row] = []
    count = 0
    for p in pdfs:
        count += 1
        try:
            logging.debug(f"→ {p.name}")
            recs = run_child_extract(
                p,
                dpi=max(72, int(args.dpi)),
                sleep_ms=max(0, int(args.sleep_ms)),
                poppler_path=args.poppler,
                tesseract_exe=args.tesseract,
                timeout_s=max(30, int(args.child_timeout_s)),
            )
            rows.extend(recs)
        except Exception as e:
            logging.error(f"Error en {p.name}: {e}")
        if (
            args.throttle_every
            and args.throttle_ms
            and (count % args.throttle_every == 0)
        ):
            logging.debug(f"Throttle: pausa {args.throttle_ms} ms tras {count} PDFs")
            time.sleep(args.throttle_ms / 1000.0)
        gc.collect()

    df = pd.DataFrame(
        rows,
        columns=[
            "fecha_factura",
            "numero_factura",
            "empresa",
            "CIF",
            "importe_base",
            "%IVA",
            "IVA",
            "importe_total",
            "Notas",
        ],
    )
    try:
        df["fecha_sort"] = pd.to_datetime(df["fecha_factura"], errors="coerce")
        df = df.sort_values(["fecha_sort", "numero_factura"]).drop(
            columns=["fecha_sort"]
        )
        df["fecha_factura"] = pd.to_datetime(df["fecha_factura"], errors="coerce")
    except Exception:
        pass

    df.to_excel(output_xlsx, index=False, engine="openpyxl")
    format_excel(output_xlsx)
    logging.info(f"Excel generado: {output_xlsx}")


if __name__ == "__main__":
    # Recomendado en Windows/VSCode debug para multiprocessing
    try:
        import multiprocessing as mp

        mp.freeze_support()
    except Exception:
        pass
    main()
