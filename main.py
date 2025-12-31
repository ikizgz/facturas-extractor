#!/usr/bin/env python3
# -*- coding: utf-8 -*-


# main.py
from __future__ import annotations

import argparse
import logging
import re
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
from PyPDF2 import PdfReader

# ---------------------------
# CONFIGURACIÓN OCR (opcional)
# ---------------------------
ENABLE_OCR = False
POPPLER_PATH: Optional[str] = None
TESSERACT_EXE: Optional[str] = None
OCR_LANG = "spa+eng"  # 'spa' si prefieres solo español

# ---------------------------
# CATÁLOGOS / CONSTANTES
# ---------------------------

CLIENT_IDS = {
    "J99198285",  # ACR S.C. (tu cliente) — normalizado
    # Si alguna factura muestra prefijo país, añade "ESJ99198285"
}

MONTHS_ES = {
    "enero": 1,
    "febrero": 2,
    "marzo": 3,
    "abril": 4,
    "mayo": 5,
    "junio": 6,
    "julio": 7,
    "agosto": 8,
    "septiembre": 9,
    "setiembre": 9,
    "octubre": 10,
    "noviembre": 11,
    "diciembre": 12,
}

VAT_RE = r"([A-Z]{1,2}[A-Z0-9\-]{2,14}|[A-Za-z]\-?\d{8}|\d{8}[A-Za-z])"

# --- Lista completa de proveedores (CIF/NIF/VAT → razón social normalizada) ---
PROVIDERS: Dict[str, str] = {
    "B05478946": "AJIRO SERVICIOS SL",
    "18014950Q": "ALBERTO GOMEZ ARDURA",
    "A28581882": "ALCAMPO SA",
    "48663511L": "ALFREDO JARA ROBLES",
    "B10547115": "ALTON CAPITAL 2000 SL",
    "ESN0186600C": "AMAZON BUSINESS EU SARL",
    "ESW0264006H": "AMAZON BUSINESS EU SARL SUCURSAL EN ESPANA",
    "DE814584193": "AMAZON EU SARL NIEDERLASSUNG DEUTSCHLAND",
    "FR12487773327": "AMAZON EU SARL SUCCURSALE FRANCAISE",
    "IT08973230967": "AMAZON EU SARL SUCCURSALE ITALIANA",
    "ESW0184081H": "AMAZON EU SARL SUCURSAL EN ESPANA",
    "PL5262907815": "AMAZON EU SARL SUCURSAL EN POLONIA",
    "N7204904B": "AN AN BEAUTY LIMITED",
    "A18096511": "ARAGONESA DE SERVICIOS ITV SA",
    "B29351616": "AUCORE SL",
    "B50187459": "B A S E SL",
    "FR76889152914": "BLUFITS SPORTS LIMITED",
    "B20089298": "CARZA GIPUZKOA MOTOR SLU",
    "A28013050": "CASER SA",
    "CY10356197E": "COMWORX LTD",
    "B29733870": "COVEY ALQUILER SL",
    "B99418808": "DECORACIONES IBOR SL",
    "25441059T": "DIEGO ALVAREZ MATEO",
    "A08338188": "DIOTRONIC SA",
    "B50935055": "DURBAN ALQUILER SL",
    "A50066190": "DURBAN MAQUINARIA CONSTRUCCION SA",
    "ESB16863532": "ECOM CIRCLE SL",
    "B45371655": "ELECTRONICA JOPAL SLU",
    "B06600035": "ELECTRONICA REY SL",
    "17151178D": "EMILIO GOMEZ GARCIA",
    "B30353254": "EMMETI IBERICA SLU",
    "Q2826004J": "FABRICA NACIONAL DE MONEDA Y TIMBRE",
    "A50047646": "FERRETERIA ARIES SA",
    "B24721524": "FERRETERIA ONLINE VTC SL",
    "B50616481": "FERRETERIA ROYMAR SL",
    "A79783254": "FEU VERT IBERICA SAU",
    "A50021831": "FONTANEROS SA",
    "B44627222": "GADIRIA INVEST SL",
    "ESB54651047": "GESCOM SPORT SL",
    "B63730071": "GESTION INTEGRAL ALMACENES SL",
    "A50307545": "HIDRONEUMATICA ANLO SA",
    "B95294757": "HOSTINET SL",
    "CY10301365E": "HOSTINGER INTERNATIONAL LIMITED",
    "B58274705": "IDIOMUND SL",
    "B50040005": "INDUSTRIAS SANITARIAS REUNIDAS SL",
    "B99399552": "JAB ARAGON DAM SL",
    "B88190830": "KOOLAIR SL",
    "N0049739F": "KW COMMERCE GMBH",
    "B66810045": "LA ESPECIALISTA DISTRIBUIDORA DE SISTEMAS CONSTRUCTIVOS SL",
    "B99134520": "LA TINTA ARAGONESA SL",
    "B84818442": "LEROY MERLIN ESPANA SLU",
    "ESA62348131": "MEDIA MARKT ZARAGOZA SA",
    "X4081053W": "MIRCEA DRAGOMIR",
    "N0072209J": "NATIONAL PEN PROMOTIONAL PRODUCTS LTD",
    "B40027062": "NEPTUNO SL",
    "A78119773": "NOROTO SAU",
    "ESB42738690": "OFAM PHARMAPLACE SL",
    "B15713407": "OFIPRO SOLUCIONES SL",
    "A82009812": "ORANGE ESPANA SAU",
    "B83107037": "OSIRA CLIMATIZACION SL",
    "B53166419": "SISTAC ILS SL",
    "B99120123": "PETROSIF 2006 SLU",
    "A58868324": "PHIL NOBEL SA",
    "A50050442": "PINTURAS ORDESA SA",
    "A28278026": "PLATAFORMA COMERCIAL DE RETAIL SAU",
    "B93275394": "PLENOIL SL",
    "ESB87064093": "QUIEROSOFT SL",
    "B02230407": "RCR PROYECTOS DE SOFTWARE SLU",
    "A08388811": "REMLE SA",
    "A80298839": "REPSOL SOLUCIONES ENERGETICAS SA",
    "B50851559": "RODRIGO LABORAL 2000 SL",
    "B99385502": "SALTOKI SUMINISTROS ZARAGOZA SL",
    "A08710006": "SALVADOR ESCODA SA",
    "B86533981": "SATURDAY TRADE SL",
    "X6526242S": "XIAOJIE WANG",
    "B50337757": "TALLERES SANTA OROSIA SL",
    "A82018474": "TELEFONICA DE ESPANA SAU",
    "A78923125": "TELEFONICA MOVILES ESPANA SAU",
    "NL812139513B01": "VISTAPRINT BV",
    "A20004008": "VIUDA DE LONDAIZ Y SOBRINOS DE L MERCADER SA",
    "A50049204": "VOLTAMPER SA",
    "B58844218": "YESYFORMA EUROPA SL",
    "ESB42651935": "ZAPATERA ALICANTINA SL",
    "A50047505": "ZOILO RIOS SA",
}

ALIASES: Dict[str, str] = {
    "GASOLEOS MERCADAIZ": "A20004008",
    "SALTOKI": "B99385502",
    "SORPRESA HOGAR": "X6526242S",
    "O2": "A82018474",
    "MOVISTAR": "A78923125",
    "PLENOIL": "B93275394",
    "NORAUTO": "A78119773",
    "OFIPRO": "B15713407",
    "INDUSAN": "B50040005",
    "REMLE": "A08388811",
    "QUIEROSOFT": "ESB87064093",
    "EUREKA PARTS": "A58868324",
    "PC BOX": "B53166419",
    "ALCAMPO": "A28581882",
    "LEROY MERLIN": "B84818442",
}

# ---------------------------
# UTILIDADES
# ---------------------------


def strip_accents_punct(s: str) -> str:
    s = s or ""
    s = "".join(
        c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn"
    )
    s = re.sub(r"[^A-Za-z0-9 ]", " ", s)
    s = re.sub(r"\s+", " ", s).strip().upper()
    return s


def norm_cif(s: Optional[str]) -> Optional[str]:
    if not s:
        return s
    return s.replace("-", "").replace(" ", "").strip().upper()


def norm_num(s: Optional[str]) -> Optional[float]:
    if s is None:
        return None
    s = str(s).strip()
    if not s:
        return None
    for sym in ("€", "EUR", " ", "\u00a0"):
        s = s.replace(sym, "")
    s = s.replace("%", "")
    try:
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        return float(s)
    except (ValueError, TypeError):
        return None


def norm_percent(s: Optional[str]) -> Optional[float]:
    if s is None:
        return None
    s0 = str(s).strip().replace(" ", "").replace("\u00a0", "")
    if not s0:
        return None
    if s0.endswith("%"):
        s0 = s0[:-1]
    val = norm_num(s0)
    return round(val / 100.0, 6) if val is not None else None


# ---------------------------
# EXTRACCIÓN DE TEXTO (PDF/OCR)
# ---------------------------


def read_pdf_text(path: Path) -> str:
    """Extrae texto del PDF. Si falla y OCR está activo, usa OCR con imports locales."""
    # 1) Texto nativo
    try:
        reader = PdfReader(str(path))
        return "\n".join([(p.extract_text() or "") for p in reader.pages])
    except Exception as e:
        logging.warning(f"No se pudo leer texto nativo: {path.name} ({e})")

    # 2) OCR (si está activo)
    if ENABLE_OCR:
        try:
            import pytesseract
            from pdf2image import convert_from_path

            # poppler path
            kwargs = {}
            if POPPLER_PATH:
                kwargs["poppler_path"] = POPPLER_PATH
            images = convert_from_path(str(path), **kwargs)
            # tesseract path
            if TESSERACT_EXE:
                pytesseract.pytesseract.tesseract_cmd = TESSERACT_EXE
            txt = []
            for img in images:
                txt.append(pytesseract.image_to_string(img, lang=OCR_LANG))
            return "\n".join(txt)
        except Exception as e:
            logging.error(f"OCR falló en {path.name}: {e}")

    return ""


# ---------------------------
# PARSEO DE FECHAS / CIF / EMPRESA / Nº FACTURA
# ---------------------------


def parse_date_any(text: str, filename: str = "") -> Optional[str]:
    m = re.search(
        r"(\d{1,2})\s+de\s+([A-Za-záéíóú]+)\s+de\s+(\d{4})", text, re.IGNORECASE
    )
    if m:
        day = int(m.group(1))
        month_name = strip_accents_punct(m.group(2).lower())
        mon = MONTHS_ES.get(month_name)
        year = int(m.group(3))
        if mon:
            try:
                return datetime(year, mon, day).date().isoformat()
            except ValueError:
                pass
    for pat, fmt in (
        (r"(\d{2}/\d{2}/\d{4})", "%d/%m/%Y"),
        (r"(\d{2}-\d{2}-\d{4})", "%d-%m-%Y"),
    ):
        m = re.search(pat, text)
        if m:
            try:
                return datetime.strptime(m.group(1), fmt).date().isoformat()
            except ValueError:
                pass
    m = re.search(r"del\s+(\d{4}-\d{2}-\d{2})", filename)
    if m:
        return m.group(1)
    return None


def find_vat(text: str) -> Optional[str]:
    ids = re.findall(VAT_RE, text)

    # VAT con prefijo país
    for cid in ids:
        cid_norm = norm_cif(cid)
        if (
            cid_norm
            and cid_norm.startswith(("ES", "FR", "DE", "IT", "NL", "PL"))
            and cid_norm not in CLIENT_IDS
        ):
            return cid_norm

    # Resto
    for cid in ids:
        cid_norm = norm_cif(cid)
        if cid_norm and cid_norm not in CLIENT_IDS:
            return cid_norm

    return None


def find_company(text: str, filename: str) -> tuple[Optional[str], Optional[str]]:
    vat = find_vat(text)
    if vat and vat in PROVIDERS:
        return PROVIDERS[vat], vat

    # razón social explícita
    for cif_key, legal_name in PROVIDERS.items():
        pattern = r"\b" + re.escape(legal_name).replace(r"\ ", r"\s+") + r"\b"
        if re.search(pattern, text, re.IGNORECASE):
            return legal_name, cif_key

    # alias → CIF → razón social
    for alias, alias_cif in ALIASES.items():
        pattern = r"\b" + re.escape(alias).replace(r"\ ", r"\s+") + r"\b"
        if re.search(pattern, text, re.IGNORECASE):
            legal = PROVIDERS.get(alias_cif)
            if legal:
                return legal, alias_cif

    base = Path(filename).name
    company = strip_accents_punct(base.split(",")[0])
    return company, None


def find_invoice_number(text: str, filename: str) -> Optional[str]:
    pats = [
        r"Factura\s*N[ºo]?\s*[:\-]?\s*([A-Za-z0-9/\-]{3,})",
        r"FACTURA\s*([A-Za-z0-9/\-]{3,})",
        r"FA\s*-\s*([A-Za-z0-9\-]{2,})",
        r"(ES[0-9A-Z]{8,})",
    ]
    for p in pats:
        m = re.search(p, text, re.IGNORECASE)
        if m:
            return m.group(1).strip()
    base = Path(filename).name
    m = re.search(r",\s*([^,]+?)\s+del\s+\d{4}-\d{2}-\d{2}", base)
    if m:
        return m.group(1).strip()
    return None


# ---------------------------
# PARSEO DE IMPORTES / IVA
# ---------------------------


def find_number_after_label(text: str, label_regex: str) -> Optional[float]:
    m = re.search(label_regex, text, re.IGNORECASE)
    if not m:
        return None
    tail = text[m.end() : m.end() + 200]
    mnum = re.search(r"([0-9]+(?:[\.,][0-9]{1,4})?)", tail)
    if not mnum:
        return None
    return norm_num(mnum.group(1))


def find_percentage_in_text(text: str) -> Optional[float]:
    m = re.search(r"([0-9]{1,2}(?:[\.,][0-9]{1,2})?)\s*%", text)
    if m:
        return norm_percent(m.group(1) + "%")
    return None


def compute_pct(base: Optional[float], iva: Optional[float]) -> Optional[float]:
    if base is not None and iva is not None and base != 0:
        return round(iva / base, 6)
    return None


def find_all_percentages(text: str) -> List[float]:
    raw = re.findall(r"([0-9]{1,2}(?:[\.,][0-9]{1,2})?)\s*%", text)
    vals: List[float] = []
    for s in raw:
        v = norm_percent(s + "%")
        if v is not None:
            vals.append(round(v * 100.0, 2))  # 21.00, 10.00...
    return vals


def detect_varios_ivas(
    text: str, base: Optional[float], iva: Optional[float], pct: Optional[float]
) -> bool:
    percents = {p for p in find_all_percentages(text)}
    if len(percents) >= 2:
        return True
    if re.search(r"\bexento\b|\bexenci[oó]n\b|\bno\s+sujeto\b", text, re.IGNORECASE):
        if percents or pct is None:
            return True
    return False


def parse_amounts_by_company(
    text: str, company: Optional[str]
) -> Dict[str, Optional[float]]:
    comp = (company or "").lower()

    def grab(label: str) -> Optional[float]:
        return find_number_after_label(text, label)

    if "alcampo" in comp:
        base = grab(r"Total\s*Base\s*Imponible")
        iva = grab(r"Total\s*Impuesto")
        total = grab(r"Total\s*Factura")
        pct = find_percentage_in_text(text) or compute_pct(base, iva)
        return {
            "importe_base": base,
            "IVA": iva,
            "importe_total": total,
            "porcentaje_iva": pct,
        }

    if "leroy" in comp:
        base = grab(r"Total\s*SI")
        iva = grab(r"Total\s*IVA/IGIC/IPSI")
        total = grab(r"Total\s*TII")
        pct = find_percentage_in_text(text) or compute_pct(base, iva)
        return {
            "importe_base": base,
            "IVA": iva,
            "importe_total": total,
            "porcentaje_iva": pct,
        }

    if "viuda de londaiz" in comp or "mercader" in comp or "gasoleos" in comp:
        base = grab(r"BASE\s*IMPONIBLE")
        iva = grab(r"TOTAL\s*I\.V\.A")
        total = grab(r"TOTAL\s*FACTURA")
        pct = find_percentage_in_text(text) or compute_pct(base, iva)
        return {
            "importe_base": base,
            "IVA": iva,
            "importe_total": total,
            "porcentaje_iva": pct,
        }

    base = (
        grab(r"Total\s*Base\s*Imponible")
        or grab(r"BASE\s*IMPONIBLE")
        or grab(r"Total\s*SI")
    )
    iva = grab(r"Total\s*IVA") or grab(r"TOTAL\s*I\.V\.A") or grab(r"Total\s*Impuesto")
    total = grab(r"Total\s*Factura") or grab(r"TOTAL\s*FACTURA") or grab(r"Total\s*TII")
    pct = find_percentage_in_text(text) or compute_pct(base, iva)
    return {
        "importe_base": base,
        "IVA": iva,
        "importe_total": total,
        "porcentaje_iva": pct,
    }


# ---------------------------
# EXTRACCIÓN POR PDF
# ---------------------------


def extract_from_pdf(path: Path) -> Dict[str, Optional[str | float]]:
    text = read_pdf_text(path)
    fecha = parse_date_any(text, filename=str(path))
    company, cif = find_company(text, str(path))
    number = find_invoice_number(text, str(path))
    amounts = parse_amounts_by_company(text, company)

    empresa_norm = strip_accents_punct(company) if company else None
    cif_norm = norm_cif(cif)
    base = amounts.get("importe_base")
    iva = amounts.get("IVA")
    total = amounts.get("importe_total")
    pct = amounts.get("porcentaje_iva")

    if total is None and base is not None and iva is not None:
        total = round(base + iva, 2)

    varios = detect_varios_ivas(text, base, iva, pct)
    notas = "VARIOS IVAs" if varios else ""

    return {
        "fecha_factura": fecha,
        "numero_factura": number,
        "empresa": empresa_norm,
        "CIF": cif_norm,
        "importe_base": base,
        "%IVA": pct,  # decimal 0.21 → 21 %
        "IVA": iva,
        "importe_total": total,
        "Notas": notas,
    }


# ---------------------------
# FORMATEO EXCEL
# ---------------------------


def format_excel(path_out: Path) -> None:
    from openpyxl import load_workbook

    wb = load_workbook(str(path_out))
    ws = wb.active
    assert ws is not None, "No se pudo obtener la hoja activa del Excel."

    headers = {cell.value: idx for idx, cell in enumerate(ws[1], start=1)}

    col_fecha = headers.get("fecha_factura")
    col_base = headers.get("importe_base")
    col_iva = headers.get("IVA")
    col_total = headers.get("importe_total")
    col_pct = headers.get("%IVA")

    fmt_fecha = "dd/mm/yyyy"
    fmt_moneda = '"€"#,##0.00'
    fmt_pct = "0.00%"

    max_row = ws.max_row or 1

    if col_fecha:
        for r in range(2, max_row + 1):
            ws.cell(row=r, column=col_fecha).number_format = fmt_fecha

    for col in (col_base, col_iva, col_total):
        if col:
            for r in range(2, max_row + 1):
                ws.cell(row=r, column=col).number_format = fmt_moneda

    if col_pct:
        for r in range(2, max_row + 1):
            ws.cell(row=r, column=col_pct).number_format = fmt_pct


# ---------------------------
# CLI
# ---------------------------


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extractor de facturas PDF → Excel consolidado"
    )
    parser.add_argument(
        "--input",
        "-i",
        type=str,
        required=True,
        help="Carpeta donde están los PDFs de facturas",
    )
    parser.add_argument(
        "--output",
        "-o",
        type=str,
        default=None,
        help="Ruta del Excel de salida (por defecto: facturas_datos_extraidos.xlsx en la carpeta de entrada)",
    )
    parser.add_argument(
        "--ocr",
        choices=["on", "off"],
        default="off",
        help="Activa OCR si el PDF no tiene texto (Poppler + Tesseract)",
    )
    parser.add_argument(
        "--poppler",
        type=str,
        default=None,
        help="Ruta a la carpeta bin de Poppler (opcional si no está en PATH)",
    )
    parser.add_argument(
        "--tesseract",
        type=str,
        default=None,
        help="Ruta completa a tesseract.exe (opcional si no está en PATH)",
    )
    parser.add_argument(
        "--log", type=str, default="INFO", choices=["DEBUG", "INFO", "WARNING", "ERROR"]
    )
    args = parser.parse_args()

    logging.basicConfig(
        level=getattr(logging, args.log), format="%(levelname)s: %(message)s"
    )

    global ENABLE_OCR, POPPLER_PATH, TESSERACT_EXE
    ENABLE_OCR = args.ocr == "on"
    POPPLER_PATH = args.poppler
    TESSERACT_EXE = args.tesseract

    input_dir = Path(args.input).expanduser().resolve()
    if not input_dir.exists() or not input_dir.is_dir():
        raise SystemExit(
            f"La carpeta de entrada no existe o no es una carpeta: {input_dir}"
        )

    default_name = "facturas_datos_extraidos.xlsx"
    output_xlsx = (
        Path(args.output).expanduser().resolve()
        if args.output
        else (input_dir / default_name)
    )

    pdfs = sorted([p for p in input_dir.glob("**/*.pdf")])
    if not pdfs:
        raise SystemExit(f"No se han encontrado PDFs en: {input_dir}")

    logging.info(f"Procesando {len(pdfs)} PDFs desde {input_dir}")
    rows: List[Dict[str, Optional[str | float]]] = []

    for p in pdfs:
        try:
            rec = extract_from_pdf(p)
            rows.append(rec)
        except Exception as e:
            logging.error(f"Error en {p.name}: {e}")

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

    # Ordenar y preparar fechas para Excel
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
    main()
