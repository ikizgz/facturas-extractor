#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# main.py
import argparse
import logging
import re
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd
from PyPDF2 import PdfReader

# --- Configuración opcional de OCR (por defecto OFF) ---
ENABLE_OCR = False
try:
    if ENABLE_OCR:
        import pdf2image
        import pytesseract
except Exception:
    ENABLE_OCR = False

# --- Diccionario de proveedores (alias → razón social) ---
PROVIDERS = {
    "A28581882": "ALCAMPO SA",
    "B84818442": "LEROY MERLIN ESPANA SLU",
    "A20004008": "VIUDA DE LONDAIZ Y SOBRINOS DE L MERCADER SA",
    "18014950Q": "ALBERTO GOMEZ ARDURA",
    "A18096511": "ARAGONESA DE SERVICIOS ITV SA",
    "A80298839": "REPSOL SOLUCIONES ENERGETICAS SA",
    "A78119773": "NOROTO SAU",
    "X6526242S": "XIAOJIE WANG",
    "B50040005": "INDUSTRIAS SANITARIAS REUNIDAS SL",
    "B02230407": "RCR PROYECTOS DE SOFTWARE SLU",
    "A82018474": "TELEFONICA DE ESPANA SAU",
    "ESW0264006H": "AMAZON BUSINESS EU SARL SUCURSAL EN ESPANA",
    "N2764308I": "CREATORS OF TOMORROW GMBH",
    # Añade aquí el resto de tu lista…
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


# --- Utilidades ---
def strip_accents_punct(s: str) -> str:
    s = s or ""
    s = "".join(
        c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn"
    )
    s = re.sub(r"[^A-Za-z0-9 ]", " ", s)
    s = re.sub(r"\s+", " ", s).strip().upper()
    return s


def norm_cif(s: str) -> str:
    if not s:
        return s
    return s.replace("-", "").strip().upper()


def parse_date_any(text: str, filename: str = "") -> str | None:
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


def read_pdf_text(path: Path) -> str:
    try:
        reader = PdfReader(str(path))
        return "\n".join([(p.extract_text() or "") for p in reader.pages])
    except Exception as e:
        logging.warning(f"No se pudo leer texto nativo: {path} ({e})")
        if ENABLE_OCR:
            try:
                images = pdf2image.convert_from_path(str(path))
                txt = []
                for img in images:
                    txt.append(pytesseract.image_to_string(img, lang="spa"))
                return "\n".join(txt)
            except Exception as e2:
                logging.error(f"OCR falló: {e2}")
                return ""
        return ""


def find_vat(text: str) -> str | None:
    ids = re.findall(VAT_RE, text)
    for cid in ids:
        cid = norm_cif(cid)
        if cid.startswith(("ES", "FR", "DE", "IT", "NL", "PL")):
            return cid
    for cid in ids:
        cid = norm_cif(cid)
        if not cid.startswith("J"):
            return cid
    return None


def find_company(text: str, filename: str) -> tuple[str | None, str | None]:
    vat = find_vat(text)
    company = None
    if vat and vat in PROVIDERS:
        company = PROVIDERS[vat]
    if not company:
        if re.search(r"ALCAMPO\s*,?\s*S\.?A\.?", text, re.IGNORECASE):
            company = "ALCAMPO SA"
        elif re.search(
            r"Leroy\s+Merlin\s+Espa(?:na|ña)\s*S\.?\s*L\.?\s*U\.?", text, re.IGNORECASE
        ):
            company = "LEROY MERLIN ESPANA SLU"
        elif re.search(r"Viuda\s+de\s+Londaiz.*Mercader", text, re.IGNORECASE):
            company = "VIUDA DE LONDAIZ Y SOBRINOS DE L MERCADER SA"
    if not company:
        base = Path(filename).name
        company = strip_accents_punct(base.split(",")[0])
    return company, vat


def find_invoice_number(text: str, filename: str) -> str | None:
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


def norm_num(s: str) -> float | None:
    if s is None:
        return None
    s = s.strip()
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except:
        return None


def find_number_after_label(text: str, label_regex: str) -> float | None:
    m = re.search(label_regex, text, re.IGNORECASE)
    if not m:
        return None
    tail = text[m.end() : m.end() + 200]
    mnum = re.search(r"([0-9]+(?:[\.,][0-9]{1,2})?)", tail)
    if not mnum:
        return None
    return norm_num(mnum.group(1))


def find_percentage(
    text: str, base: float | None = None, iva: float | None = None
) -> float | None:
    m = re.search(r"([0-9]{1,2}(?:[\.,][0-9]{1,2})?)\s*%", text)
    if m:
        return float(m.group(1).replace(",", ".")) / 100.0
    if base and iva and base != 0:
        return round(iva / base, 4)
    return None


def parse_amounts_by_company(text: str, company: str) -> dict:
    comp = (company or "").lower()

    def grab(label):
        return find_number_after_label(text, label)

    if "alcampo" in comp:
        base = grab(r"Total\s*Base\s*Imponible")
        iva = grab(r"Total\s*Impuesto")
        total = grab(r"Total\s*Factura")
        pct = find_percentage(text, base, iva)
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
        pct = find_percentage(text, base, iva)
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
        pct = find_percentage(text, base, iva)
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
    pct = find_percentage(text, base, iva)
    return {
        "importe_base": base,
        "IVA": iva,
        "importe_total": total,
        "porcentaje_iva": pct,
    }


def extract_from_pdf(path: Path) -> dict:
    text = read_pdf_text(path)
    fecha = parse_date_any(text, filename=str(path))
    company, cif = find_company(text, str(path))
    number = find_invoice_number(text, str(path))
    amounts = parse_amounts_by_company(text, company)
    empresa_norm = strip_accents_punct(company) if company else None
    cif_norm = norm_cif(cif)
    pct = amounts.get("porcentaje_iva")
    base = amounts.get("importe_base")
    iva = amounts.get("IVA")
    total = amounts.get("importe_total")
    if total is None and base is not None and iva is not None:
        total = round(base + iva, 2)
    return {
        "fecha_factura": fecha,
        "numero_factura": number,
        "empresa": empresa_norm,
        "CIF": cif_norm,
        "importe_base": base,
        "%IVA": pct,  # decimal: 0.21 ⇒ 21 %
        "IVA": iva,
        "importe_total": total,
        "source_file": path.name,
    }


def format_excel(path_out: Path):
    from openpyxl import load_workbook

    wb = load_workbook(str(path_out))
    ws = wb.active
    headers = {cell.value: idx for idx, cell in enumerate(ws[1], start=1)}
    col_fecha = headers.get("fecha_factura")
    col_base = headers.get("importe_base")
    col_iva = headers.get("IVA")
    col_total = headers.get("importe_total")
    col_pct = headers.get("%IVA")
    fmt_fecha = "dd/mm/yyyy"
    fmt_moneda = '"€"#,##0.00'
    fmt_pct = "0.00%"
    max_row = ws.max_row
    # Fecha
    if col_fecha:
        for r in range(2, max_row + 1):
            ws.cell(row=r, column=col_fecha).number_format = fmt_fecha
    # Moneda
    for col in (col_base, col_iva, col_total):
        if col:
            for r in range(2, max_row + 1):
                ws.cell(row=r, column=col).number_format = fmt_moneda
    # Porcentaje
    if col_pct:
        for r in range(2, max_row + 1):
            ws.cell(row=r, column=col_pct).number_format = fmt_pct
    wb.save(str(path_out))


# --- CLI ---
def main():
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
        help="Ruta del Excel de salida (por defecto se crea en la carpeta de entrada)",
    )
    parser.add_argument(
        "--ocr",
        choices=["on", "off"],
        default="off",
        help="Activa OCR si el PDF no tiene texto (requiere instalación)",
    )
    parser.add_argument(
        "--log", type=str, default="INFO", choices=["DEBUG", "INFO", "WARNING", "ERROR"]
    )
    args = parser.parse_args()

    logging.basicConfig(
        level=getattr(logging, args.log), format="%(levelname)s: %(message)s"
    )

    global ENABLE_OCR
    ENABLE_OCR = args.ocr == "on"

    input_dir = Path(args.input).expanduser().resolve()
    if not input_dir.exists() or not input_dir.is_dir():
        raise SystemExit(
            f"La carpeta de entrada no existe o no es una carpeta: {input_dir}"
        )

    # ⚠️ CAMBIO CLAVE: si no pasas --output, creamos por defecto el Excel en la carpeta de entrada
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

    rows = []
    for p in pdfs:
        try:
            rows.append(extract_from_pdf(p))
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
            "source_file",
        ],
    )

    try:
        df["fecha_sort"] = pd.to_datetime(df["fecha_factura"])
        df = df.sort_values(["fecha_sort", "numero_factura"]).drop(
            columns=["fecha_sort"]
        )
        df["fecha_factura"] = pd.to_datetime(df["fecha_factura"])
    except Exception:
        pass

    df.to_excel(output_xlsx, index=False, engine="openpyxl")
    format_excel(output_xlsx)
    logging.info(f"Excel generado: {output_xlsx}")


if __name__ == "__main__":
    main()
