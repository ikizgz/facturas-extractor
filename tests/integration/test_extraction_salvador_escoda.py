"""Pruebas de integración para verificar la extracción con PDFs reales."""

from __future__ import annotations

from pathlib import Path

import pdfplumber
import pytest

from gestor_facturas.providers.salvador_escoda import SalvadorEscodaParser


def test_extraccion_pdf_real_salvador_escoda():
    """Prueba el pipeline completo de lectura y extracción con un PDF real de Salvador Escoda."""
    # Ruta relativa o absoluta hacia donde guardes una de tus facturas reales de prueba
    # Puedes ajustar la ruta al PDF que subiste o tienes en tu proyecto
    ruta_pdf = Path(
        "tests/integration/factura_salvador_escoda_2.pdf"
    )  # Cambia esto por la ruta real si la tienes colocada aquí

    # Si no tienes el PDF físicamente en esa ruta exacta para el test automatizado,
    # este test se saltará automáticamente para que no rompa el `pytest` general.
    if not ruta_pdf.exists():
        pytest.skip(
            f"No se encontró el archivo PDF real en {ruta_pdf} para la prueba de integración."
        )

    # 1. Extracción de texto con pdfplumber (igual que lo hará tu pipeline principal)
    texto_extraido = ""
    with pdfplumber.open(ruta_pdf) as pdf:
        for pagina in pdf.pages:
            texto_extraido += pagina.extract_text() or ""

    # 2. Instanciamos el parser y comprobamos que identifica el documento
    parser = SalvadorEscodaParser()
    assert parser.puede_procesar(texto_extraido) is True

    # 3. Extraemos los datos estructurados
    datos = parser.extraer_datos(texto_extraido, ruta_pdf)

    # 4. Validamos que se hayan extraído campos obligatorios correctamente (que no estén vacíos o por defecto)
    assert datos["numero_factura"] != "DESCONOCIDO"
    assert datos["nif_cif"] == "A08710006"
    assert isinstance(datos["importe_base"], float)
    assert isinstance(datos["tasa_iva"], float)
    assert isinstance(datos["importe_iva"], float)
    assert isinstance(datos["importe_total"], float)
    # Imprimimos los datos extraídos para inspección visual durante la prueba
    print("\n--- DATOS EXTRAÍDOS DEL PDF REAL ---")
    for k, v in datos.items():
        print(f"{k}: {v}")
