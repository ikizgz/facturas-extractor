"""Pruebas unitarias para el parser de Salvador Escoda S.A."""

from __future__ import annotations

from pathlib import Path

from gestor_facturas.providers.salvador_escoda import SalvadorEscodaParser


def test_puede_procesar_salvador_escoda():
    """Verifica que el parser detecta correctamente los textos identificativos de Salvador Escoda."""
    parser = SalvadorEscodaParser()

    texto_valido = "SALVADOR ESCODA S.A. - C.I.F. A-08710006"
    texto_valido_alternativo = "Suministro de productos... C.I.F. A08710006"
    texto_invalido = "FACTURA DE AMAZON EU S.À R.L."

    assert parser.puede_procesar(texto_valido) is True
    assert parser.puede_procesar(texto_valido_alternativo) is True
    assert parser.puede_procesar(texto_invalido) is False


def test_extraer_datos_factura_normal():
    """Simula la extracción de datos de una factura estándar de Salvador Escoda."""
    parser = SalvadorEscodaParser()

    texto_simulado = """
    SALVADOR ESCODA S.A.
    FACTURA 10092418
    Fecha Factura: 10/07/2026
    BASE IMPONIBLE 44,55
    21,00 %IVA 9,36
    TOTAL 53,91
    """

    datos = parser.extraer_datos(texto_simulado, Path("factura_normal.pdf"))

    assert datos["numero_factura"] == "10092418"
    assert datos["fecha"] == "2026-07-10"
    assert datos["proveedor"] == "SALVADOR ESCODA S.A."
    assert datos["nif_cif"] == "A08710006"
    assert datos["importe_base"] == 44.55
    assert datos["tasa_iva"] == 21.0
    assert datos["importe_iva"] == 9.36
    assert datos["importe_total"] == 53.91
    assert datos["notas"] == ""
    # imprimimos los datos extraídos para inspección visual
    print("\n--- DATOS EXTRAÍDOS ---")
    for k, v in datos.items():
        print(f"{k}: {v}")


def test_extraer_datos_factura_abono():
    """Simula la extracción de datos de una factura rectificativa / abono de Salvador Escoda."""
    parser = SalvadorEscodaParser()

    texto_simulado = """
    SALVADOR ESCODA S.A.
    FACTURA RECTIFICATIVA 10095637
    Fecha Factura: 13/07/2026
    Abono por devolución Factura: 10092411
    BASE IMPONIBLE -41,16
    21,00 %IVA -8,64
    TOTAL -49,80
    """

    datos = parser.extraer_datos(texto_simulado, Path("factura_abono.pdf"))

    assert datos["numero_factura"] == "10095637"
    assert datos["fecha"] == "2026-07-13"
    assert datos["proveedor"] == "SALVADOR ESCODA S.A."
    assert datos["nif_cif"] == "A08710006"
    assert datos["importe_base"] == -41.16
    assert datos["tasa_iva"] == 21.0
    assert datos["importe_iva"] == -8.64
    assert datos["importe_total"] == -49.80
    # Verificamos que la traza de notas detecta el abono y la factura original
    assert "Factura Abono / Rectificativa" in datos["notas"]
    assert "10092411" in datos["notas"]
    # imprimimos los datos extraídos para inspección visual
    print("\n--- DATOS EXTRAÍDOS ---")
    for k, v in datos.items():
        print(f"{k}: {v}")
