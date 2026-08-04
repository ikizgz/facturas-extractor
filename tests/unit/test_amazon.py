"""Pruebas unitarias para el parser de facturas de Amazon."""

from __future__ import annotations

from gestor_facturas.providers.amazon import AmazonParser


def test_puede_procesar_amazon():
    """Verifica que el parser detecta correctamente las facturas de Amazon."""
    parser = AmazonParser()

    texto_valido = "Factura de Amazon Business EU S.à.r.l..."
    texto_invalido = "Factura de Leroy Merlin España S.A.U."

    assert parser.puede_procesar(texto_valido) is True
    assert parser.puede_procesar(texto_invalido) is False


def test_extraer_datos_amazon_ejemplo():
    """Simula la extracción de datos de una factura de Amazon con IVA al 21%."""
    parser = AmazonParser()

    # Texto simulado similar al contenido real de una factura de Amazon Business
    texto_simulado = """
    amazon business
    Factura
    Vendido por Amazon Business EU S.à.r.l, Sucursal en España
    IVA ESW0264006H
    Fecha de la factura/Fecha de la entrega 23 julio 2026
    Número de la factura ES61DXYAABEI
    Bruguer Masilla Grietas Profundas 1 4,95 € 21.0% 5,99 €
    Total de la factura 5,99 €
    21.0% 4,95 € 1,04 €
    Total 4,95 € 1,04 €
    """

    # Ejecutamos la extracción pasando un path ficticio
    datos = parser.extraer_datos(texto_simulado, "factura_test.pdf")

    # Comprobamos que los campos coinciden exactamente con la estructura esperada
    assert datos["numero_factura"] == "ES61DXYAABEI"
    assert datos["nif_cif"] == "ESW0264006H"
    assert datos["importe_base"] == 4.95
    assert datos["tasa_iva"] == 21.0
    assert datos["importe_iva"] == 1.04
    assert datos["importe_total"] == 5.99

    print("\n--- DATOS EXTRAÍDOS ---")
    for k, v in datos.items():
        print(f"{k}: {v}")


def test_extraer_datos_amazon_nota_credito():
    """Simula la extracción de datos de una nota de crédito de Amazon con importes negativos."""
    parser = AmazonParser()

    # Texto simulado basado en tus notas de crédito reales
    texto_simulado = """
    Nota de crédito
    Productos enviados desde: Alemania
    IVA % Precio total (IVA excluido) IVA
    0% -8,39 € 0,00 €
    Total -8,39 € 0,00 €
    Total pendiente -8,39 €
    Número de la nota de crédito DE60000LRIYTPC
    Número de la factura original DE600069RIYTPI
    Fecha de la nota de crédito 26 julio 2026
    Vendido por guangzhouchensaidianzishangwuyouxiangongsi
    IVA DE450979112
    """

    datos = parser.extraer_datos(texto_simulado, "nota_credito_test.pdf")

    assert datos["numero_factura"] == "DE60000LRIYTPC"
    assert datos["nif_cif"] == "DE450979112"
    assert datos["importe_base"] == -8.39
    assert datos["tasa_iva"] == 0.0
    assert datos["importe_iva"] == 0.0
    assert datos["importe_total"] == -8.39

    print("\n--- DATOS EXTRAÍDOS ---")
    for k, v in datos.items():
        print(f"{k}: {v}")
