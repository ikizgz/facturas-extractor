"""Parser específico para la extracción de datos en facturas y notas rectificativas de Salvador Escoda S.A."""

from __future__ import annotations

import re
from pathlib import Path

from gestor_facturas.providers.base import ProveedorBase
from gestor_facturas.services.common import _normalizar_numero, _parsear_fecha


class SalvadorEscodaParser(ProveedorBase):
    """Parser para procesar documentos de Salvador Escoda S.A."""

    def puede_procesar(self, raw_text: str) -> bool:
        """Determina si el texto pertenece a una factura o abono de Salvador Escoda."""
        t = raw_text.upper()
        return any(
            x in t for x in ["SALVADOR ESCODA", "ESCODA", "A08710006", "A-08710006"]
        )

    def extraer_datos(self, raw_text: str, pdf_path: str | Path) -> dict:
        """Extrae los campos normalizados de una factura o abono de Salvador Escoda."""
        texto_unificado = " ".join(raw_text.split())
        # Impresión de raw_test para depuración si es necesario
        # print("\nDEBUG: Texto unificado extraído del PDF:\n", texto_unificado)

        # 0. Detección y trazabilidad de abonos / facturas rectificativas
        es_rectificativa = (
            "rectificativa" in texto_unificado.lower()
            or "abono" in texto_unificado.lower()
        )

        # Buscamos el número de factura original mencionado en las líneas de abono (ej: "Abono por devolución Factura: 10092411")
        orig_match = re.search(
            r"factura[:\s]*(\d{6,10})", texto_unificado, re.IGNORECASE
        )

        # 1. Número de factura (o factura rectificativa)
        mnum = re.search(
            r"(?:FACTURA\s+RECTIFICATIVA|FACTURA)\D*(\d{6,10})",
            texto_unificado,
            re.IGNORECASE,
        )
        numero_factura = mnum.group(1) if mnum else Path(pdf_path).stem

        # Construcción de las notas para la trazabilidad
        notas = ""
        if es_rectificativa:
            num_orig = (
                orig_match.group(1)
                if orig_match and orig_match.group(1) != numero_factura
                else ""
            )
            if num_orig:
                notas = f"Factura Abono / Rectificativa - Factura original: {num_orig}"
            else:
                notas = "Factura Abono / Rectificativa"

        # 2. Fecha de la factura
        fecha_match = re.search(
            r"Fecha\s+Factura[:\s]*(\d{1,2}/\d{1,2}/\d{4})",
            texto_unificado,
            re.IGNORECASE,
        )
        if fecha_match:
            fecha_str = fecha_match.group(1)
            # Convertimos formato DD/MM/YYYY a YYYY-MM-DD
            partes = fecha_str.split("/")
            if len(partes) == 3:
                fecha = f"{partes[2]}-{partes[1]}-{partes[0]}"
            else:
                fecha = fecha_str
        else:
            fecha = _parsear_fecha(texto_unificado)

        # 3. Datos del emisor
        proveedor = "SALVADOR ESCODA S.A."
        nif_cif = "A08710006"

        # 4. Tasa de IVA: Extracción directa y limpia del texto
        tasa_iva = 21.0  # Valor por defecto seguro
        iva_pct_match = re.search(
            r"([0-9]+(?:[,.][0-9]{1,2})?)\s*%\s*IVA",
            texto_unificado,
            re.IGNORECASE,
        )
        if iva_pct_match:
            val_iva_str = iva_pct_match.group(1).replace(",", ".")
            try:
                tasa_iva = float(val_iva_str)
            except ValueError:
                pass

        # 5. Importes basados en la estructura real de columnas de Salvador Escoda
        importe_base = 0.0
        importe_iva = 0.0
        importe_total = 0.0

        total_anchor_match = re.search(r"TOTAL\b(.*)", texto_unificado, re.IGNORECASE)

        if total_anchor_match:
            segmento_post_total = total_anchor_match.group(1)

            # 1. El importe total es siempre el que va detrás de 'EUR'
            eur_match = re.search(
                r"EUR\s*(-?\d{1,3}(?:[.,]\d{3})*[.,]\d{2})",
                segmento_post_total,
                re.IGNORECASE,
            )
            if eur_match:
                importe_total = _normalizar_numero(eur_match.group(1))

            # 2. Recogemos todos los números previos a EUR
            parte_antes_eur = (
                segmento_post_total.split("EUR")[0]
                if "EUR" in segmento_post_total
                else segmento_post_total
            )
            nums_encontrados = re.findall(
                r"-?\b\d{1,3}(?:[.,]\d{3})*[.,]\d{2}\b", parte_antes_eur
            )
            nums_limpios = [
                _normalizar_numero(n)
                for n in nums_encontrados
                if _normalizar_numero(n) is not None
            ]

            # 3. Asignación robusta según los números encontrados
            if len(nums_limpios) >= 2:
                # Intentamos primero buscar el binomio que sume el total si hay más de 3 números
                encontrado = False
                if len(nums_limpios) > 3 and importe_total != 0.0:
                    for i in range(len(nums_limpios) - 1):
                        posible_base = nums_limpios[i]
                        posible_iva = nums_limpios[i + 1]
                        if (
                            abs((posible_base + posible_iva) - abs(importe_total))
                            < 0.05
                        ):
                            importe_base = posible_base
                            importe_iva = posible_iva
                            encontrado = True
                            break

                # Si no se encontró por binomio estricto (o hay 3 o menos números como tu caso de 3.20, 3.20, 0.67),
                # los dos últimos de la secuencia antes de EUR son infaliblemente la Base y el IVA.
                if not encontrado:
                    importe_base = nums_limpios[-2]
                    importe_iva = nums_limpios[-1]
            elif len(nums_limpios) == 1:
                importe_base = nums_limpios[0]

        # Respaldo de seguridad si el total no se capturó con EUR
        if importe_total == 0.0 and total_anchor_match:
            nums_totales = re.findall(
                r"-?\b\d{1,3}(?:[.,]\d{3})*[.,]\d{2}\b", total_anchor_match.group(1)
            )
            if nums_totales:
                importe_total = _normalizar_numero(nums_totales[-1])

        # Red de seguridad final por diferencia
        if importe_iva == 0.0 and importe_base != 0.0 and importe_total != 0.0:
            importe_iva = round(importe_total - importe_base, 2)

        # Red de seguridad: Si hay base e IVA, contrastamos con el cálculo real redondeado a enteros
        if importe_iva != 0.0 and importe_base != 0.0:
            tasa_calculada = round(abs(importe_iva / importe_base) * 100.0, 0)
            if tasa_iva != tasa_calculada:
                tasa_iva = tasa_calculada

        return {
            "fecha": fecha,
            "numero_factura": numero_factura,
            "proveedor": proveedor,
            "nif_cif": nif_cif,
            "importe_base": importe_base,
            "tasa_iva": tasa_iva,
            "importe_iva": importe_iva,
            "importe_total": importe_total,
            "notas": notas,
        }
