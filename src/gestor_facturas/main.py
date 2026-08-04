"""Script principal (Entrypoint) para la ejecución del gestor de facturas."""

from __future__ import annotations

import logging
from pathlib import Path

from gestor_facturas.core.config import (
    ARCHIVO_EXCEL_SALIDA,
    CARPETA_LOGS,
)
from gestor_facturas.core.pipeline import procesar_pdf
from gestor_facturas.services.excel import guardar_factura_excel
from gestor_facturas.services.logging_utils import setup_logging

logger = logging.getLogger(__name__)


def ejecutar_lote(carpeta_facturas: str | Path = "facturas") -> None:
    """Procesa todos los archivos PDF encontrados en la carpeta especificada."""
    # Inicializar el sistema de logs centralizado (consola + archivo con timestamp)
    setup_logging(
        app_name="gestor_facturas",
        target_dir=Path.cwd(),
        level=logging.INFO,
    )

    dir_path = Path(carpeta_facturas)

    if not dir_path.exists() or not dir_path.is_dir():
        logger.error(
            "La carpeta de facturas '%s' no existe o no es un directorio.",
            dir_path.resolve(),
        )
        print(
            f"\n[ERROR] Por favor, crea una carpeta llamada '{carpeta_facturas}' y coloca dentro tus PDFs de facturas,\n o pasa el nombre de la carpeta donde las tienes."
        )
        return

    archivos_pdf = list(dir_path.glob("*.pdf"))

    if not archivos_pdf:
        logger.warning(
            "No se han encontrado archivos PDF en la carpeta: %s",
            dir_path.resolve(),
        )
        print(
            f"\n[AVISO] No hay archivos PDF en la carpeta '{carpeta_facturas}'.\n"
        )
        return

    logger.info("Se han encontrado %d archivos PDF para procesar.", len(archivos_pdf))
    print(f"\n--- Iniciando procesamiento de {len(archivos_pdf)} facturas ---\n")

    exitosos = 0
    fallidos = 0

    for pdf_path in archivos_pdf:
        print(f"Procesando: {pdf_path.name}...")
        datos_factura = procesar_pdf(pdf_path)

        if datos_factura:
            if guardar_factura_excel(datos_factura, ARCHIVO_EXCEL_SALIDA):
                exitosos += 1
                print(f"  -> OK: Guardado en {ARCHIVO_EXCEL_SALIDA}")
            else:
                fallidos += 1
                print("  -> ERROR: No se pudo guardar en el Excel.")
        else:
            fallidos += 1
            print(
                "  -> ADVERTENCIA: No se pudo extraer información válida o procesar el proveedor."
            )

    print("\n--- Proceso finalizado ---")
    print(f"Procesadas con éxito: {exitosos}")
    print(f"Con errores o no reconocidas: {fallidos}")
    print(f"Resultados guardados en: {ARCHIVO_EXCEL_SALIDA}\n")


if __name__ == "__main__":
    ejecutar_lote("facturas")
