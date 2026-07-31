"""Definición de la clase base abstracta para los proveedores de facturas."""

from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path


class ProveedorBase(ABC):
    """Clase base abstracta que deben implementar todos los extractores de facturas

    específicos por proveedor.
    """

    @abstractmethod
    def puede_procesar(self, texto_pdf: str) -> bool:
        """Determina si este proveedor es el encargado de procesar la factura

        analizando patrones clave en el texto extraído.
        """

    @abstractmethod
    def extraer_datos(self, texto_pdf: str, archivo_pdf: str | Path) -> dict:
        """Extrae los campos normalizados de la factura y devuelve un diccionario

        compatible con la estructura del archivo Excel.
        """
