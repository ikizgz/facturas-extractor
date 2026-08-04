"""Paquete de parsers específicos por proveedor.

Centraliza el registro de todos los proveedores soportados por la aplicación.
"""

from __future__ import annotations

from gestor_facturas.providers.amazon import AmazonParser
from gestor_facturas.providers.base import ProveedorBase

# Lista oficial de parsers registrados en la aplicación
# A medida que crees nuevos proveedores (ej. LeroyMerlinParser, LuzParser),
# solo tienes que importarlos y añadirlos a esta lista.
PROVEEDORES_REGISTRADOS: list[ProveedorBase] = [
    AmazonParser(),
]
