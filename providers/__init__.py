# -*- coding: utf-8

from .generic import GenericParser
from .itv import ItvParser

# Orden de detección: específicos primero, genérico al final
PROVIDERS = [
    ItvParser(),
    GenericParser(),
]
