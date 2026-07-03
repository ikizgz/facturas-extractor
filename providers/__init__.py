# -*- coding: utf-8

from .generic import GenericParser
from .itv import ItvParser
from .salvadorescoda import SalvadorEscodaParser

# Orden de detección: específicos primero, genérico al final
PROVIDERS = [
    ItvParser(),
    SalvadorEscodaParser(),
    GenericParser(),
]
