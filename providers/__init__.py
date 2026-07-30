#!/usr/bin/env python3
# __init__.py

from .amazon import AmazonParser
from .generic import GenericParser
from .itv import ItvParser
from .leroymerlin import LeroyMerlinParser
from .salvadorescoda import SalvadorEscodaParser

# Orden de detección: específicos primero, genérico al final
PROVIDERS = [
    ItvParser(),
    SalvadorEscodaParser(),
    AmazonParser(),
    LeroyMerlinParser(),
    GenericParser(),
]
