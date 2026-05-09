# -*- coding: utf-8

from .alcampo import AlcampoParser
from .generic import GenericParser
from .indusan import IndusanParser
from .itv import ItvParser
from .mercadaiz import MercadaizParser
from .o2 import O2Parser
from .repsol import RepsolParser
from .sorpresa import SorpresaParser
from .supercontable import SupercontableParser

# Orden de detección: específicos primero, genérico al final
PROVIDERS = [
    AlcampoParser(),
    IndusanParser(),
    ItvParser(),
    MercadaizParser(),
    O2Parser(),
    RepsolParser(),
    SorpresaParser(),
    SupercontableParser(),
    GenericParser(),
]
