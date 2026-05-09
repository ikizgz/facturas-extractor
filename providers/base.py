# -*- coding: utf-8 -*-

from __future__ import annotations

from typing import List

from .common import Row


class ProviderParser:
    name: str = "GENERIC"

    def detect(self, text: str) -> bool:
        return False

    def parse(self, text: str, path) -> List[Row]:
        raise NotImplementedError
