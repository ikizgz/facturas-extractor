#!/usr/bin/env python3
# base.py

from __future__ import annotations

from .common import Row


class ProviderParser:
    name: str = "GENERIC"

    def detect(self, text: str) -> bool:
        return False

    def parse(self, text: str, path) -> list[Row]:
        raise NotImplementedError
