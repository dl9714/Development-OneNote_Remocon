# -*- coding: utf-8 -*-
from __future__ import annotations


class LazyModule:
    __slots__ = ("_name", "_module")

    def __init__(self, name: str):
        self._name = name
        self._module = None

    def __getattr__(self, item: str):
        if self._module is None:
            self._module = __import__(self._name, fromlist=["*"])
        return getattr(self._module, item)


class LazyPath:
    __slots__ = ("_path", "_value")

    def __init__(self, path: str):
        self._path = path
        self._value = None

    def _get(self):
        if self._value is None:
            from pathlib import Path

            self._value = Path(self._path)
        return self._value

    def __getattr__(self, item: str):
        return getattr(self._get(), item)

    def __fspath__(self):
        return self._path

    def __str__(self):
        return self._path
