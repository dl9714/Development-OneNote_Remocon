# -*- coding: utf-8 -*-
"""Public facade for macOS OneNote UI automation helpers.

The implementation is split under ``src.macos_ui_parts`` to keep individual
Python files small while preserving the historical ``src.macos_ui`` import path.
"""

from importlib import import_module

from src.macos_ui_parts._context import (
    exported_globals as _exported_globals,
    finalize_context as _finalize_context,
)

_PART_MODULES = (
    'part_01',
    'part_02',
    'part_03',
    'part_04',
    'part_05',
    'part_06',
    'part_07',
    'part_08',
    'part_09',
    'part_10',
    'part_11',
    'part_12',
    'part_13',
    'part_14',
    'part_15',
    'part_16',
    'part_17',
)

for _module_name in _PART_MODULES:
    import_module(f"src.macos_ui_parts.{_module_name}")

_finalize_context()
globals().update(_exported_globals())

for _name, _value in list(globals().items()):
    if getattr(_value, "__module__", "").startswith("src.macos_ui_parts."):
        try:
            _value.__module__ = __name__
        except Exception:
            pass

__all__ = [name for name in globals() if not name.startswith("_")]
