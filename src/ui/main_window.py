# -*- coding: utf-8 -*-
"""Public facade for the OneNote Remocon main window.

The implementation is split under ``src.ui.main_window_parts`` so each source
file stays small while preserving the historical ``src.ui.main_window`` import
path used by ``main.py`` and older builds.
"""

from importlib import import_module

from src.ui.main_window_parts._context import (
    exported_globals as _exported_globals,
    finalize_context as _finalize_context,
)

_PART_MODULES = (
    'pre_01',
    'pre_02',
    'pre_03',
    'pre_04',
    'pre_05',
    'pre_06',
    'pre_07',
    'pre_08',
    'pre_09',
    'pre_10',
    'pre_11',
    'pre_12',
    'pre_13',
    'pre_14',
    'pre_15_mac_collect',
    'pre_15_mac_prepare',
    'pre_15_mac_launch',
    'pre_15',
    'pre_16',
    'mixin_01',
    'mixin_02',
    'mixin_03',
    'mixin_04',
    'mixin_05',
    'mixin_06_windows_a',
    'mixin_06_windows_b',
    'mixin_07',
    'mixin_08',
    'mixin_09',
    'mixin_10',
    'mixin_11',
    'mixin_12',
    'mixin_13',
    'mixin_14',
    'mixin_15',
    'mixin_16',
    'mixin_17',
    'mixin_18',
    'mixin_19',
    'mixin_20',
    'mixin_21',
    'mixin_22',
    'mixin_23',
    'mixin_24',
    'mixin_25',
    'mixin_26_left',
    'mixin_26_right',
    'mixin_26',
    'mixin_27',
    'mixin_28',
    'mixin_29',
    'mixin_30',
    'mixin_31',
    'mixin_32',
    'mixin_33',
    'mixin_34',
    'mixin_35',
    'mixin_36',
    'mixin_37',
    'mixin_38',
    'mixin_39',
    'mixin_40',
    'mixin_41',
    'mixin_42',
    'mixin_43',
    'app',
)

for _module_name in _PART_MODULES:
    import_module(f"src.ui.main_window_parts.{_module_name}")

_finalize_context()
globals().update(_exported_globals())

for _name, _value in list(globals().items()):
    if getattr(_value, "__module__", "").startswith("src.ui.main_window_parts."):
        try:
            _value.__module__ = __name__
        except Exception:
            pass

__all__ = [name for name in globals() if not name.startswith("_")]
