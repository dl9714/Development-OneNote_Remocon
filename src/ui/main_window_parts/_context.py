# -*- coding: utf-8 -*-
from __future__ import annotations

from typing import Any, Dict, MutableMapping

_SHARED: Dict[str, Any] = {}
_NAMESPACES: list[MutableMapping[str, Any]] = []

_RESERVED = {
    "_bind_context",
    "_publish_context",
    "bind_context",
    "publish_context",
    "exported_globals",
}


def _is_exportable(name: str) -> bool:
    return not (
        name in _RESERVED
        or name in {"__builtins__", "__cached__", "__doc__", "__file__", "__loader__", "__name__", "__package__", "__spec__"}
    )


def bind_context(namespace: MutableMapping[str, Any]) -> None:
    namespace.update(_SHARED)
    if namespace not in _NAMESPACES:
        _NAMESPACES.append(namespace)


def publish_context(namespace: MutableMapping[str, Any]) -> None:
    for name, value in list(namespace.items()):
        if _is_exportable(name):
            _SHARED[name] = value
    for registered in list(_NAMESPACES):
        registered.update(_SHARED)


def exported_globals() -> Dict[str, Any]:
    return dict(_SHARED)
