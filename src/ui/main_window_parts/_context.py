# -*- coding: utf-8 -*-
from __future__ import annotations

_SHARED = {}
_NAMESPACES = []
_NAMESPACE_IDS: set[int] = set()
_BASELINE_POS = {}
_PREEXISTING_NAMES = {}

_RESERVED = {
    "_bind_context",
    "_publish_context",
    "bind_context",
    "finalize_context",
    "publish_context",
    "exported_globals",
}


def _is_exportable(name: str) -> bool:
    return not (
        name in _RESERVED
        or name in {"__builtins__", "__cached__", "__doc__", "__file__", "__loader__", "__name__", "__package__", "__spec__"}
    )


def bind_context(namespace) -> None:
    namespace_id = id(namespace)
    _PREEXISTING_NAMES[namespace_id] = set(namespace)
    namespace.update(_SHARED)
    _BASELINE_POS[namespace_id] = len(namespace)
    if namespace_id not in _NAMESPACE_IDS:
        _NAMESPACE_IDS.add(namespace_id)
        _NAMESPACES.append(namespace)


def publish_context(namespace) -> None:
    start = _BASELINE_POS.get(id(namespace), 0)
    preexisting = _PREEXISTING_NAMES.get(id(namespace), set())
    for index, (name, value) in enumerate(namespace.items()):
        if index < start and name not in preexisting:
            continue
        if _is_exportable(name):
            _SHARED[name] = value


def finalize_context() -> None:
    for registered in list(_NAMESPACES):
        registered.update(_SHARED)


def exported_globals() -> dict:
    return dict(_SHARED)
