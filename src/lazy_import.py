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


class LazyAttr:
    __slots__ = ("_module", "_name")

    def __init__(self, module_name: str, name: str):
        self._module = LazyModule(module_name)
        self._name = name

    def _get(self):
        return getattr(self._module, self._name)

    def __call__(self, *args, **kwargs):
        return self._get()(*args, **kwargs)

    def __getattr__(self, item: str):
        return getattr(self._get(), item)


class _LazyClassMeta(type):
    def _get(cls):
        return getattr(cls._lazy_module, cls._lazy_name)

    def __call__(cls, *args, **kwargs):
        return cls._get()(*args, **kwargs)

    def __instancecheck__(cls, instance):
        return isinstance(instance, cls._get())

    def __subclasscheck__(cls, subclass):
        return issubclass(subclass, cls._get())

    def __getattr__(cls, item: str):
        return getattr(cls._get(), item)


def lazy_class(module_name: str, name: str, base=object):
    return _LazyClassMeta(
        name,
        (base,),
        {
            "__module__": module_name,
            "_lazy_module": LazyModule(module_name),
            "_lazy_name": name,
        },
    )


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
