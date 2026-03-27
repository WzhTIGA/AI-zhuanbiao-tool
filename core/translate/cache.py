from __future__ import annotations

from collections import OrderedDict


class TranslationCache:
    def __init__(self, *, max_size: int = 2048) -> None:
        self._max_size = max_size
        self._data: OrderedDict[tuple[str, str, str], str] = OrderedDict()

    def get(self, key: tuple[str, str, str]) -> str | None:
        v = self._data.get(key)
        if v is None:
            return None
        self._data.move_to_end(key)
        return v

    def set(self, key: tuple[str, str, str], value: str) -> None:
        self._data[key] = value
        self._data.move_to_end(key)
        while len(self._data) > self._max_size:
            self._data.popitem(last=False)

    def stats(self) -> dict[str, int]:
        return {"size": len(self._data), "max_size": self._max_size}

