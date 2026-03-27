from __future__ import annotations

from abc import ABC, abstractmethod


class Translator(ABC):
    @abstractmethod
    def translate(self, *, text: str, src_lang: str, dst_lang: str) -> str:
        raise NotImplementedError()

