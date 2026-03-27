from __future__ import annotations

import re
import time
from typing import Any

import requests

from .base import Translator
from .cache import TranslationCache


_RE_HAS_CJK = re.compile(r"[\u4e00-\u9fff]")
_RE_NUMERIC = re.compile(r"^[\s\d\.\-_/]+$")


def _normalize_text(text: str) -> str:
    return " ".join(text.strip().split())


def _should_skip(text: str) -> bool:
    if not text:
        return True
    if _RE_NUMERIC.match(text):
        return True
    if _RE_HAS_CJK.search(text):
        return True
    return False


class _RateLimiter:
    def __init__(self, *, min_interval_s: float) -> None:
        self._min_interval_s = min_interval_s
        self._next_allowed = 0.0

    def wait(self) -> None:
        now = time.monotonic()
        if now < self._next_allowed:
            time.sleep(self._next_allowed - now)
        self._next_allowed = time.monotonic() + self._min_interval_s


class MyMemoryTranslator(Translator):
    def __init__(
        self,
        *,
        cache: TranslationCache | None = None,
        min_interval_s: float = 0.12,
        timeout_s: float = 10.0,
        max_retries: int = 2,
    ) -> None:
        self._cache = cache or TranslationCache()
        self._limiter = _RateLimiter(min_interval_s=min_interval_s)
        self._timeout_s = timeout_s
        self._max_retries = max_retries
        self._request_count = 0
        self._cache_hit_count = 0
        self._error_count = 0

    def stats(self) -> dict[str, Any]:
        return {
            "provider": "MyMemory",
            "requests": self._request_count,
            "cache_hits": self._cache_hit_count,
            "errors": self._error_count,
            "cache": self._cache.stats(),
        }

    def translate(self, *, text: str, src_lang: str, dst_lang: str) -> str:
        raw = text
        text = _normalize_text(text)
        if _should_skip(text):
            return raw.strip()

        key = (src_lang, dst_lang, text)
        cached = self._cache.get(key)
        if cached is not None:
            self._cache_hit_count += 1
            return cached

        params = {"q": text, "langpair": f"{src_lang}|{dst_lang}"}
        url = "https://api.mymemory.translated.net/get"

        last_error: Exception | None = None
        for attempt in range(self._max_retries + 1):
            try:
                self._limiter.wait()
                self._request_count += 1
                resp = requests.get(url, params=params, timeout=self._timeout_s)
                resp.raise_for_status()
                payload = resp.json()
                translated = str(payload.get("responseData", {}).get("translatedText", "")).strip()
                if translated:
                    self._cache.set(key, translated)
                    return translated
                self._error_count += 1
                break
            except Exception as e:
                last_error = e
                self._error_count += 1
                if attempt < self._max_retries:
                    time.sleep(0.6 * (2**attempt))
                else:
                    break

        if last_error is not None:
            return raw.strip()
        return raw.strip()
