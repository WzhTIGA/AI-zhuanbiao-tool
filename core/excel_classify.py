from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class ClassifiedWorkbook:
    output_filename: str
    original_base: str
    category_en: str | None
    category_cn: str


def classify_by_filename(
    *,
    original_base: str,
    type_dict: dict[str, str],
    uncategorized_cn: str = "未分类",
) -> tuple[str | None, str]:
    name = original_base.casefold()
    best_key: str | None = None
    for k in type_dict.keys():
        kk = str(k).strip()
        if not kk:
            continue
        if kk.casefold() in name:
            if best_key is None or len(kk) > len(best_key):
                best_key = kk
    if best_key is None:
        return None, uncategorized_cn
    return best_key, type_dict.get(best_key, uncategorized_cn)

