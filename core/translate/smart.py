from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

from .base import Translator


_RE_EN_SEGMENT = re.compile(r"[A-Za-z][A-Za-z0-9_./\-]*")
_RE_SPLIT_CAMEL = re.compile(r"[A-Z]+(?=[A-Z][a-z])|[A-Z]?[a-z]+|[A-Z]+|\d+")
_RE_SEP = re.compile(r"[_./\-]+")


def _normalize_spaces(s: str) -> str:
    return " ".join(str(s).strip().split())


def _split_english_token(token: str) -> list[str]:
    parts: list[str] = []
    for chunk in _RE_SEP.split(token):
        chunk = chunk.strip()
        if not chunk:
            continue
        parts.extend([m.group(0) for m in _RE_SPLIT_CAMEL.finditer(chunk)])
    return [p for p in parts if p]


def _translate_piece(token: str, *, dict_map: dict[str, str], inner: Translator, src_lang: str, dst_lang: str) -> str:
    k = token.casefold()
    mapped = dict_map.get(k)
    if mapped:
        return mapped
    return inner.translate(text=token, src_lang=src_lang, dst_lang=dst_lang).strip() or token


@dataclass(frozen=True)
class SmartTranslator(Translator):
    inner: Translator
    dict_map: dict[str, str]

    def translate(self, *, text: str, src_lang: str, dst_lang: str) -> str:
        raw = str(text)
        text = _normalize_spaces(raw)
        if not text:
            return ""

        out: list[str] = []
        last = 0
        for m in _RE_EN_SEGMENT.finditer(text):
            if m.start() > last:
                out.append(text[last : m.start()])
            seg = m.group(0)
            tokens = _split_english_token(seg)
            if not tokens:
                out.append(seg)
            else:
                translated = "".join(
                    _translate_piece(t, dict_map=self.dict_map, inner=self.inner, src_lang=src_lang, dst_lang=dst_lang)
                    for t in tokens
                )
                out.append(translated or seg)
            last = m.end()

        if last < len(text):
            out.append(text[last:])

        return "".join(out).strip()

    def stats(self) -> dict[str, Any] | None:
        s = getattr(self.inner, "stats", None)
        if callable(s):
            return {"smart": {"dict_size": len(self.dict_map)}, "inner": s()}
        return None


def default_smart_dict() -> dict[str, str]:
    return {
        "id": "编号",
        "key": "键",
        "name": "名称",
        "desc": "描述",
        "description": "描述",
        "type": "类型",
        "level": "等级",
        "lvl": "等级",
        "exp": "经验",
        "gold": "金币",
        "cost": "消耗",
        "price": "价格",
        "value": "数值",
        "rate": "概率",
        "chance": "概率",
        "prob": "概率",
        "min": "最小",
        "max": "最大",
        "hp": "生命",
        "mp": "法力",
        "atk": "攻击",
        "attack": "攻击",
        "def": "防御",
        "defense": "防御",
        "crit": "暴击",
        "critical": "暴击",
        "spd": "速度",
        "speed": "速度",
        "time": "时间",
        "duration": "持续时间",
        "cooldown": "冷却",
        "cd": "冷却",
        "buff": "增益",
        "debuff": "减益",
        "skill": "技能",
        "item": "道具",
        "equip": "装备",
        "hero": "英雄",
        "monster": "怪物",
        "npc": "机器人",
        "pet": "宠物",
        "chapter": "章节",
        "task": "任务",
        "drop": "掉落",
        "shop": "商店",
        "store": "商店",
        "language": "语言",
        "activity": "活动",
        "arena": "竞技场",
        "attr": "属性",
        "attribute": "属性",
        "artifact": "神器",
        "comconf": "组合配置",
        "model": "模型",
        "pay": "支付",
    }

