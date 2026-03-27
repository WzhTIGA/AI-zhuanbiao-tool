from __future__ import annotations

from dataclasses import dataclass
import inspect
from typing import Any, Callable

from .excel_classify import classify_by_filename
from .excel_search import (
    MatchMode,
    NameHit,
    SecondaryHit,
    count_sheets_in_xlsx_bytes,
    find_name_hits,
    find_secondary_hits,
)
from .excel_transform import ProcessedWorkbook, filename_base, process_workbook
from .hits_export import export_hits_xlsx
from .translate.base import Translator
from .translate.mymemory import MyMemoryTranslator
from .translate.smart import SmartTranslator, default_smart_dict
from .zip_io import read_xlsx_files_from_archive, write_zip


@dataclass(frozen=True)
class PipelineOutput:
    summary: dict[str, Any]
    processed_zip_bytes: bytes


@dataclass(frozen=True)
class SearchOutput:
    summary: dict[str, Any]
    name_hits: list[NameHit]
    secondary_hits: list[SecondaryHit]


def default_type_dict() -> dict[str, str]:
    return {
        "item": "道具",
        "equip": "装备",
        "hero": "英雄",
        "monster": "怪物",
        "npc": "机器人",
        "pet": "宠物",
        "skill": "技能",
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


def run_process(
    *,
    zip_bytes: bytes,
    type_dict: dict[str, str],
    translator: Translator | None = None,
    progress_cb: Callable[[float, str], None] | None = None,
) -> PipelineOutput:
    last_percent = -1.0

    def report(percent: float, message: str) -> None:
        nonlocal last_percent
        if progress_cb is None:
            return
        p = float(percent)
        if p < 0:
            p = 0.0
        if p > 100:
            p = 100.0
        if p < last_percent:
            p = last_percent
        last_percent = p
        progress_cb(p, message)

    report(0, "读取压缩包…")
    extracted = read_xlsx_files_from_archive(zip_bytes)
    if not extracted:
        raise ValueError("压缩包中未找到任何 xlsx 文件。")
    if translator is None:
        translator = SmartTranslator(inner=MyMemoryTranslator(), dict_map=default_smart_dict())
    report(2, f"发现 {len(extracted)} 个 xlsx，开始处理…")

    processed: list[ProcessedWorkbook] = []
    processed_files: dict[str, bytes] = {}
    rename_rows: list[dict[str, str]] = []

    total_files = len(extracted)
    for i, (original_filename, xlsx_bytes) in enumerate(extracted.items(), start=1):
        report(2 + (i - 1) / max(total_files, 1) * 88, f"处理文件 {i}/{total_files}：{original_filename}")
        p = process_workbook(original_filename=original_filename, xlsx_bytes=xlsx_bytes, translator=translator)
        processed.append(p)
        
        _, category_cn = classify_by_filename(original_base=p.original_base, type_dict=type_dict)
        category_folder = category_cn if category_cn else "未分类"
        out_path = f"{category_folder}/{p.output_filename}"
        
        processed_files[out_path] = p.output_bytes
        rename_rows.append({"原文件": original_filename, "新文件": out_path})
        report(2 + i / max(total_files, 1) * 88, f"已处理 {i}/{total_files}：{p.output_filename}")

    report(92, "分类统计…")
    counts: dict[str, int] = {}
    for p in processed:
        _, category_cn = classify_by_filename(original_base=p.original_base, type_dict=type_dict)
        category_folder = category_cn if category_cn else "未分类"
        counts[category_folder] = counts.get(category_folder, 0) + 1

    report(98, "打包结果…")
    processed_zip_bytes = write_zip(processed_files)
    translator_stats = getattr(translator, "stats", None)
    summary = {
        "第2步_重命名": rename_rows,
        "第5步_分类统计": counts,
        "翻译统计": translator_stats() if callable(translator_stats) else None,
        "处理文件数": len(processed_files),
    }

    report(100, "完成")
    return PipelineOutput(summary=summary, processed_zip_bytes=processed_zip_bytes)


def run_search(
    *,
    zip_bytes: bytes,
    query_name: str,
    match_mode: MatchMode,
    query_type_en: str,
    type_dict: dict[str, str],
    progress_cb: Callable[..., None] | None = None,
) -> SearchOutput:
    last_percent = -1.0
    wants_sheet = False
    if progress_cb is not None:
        try:
            wants_sheet = len(inspect.signature(progress_cb).parameters) >= 4
        except Exception:
            wants_sheet = False

    def report(percent: float, message: str, *, sheet_index: int = 0, sheet_total: int = 0) -> None:
        nonlocal last_percent
        if progress_cb is None:
            return
        p = float(percent)
        if p < 0:
            p = 0.0
        if p > 100:
            p = 100.0
        if p < last_percent:
            p = last_percent
        last_percent = p
        if wants_sheet:
            progress_cb(p, message, int(sheet_index), int(sheet_total))
        else:
            progress_cb(p, message)

    report(0, "读取压缩包…")
    extracted = read_xlsx_files_from_archive(zip_bytes)
    if not extracted:
        raise ValueError("压缩包中未找到任何 xlsx 文件。")

    total_files = len(extracted)
    total_sheets = sum(count_sheets_in_xlsx_bytes(b) for b in extracted.values())
    report(2, f"发现 {total_files} 个 xlsx，共 {total_sheets} 张表，开始搜索…", sheet_index=0, sheet_total=total_sheets)

    def on_name_sheet(cur: int, total: int, file_name: str, sheet_name: str) -> None:
        report(
            15 + (cur / max(total, 1)) * 50,
            f"名称搜索… {file_name} / {sheet_name}",
            sheet_index=cur,
            sheet_total=total,
        )

    report(15, "名称搜索…", sheet_index=0, sheet_total=total_sheets)
    name_hits = find_name_hits(
        workbooks=extracted,
        query_name=query_name,
        match_mode=match_mode,
        progress_cb=on_name_sheet,
        sheet_total=total_sheets,
    )
    report(65, f"名称搜索完成：{'命中' if name_hits else '未命中'}")

    report(70, "分类统计…")
    category_by_file: dict[str, dict[str, str | None]] = {}
    counts: dict[str, int] = {}
    for fn in extracted.keys():
        base = filename_base(fn)
        category_en, category_cn = classify_by_filename(original_base=base, type_dict=type_dict)
        category_by_file[fn] = {"类型英文": category_en, "类型中文": category_cn}
        category_folder = category_cn if category_cn else "未分类"
        counts[category_folder] = counts.get(category_folder, 0) + 1

    step4: dict[str, Any]
    source: NameHit | None = None
    if not name_hits:
        step4 = {"状态": "未命中", "提示": "未找到对应名称物品"}
    else:
        source = name_hits[0]
        step4 = {
            "状态": "已命中",
            "命中数量": len(name_hits),
            "首条命中": {
                "表文件": source.file_name,
                "工作表": source.sheet_name,
                "行号": source.row,
                "id": source.id_value,
                "key": source.key_value,
                "B列": source.b_value,
            },
        }

    query_type_en = query_type_en.strip()
    step6: dict[str, Any]
    secondary_hits: list[SecondaryHit] = []
    if not query_type_en:
        step6 = {"状态": "未执行", "提示": "未填写类型"}
    else:
        same_type_files = {
            fn
            for fn, meta in category_by_file.items()
            if (meta.get("类型英文") or "").casefold() == query_type_en.casefold()
        }
        if not same_type_files:
            step6 = {"状态": "未命中", "提示": "未找到相同类的物品", "类型英文": query_type_en}
        elif source is None:
            step6 = {"状态": "未执行", "提示": "未找到对应名称物品，无法进行二次检索", "类型英文": query_type_en}
        else:
            secondary_sheets = sum(count_sheets_in_xlsx_bytes(extracted[fn]) for fn in same_type_files if fn in extracted)

            def on_secondary_sheet(cur: int, total: int, file_name: str, sheet_name: str) -> None:
                report(
                    80 + (cur / max(total, 1)) * 15,
                    f"同类二次检索… {file_name} / {sheet_name}",
                    sheet_index=cur,
                    sheet_total=total,
                )

            report(80, "同类二次检索…", sheet_index=0, sheet_total=secondary_sheets)
            secondary_hits = find_secondary_hits(
                workbooks=extracted,
                source=source,
                only_files=set(same_type_files),
                progress_cb=on_secondary_sheet,
                sheet_total=secondary_sheets,
            )
            if not secondary_hits:
                step6 = {
                    "状态": "未命中",
                    "提示": "未找到同名称同类型的物品",
                    "类型英文": query_type_en,
                    "源数据": {"id": source.id_value, "key": source.key_value, "B列": source.b_value},
                }
            else:
                step6 = {
                    "状态": "已命中",
                    "命中数量": len(secondary_hits),
                    "类型英文": query_type_en,
                    "源数据": {"id": source.id_value, "key": source.key_value, "B列": source.b_value},
                }
            report(95, f"二次检索完成：{'命中' if secondary_hits else '未命中'}")

    summary = {
        "第5步_分类统计": counts,
        "第4步_名称搜索": step4,
        "第6步_同类二次检索": step6,
        "搜索文件数": len(extracted),
    }
    report(100, "完成")
    return SearchOutput(summary=summary, name_hits=name_hits, secondary_hits=secondary_hits)


def run_pipeline(
    *,
    zip_bytes: bytes,
    query_type_en: str,
    query_name: str,
    match_mode: MatchMode,
    type_dict: dict[str, str],
    translator: Translator | None = None,
    progress_cb: Callable[[float, str], None] | None = None,
) -> PipelineOutput:
    last_percent = -1.0

    def report(percent: float, message: str) -> None:
        nonlocal last_percent
        if progress_cb is None:
            return
        p = float(percent)
        if p < 0:
            p = 0.0
        if p > 100:
            p = 100.0
        if p < last_percent:
            p = last_percent
        last_percent = p
        progress_cb(p, message)

    report(0, "读取压缩包…")
    extracted = read_xlsx_files_from_archive(zip_bytes)
    if not extracted:
        raise ValueError("压缩包中未找到任何 xlsx 文件。")
    if translator is None:
        translator = SmartTranslator(inner=MyMemoryTranslator(), dict_map=default_smart_dict())
    report(2, f"发现 {len(extracted)} 个 xlsx，开始处理…")

    processed: list[ProcessedWorkbook] = []
    processed_files: dict[str, bytes] = {}
    rename_rows: list[dict[str, str]] = []

    total_files = len(extracted)
    for i, (original_filename, xlsx_bytes) in enumerate(extracted.items(), start=1):
        report(2 + (i - 1) / max(total_files, 1) * 78, f"处理文件 {i}/{total_files}：{original_filename}")
        p = process_workbook(original_filename=original_filename, xlsx_bytes=xlsx_bytes, translator=translator)
        processed.append(p)
        processed_files[p.output_filename] = p.output_bytes
        rename_rows.append({"原文件": original_filename, "新文件": p.output_filename})
        report(2 + i / max(total_files, 1) * 78, f"已处理 {i}/{total_files}：{p.output_filename}")

    report(82, "分类统计…")
    category_by_file: dict[str, dict[str, str | None]] = {}
    counts: dict[str, int] = {}
    for p in processed:
        category_en, category_cn = classify_by_filename(original_base=p.original_base, type_dict=type_dict)
        category_by_file[p.output_filename] = {"类型英文": category_en, "类型中文": category_cn}
        counts[category_cn] = counts.get(category_cn, 0) + 1

    report(86, "名称搜索…")
    hits = find_name_hits(
        workbooks=processed_files,
        query_name=query_name,
        match_mode=match_mode,
    )

    step4: dict[str, Any]
    source: Any = None
    if not hits:
        step4 = {"状态": "未命中", "提示": "未找到对应名称物品"}
    else:
        source = hits[0]
        step4 = {
            "状态": "已命中",
            "命中数量": len(hits),
            "首条命中": {
                "表文件": source.file_name,
                "工作表": source.sheet_name,
                "行号": source.row,
                "id": source.id_value,
                "key": source.key_value,
                "B列": source.b_value,
            },
        }
    report(90, f"名称搜索完成：{'命中' if hits else '未命中'}")

    query_type_en = query_type_en.strip()
    step6: dict[str, Any]
    secondary_hits = []
    if not query_type_en:
        step6 = {"状态": "未执行", "提示": "未填写类型"}
    else:
        same_type_files = {
            fn
            for fn, meta in category_by_file.items()
            if (meta.get("类型英文") or "").casefold() == query_type_en.casefold()
        }
        if not same_type_files:
            step6 = {"状态": "未命中", "提示": "未找到相同类的物品", "类型英文": query_type_en}
        elif source is None:
            step6 = {"状态": "未执行", "提示": "未找到对应名称物品，无法进行二次检索", "类型英文": query_type_en}
        else:
            only_files = set(same_type_files)
            report(92, "同类二次检索…")
            secondary_hits = find_secondary_hits(workbooks=processed_files, source=source, only_files=only_files)
            if not secondary_hits:
                step6 = {
                    "状态": "未命中",
                    "提示": "未找到同名称同类型的物品",
                    "类型英文": query_type_en,
                    "源数据": {"id": source.id_value, "key": source.key_value, "B列": source.b_value},
                }
            else:
                step6 = {
                    "状态": "已命中",
                    "命中数量": len(secondary_hits),
                    "类型英文": query_type_en,
                    "源数据": {"id": source.id_value, "key": source.key_value, "B列": source.b_value},
                    "命中列表": [
                        {
                            "表文件": h.file_name,
                            "工作表": h.sheet_name,
                            "行号": h.row,
                            "匹配方式": h.matched_by,
                            "行数据": h.row_data,
                        }
                        for h in secondary_hits
                    ],
                }
            report(95, f"二次检索完成：{'命中' if secondary_hits else '未命中'}")

    report(97, "生成命中数据文件…")
    hits_xlsx = export_hits_xlsx(
        query_name=query_name,
        query_type_en=query_type_en,
        name_hits=hits,
        secondary_hits=secondary_hits,
    )
    hits_name = "命中数据.xlsx"
    if hits_name in processed_files:
        for i in range(2, 2000):
            cand = f"命中数据({i}).xlsx"
            if cand not in processed_files:
                hits_name = cand
                break
    processed_files[hits_name] = hits_xlsx

    report(99, "打包结果…")
    processed_zip_bytes = write_zip(processed_files)
    translator_stats = getattr(translator, "stats", None)
    summary = {
        "第2步_重命名": rename_rows,
        "第5步_分类统计": counts,
        "第4步_名称搜索": step4,
        "第6步_同类二次检索": step6,
        "命中数据文件": hits_name,
        "翻译统计": translator_stats() if callable(translator_stats) else None,
        "处理文件数": len(processed_files),
    }

    report(100, "完成")
    return PipelineOutput(summary=summary, processed_zip_bytes=processed_zip_bytes)
