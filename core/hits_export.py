from __future__ import annotations

import io
import json

import openpyxl

from .excel_search import NameHit, SecondaryHit


def export_hits_xlsx(
    *,
    query_name: str,
    query_type_en: str,
    name_hits: list[NameHit],
    secondary_hits: list[SecondaryHit],
) -> bytes:
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "名称搜索命中"
    ws1.append(["查询名称", str(query_name).strip()])
    ws1.append([])
    ws1.append(["表文件", "工作表", "行号", "id", "key", "B列"])
    for h in name_hits:
        ws1.append([h.file_name, h.sheet_name, h.row, h.id_value or "", h.key_value or "", h.b_value or ""])

    ws2 = wb.create_sheet("同类二次检索命中")
    ws2.append(["查询类型(英文)", str(query_type_en).strip()])
    ws2.append([])
    ws2.append(["表文件", "工作表", "行号", "匹配方式", "源id", "源key", "源B列", "行数据(json)"])
    for h in secondary_hits:
        ws2.append(
            [
                h.file_name,
                h.sheet_name,
                h.row,
                h.matched_by,
                h.source_id or "",
                h.source_key or "",
                h.source_b or "",
                json.dumps(h.row_data, ensure_ascii=False, sort_keys=True),
            ]
        )

    buf = io.BytesIO()
    wb.save(buf)
    wb.close()
    return buf.getvalue()

