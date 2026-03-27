from __future__ import annotations

import os
import json

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))  # 把项目根目录加入 sys.path

import streamlit as st

from core.pipeline import MatchMode, default_type_dict, run_process, run_search

DICT_FILE = "type_dict.json"

def load_type_dict() -> dict[str, str]:
    if os.path.exists(DICT_FILE):
        try:
            with open(DICT_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return default_type_dict()

def save_type_dict(d: dict[str, str]) -> None:
    with open(DICT_FILE, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=2)

def _download_button(*, label: str, data: bytes, file_name: str) -> None:
    st.download_button(label=label, data=data, file_name=file_name, mime="application/zip")

def main() -> None:
    st.set_page_config(page_title="Excel 表转换与道具检索工具", layout="wide")
    st.title("Excel 表转换与道具检索工具")

    # Load type dict to session state if not present
    if "type_dict" not in st.session_state:
        st.session_state["type_dict"] = load_type_dict()

    tab1, tab2, tab3 = st.tabs(["物品搜索", "表格转换", "类型词典"])

    with tab1:
        st.subheader("物品搜索")
        
        col1, col2, col3 = st.columns([2, 2, 2])
        with col1:
            query_name = st.text_input("物品名称", value="", key="search_query_name")
        with col2:
            match_mode: MatchMode = st.radio("名称匹配方式", options=["精准匹配", "包含匹配"], horizontal=True, key="search_match_mode")
        with col3:
            query_type_en = st.text_input("物品类型（英文）", value="item", key="search_query_type_en")

        uploaded_zip_search = st.file_uploader("上传需要搜索的 zip/rar 压缩包", type=["zip", "rar"], key="search_zip")

        search_disabled = uploaded_zip_search is None or not query_name.strip()
        do_search = st.button("搜索", disabled=search_disabled, key="search_btn")
        search_msg = st.empty()
        search_sheet_ph = st.empty()
        search_progress_ph = st.empty()

        if do_search:
            if uploaded_zip_search is None:
                search_msg.error("请先上传 zip/rar 压缩包。")
            else:
                search_msg.empty()
                search_sheet_ph.empty()
                progress = search_progress_ph.progress(0, text="准备开始…")

                def on_progress_search(percent: float, message: str, current_sheet: int = 0, total_sheets: int = 0) -> None:
                    if total_sheets > 0:
                        search_sheet_ph.markdown(f"当前第 {current_sheet} 张表 / 总共 {total_sheets} 张表")
                    else:
                        search_sheet_ph.empty()
                    progress.progress(int(round(percent)), text=message)

                try:
                    out = run_search(
                        zip_bytes=uploaded_zip_search.getvalue(),
                        query_name=query_name.strip(),
                        match_mode=match_mode,
                        query_type_en=query_type_en.strip(),
                        type_dict=st.session_state["type_dict"],
                        progress_cb=on_progress_search,
                    )
                    st.session_state["search_out"] = out
                    progress.progress(100, text="搜索完成")
                except Exception as e:
                    progress.progress(0, text="搜索失败")
                    search_msg.error(f"搜索失败：{e}")

        search_out = st.session_state.get("search_out")
        if search_out is not None:
            st.markdown("### 搜索结果")

            name_rows = [
                {
                    "表文件": h.file_name,
                    "工作表": h.sheet_name,
                    "行号": h.row,
                    "id": h.id_value,
                    "key": h.key_value,
                    "B列": h.b_value,
                }
                for h in search_out.name_hits
            ]
            st.markdown(f"**首条命中（名称命中）：{len(name_rows)} 条**")
            st.dataframe(name_rows, use_container_width=True, hide_index=True)

            secondary_rows = [
                {
                    "表文件": h.file_name,
                    "工作表": h.sheet_name,
                    "行号": h.row,
                    "匹配方式": h.matched_by,
                    "源id": h.source_id,
                    "源key": h.source_key,
                    "源B列": h.source_b,
                    "行数据": json.dumps(h.row_data, ensure_ascii=False),
                }
                for h in search_out.secondary_hits
            ]
            st.markdown(f"**二次检索命中（同类命中）：{len(secondary_rows)} 条**")
            st.dataframe(secondary_rows, use_container_width=True, hide_index=True)

    with tab2:
        st.subheader("表格转换")
        uploaded_zip_process = st.file_uploader("上传需要处理的 zip/rar 压缩包", type=["zip", "rar"], key="process_zip")
        
        process_disabled = uploaded_zip_process is None
        do_process = st.button("开始处理", type="primary", disabled=process_disabled, key="process_btn")
        process_msg = st.empty()
        process_progress_ph = st.empty()

        if do_process:
            if uploaded_zip_process is None:
                process_msg.error("请先上传 zip/rar 压缩包。")
            else:
                process_msg.empty()
                progress = process_progress_ph.progress(0, text="准备开始…")

                def on_progress_process(percent: float, message: str) -> None:
                    progress.progress(int(round(percent)), text=message)

                try:
                    out = run_process(
                        zip_bytes=uploaded_zip_process.getvalue(),
                        type_dict=st.session_state["type_dict"],
                        progress_cb=on_progress_process,
                    )
                    st.session_state["process_out"] = out
                    progress.progress(100, text="处理完成")
                except Exception as e:
                    progress.progress(0, text="处理失败")
                    process_msg.error(f"处理失败：{e}")

        process_out = st.session_state.get("process_out")
        if process_out is not None:
            st.markdown("### 处理结果")
            st.json(process_out.summary)
            _download_button(label="下载处理结果.zip", data=process_out.processed_zip_bytes, file_name="处理结果.zip")

    with tab3:
        st.subheader("类型词典管理")
        type_rows = [{"英文": k, "中文": v} for k, v in st.session_state["type_dict"].items()]
        edited = st.data_editor(
            type_rows,
            num_rows="dynamic",
            width="stretch",
            hide_index=True,
            column_config={
                "英文": st.column_config.TextColumn(required=True),
                "中文": st.column_config.TextColumn(required=True),
            },
            key="type_dict_editor"
        )
        new_dict = {str(r.get("英文", "")).strip(): str(r.get("中文", "")).strip() for r in edited}
        new_dict = {k: v for k, v in new_dict.items() if k and v}
        
        if st.button("保存类型词典", type="primary"):
            save_type_dict(new_dict)
            st.session_state["type_dict"] = new_dict
            st.success("类型词典已保存！")

if __name__ == "__main__":
    main()
