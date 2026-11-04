"""拼多多商品 Excel 批量改价 Streamlit 应用。"""

from __future__ import annotations

import json
import math
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook


st.set_page_config(page_title="拼多多商品批量改价", page_icon="💰", layout="centered")

CONFIG_PATH = Path(__file__).resolve().with_name("config.json")
REQUIRED_COLUMNS = ["商品名称", "规格名称", "单买价", "拼单价"]

# 默认规则
DEFAULT_RULES = [
    {"关键词": "10cm", "拼单价": 22.99},
    {"关键词": "20cm", "拼单价": 26.99},
    {"关键词": "30cm", "拼单价": 27.89},
]


def load_config() -> list[dict]:
    """读取规则配置。"""
    if not CONFIG_PATH.exists():
        return DEFAULT_RULES.copy()
    try:
        with CONFIG_PATH.open("r", encoding="utf-8") as f:
            saved = json.load(f)
        rules = saved.get("price_rules", [])
        if isinstance(rules, list):
            return rules
    except Exception:
        pass
    return DEFAULT_RULES.copy()


def write_config(rules: list[dict]) -> bool:
    """保存规则到本地文件。"""
    try:
        with CONFIG_PATH.open("w", encoding="utf-8") as f:
            json.dump({"price_rules": rules}, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


def clean_rules(df: pd.DataFrame) -> list[dict]:
    """清理并校验规则。"""
    rules = []
    seen = set()
    for _, row in df.iterrows():
        kw = "" if pd.isna(row["关键词"]) else str(row["关键词"]).strip()
        if not kw and pd.isna(row["拼单价"]):
            continue
        if not kw:
            raise ValueError("有关键词为空，请填写或删除空行。")
        if kw in seen:
            raise ValueError(f"关键词“{kw}”重复了。")
        try:
            price = float(row["拼单价"])
        except (TypeError, ValueError):
            raise ValueError(f"关键词“{kw}”的价格不是数字。") from None
        if not math.isfinite(price):
            raise ValueError(f"关键词“{kw}”的价格无效。")
        seen.add(kw)
        rules.append({"关键词": kw, "拼单价": price})
    if not rules:
        raise ValueError("请至少保留一条规则。")
    return rules


def find_matched_rule(product_name, spec_name, rules):
    """匹配关键词规则。"""
    p_text = "" if pd.isna(product_name) else str(product_name)
    s_text = "" if pd.isna(spec_name) else str(spec_name)
    sorted_rules = sorted(rules, key=lambda r: len(r["关键词"]), reverse=True)
    for rule in sorted_rules:
        if rule["关键词"] in p_text or rule["关键词"] in s_text:
            return rule
    return None


def process_excel(source_bytes: bytes, rules: list[dict]) -> tuple[bytes, int]:
    """处理 Excel 文件。"""
    excel_file = pd.ExcelFile(BytesIO(source_bytes))
    sheet_name = excel_file.sheet_names[0]
    df = pd.read_excel(excel_file, sheet_name=sheet_name, dtype=object)

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError("Excel 缺少表头：" + "、".join(missing))

    wb = load_workbook(BytesIO(source_bytes))
    ws = wb[sheet_name]

    count = 0
    for row_idx, (p_name, s_name) in enumerate(
        zip(df["商品名称"], df["规格名称"]), start=2
    ):
        rule = find_matched_rule(p_name, s_name, rules)
        if rule is None:
            continue
        ws.cell(row=row_idx, column=5).value = rule["拼单价"] + 2
        ws.cell(row=row_idx, column=6).value = rule["拼单价"]
        count += 1

    out = BytesIO()
    wb.save(out)
    return out.getvalue(), count


def main():
    st.title("💰 拼多多商品 Excel 批量改价")
    st.caption("配置价格写入 F 列“拼单价”，E 列“单买价”自动为拼单价 + 2。")

    uploaded = st.file_uploader("上传拼多多商品 Excel 文件（.xlsx）", type=["xlsx"])

    st.subheader("价格规则")
    current_rules = load_config()
    df = pd.DataFrame(current_rules, columns=["关键词", "拼单价"])
    edited = st.data_editor(
        df,
        column_config={
            "关键词": st.column_config.TextColumn("关键词", required=True),
            "拼单价": st.column_config.NumberColumn("拼单价", min_value=0.0, format="%.2f", required=True),
        },
        num_rows="dynamic",
        hide_index=True,
        use_container_width=True,
        key="rules_editor",
    )

    if st.button("保存配置", use_container_width=True):
        try:
            valid = clean_rules(edited)
            if write_config(valid):
                st.success(f"已保存 {len(valid)} 条规则。")
            else:
                st.error("保存失败。")
        except ValueError as e:
            st.error(f"保存失败：{e}")

    st.divider()
    if st.button("开始处理", type="primary", use_container_width=True):
        if uploaded is None:
            st.warning("请先上传 Excel 文件。")
        else:
            try:
                valid = clean_rules(edited)
                out_bytes, count = process_excel(uploaded.getvalue(), valid)
                st.session_state["output"] = out_bytes
                st.session_state["filename"] = f"已改价_{uploaded.name}"
                st.success(f"处理完成，共更新 {count} 行。")
            except Exception as e:
                st.error(f"处理失败：{e}")

    if "output" in st.session_state:
        st.download_button(
            "下载处理后的 Excel",
            data=st.session_state["output"],
            file_name=st.session_state["filename"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )


if __name__ == "__main__":
    main()
