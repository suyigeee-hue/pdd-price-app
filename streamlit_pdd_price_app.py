"""拼多多商品 Excel 批量改价 Streamlit 应用。

启动方式：
    pip install streamlit pandas openpyxl
    streamlit run streamlit_pdd_price_app.py

规则存储说明：
- 本地运行（未配置 GITHUB_TOKEN）：规则保存在本脚本同目录的 config.json。
- 云端运行（Streamlit Community Cloud 中配置了 GITHUB_TOKEN 等 Secrets）：
  规则通过 GitHub API 保存到仓库里的 config.json，永久有效、多人共用。
"""

from __future__ import annotations

import base64
import json
import math
import urllib.error
import urllib.request
from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
from openpyxl import load_workbook


st.set_page_config(page_title="拼多多商品批量改价", page_icon="💰", layout="centered")

CONFIG_PATH = Path(__file__).resolve().with_name("config.json")
REQUIRED_COLUMNS = ["商品名称", "规格名称", "单买价", "拼单价"]
GITHUB_API = "https://api.github.com"

# 首次运行时写入本地 config.json 的默认规则，可在页面中自行调整。
DEFAULT_RULES = [
    {"关键词": "10cm", "拼单价": 22.68},
    {"关键词": "20cm", "拼单价": 26.79},
    {"关键词": "30cm", "拼单价": 27.89},
    {"关键词": "40cm", "拼单价": 29.89},
    {"关键词": "鼠标垫", "拼单价": 34.69},
    {"关键词": "老鼠干30cm", "拼单价": 10.75},
    {"关键词": "老鼠干15cm", "拼单价": 9.68},
    {"关键词": "15cm", "拼单价": 9.68},
    {"关键词": "老鼠干", "拼单价": 10.75},
    {"关键词": "7cm猫耳", "拼单价": 24.89},
    {"关键词": "吱吱叫", "拼单价": 23.69},
    {"关键词": "吐泡泡", "拼单价": 24.68},

]


def cloud_mode() -> bool:
    """是否处于云端模式：配置了 GITHUB_TOKEN 时走 GitHub 仓库存储。"""
    try:
        _ = st.secrets["GITHUB_TOKEN"]
        return True
    except Exception:
        return False


def _github_headers() -> dict[str, str]:
    return {
        "Authorization": f"Bearer {st.secrets['GITHUB_TOKEN']}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }


def _github_repo() -> str:
    return st.secrets.get("GITHUB_REPO", "suyigeee-hue/pdd-price-app")


def _github_path() -> str:
    return st.secrets.get("GITHUB_CONFIG_PATH", "config.json")


def _github_get_file() -> list[dict[str, Any]] | None:
    """从 GitHub 仓库读取 config.json 并返回规则列表；文件不存在时返回 None。"""
    url = f"{GITHUB_API}/repos/{_github_repo()}/contents/{_github_path()}"
    request = urllib.request.Request(url, headers=_github_headers())
    try:
        with urllib.request.urlopen(request, timeout=15) as response:
            body = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as error:
        if error.code == 404:
            return None
        raise error
    content = body.get("content", "")
    raw_text = base64.b64decode(content).decode("utf-8")
    saved_config = json.loads(raw_text)
    rules = saved_config.get("price_rules", [])
    if not isinstance(rules, list):
        raise ValueError("price_rules 不是列表")
    return rules


def _github_write_file(rules: list[dict[str, Any]]) -> None:
    """把规则列表写回 GitHub 仓库的 config.json（自动附带 sha 完成更新）。"""
    url = f"{GITHUB_API}/repos/{_github_repo()}/contents/{_github_path()}"
    headers = _github_headers()

    # 先读取当前文件的 sha（GitHub 更新文件必须携带原 sha）。
    sha: str | None = None
    get_request = urllib.request.Request(url, headers=headers)
    try:
        with urllib.request.urlopen(get_request, timeout=15) as response:
            body = json.loads(response.read().decode("utf-8"))
            sha = body.get("sha")
    except urllib.error.HTTPError as error:
        if error.code != 404:
            raise error
        sha = None  # 仓库中尚无 config.json，直接新建。

    new_content = json.dumps({"price_rules": rules}, ensure_ascii=False, indent=2).encode("utf-8")
    payload = {
        "message": "Update price rules from web app",
        "content": base64.b64encode(new_content).decode("ascii"),
    }
    if sha:
        payload["sha"] = sha

    put_request = urllib.request.Request(
        url,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={**headers, "Content-Type": "application/json"},
        method="PUT",
    )
    with urllib.request.urlopen(put_request, timeout=15) as response:
        response.read()


def write_config(rules: list[dict[str, Any]]) -> bool:
    """保存规则。云端模式写回 GitHub 仓库；本地模式写本地 JSON。

    返回 True 表示保存成功；云端写入失败时返回 False（并回退写入本地文件）。
    """
    if cloud_mode():
        try:
            _github_write_file(rules)
            return True
        except Exception:
            # 云端写入失败时回退到本地文件，避免本次编辑丢失。
            try:
                with CONFIG_PATH.open("w", encoding="utf-8") as config_file:
                    json.dump({"price_rules": rules}, config_file, ensure_ascii=False, indent=2)
            except OSError:
                pass
            return False

    with CONFIG_PATH.open("w", encoding="utf-8") as config_file:
        json.dump({"price_rules": rules}, config_file, ensure_ascii=False, indent=2)
    return True


def load_config() -> list[dict[str, Any]]:
    """读取规则。云端模式优先从 GitHub 仓库读取，失败时回退本地；本地模式读本地文件。

    首次启动或配置损坏时使用并保存默认规则。
    """
    if cloud_mode():
        try:
            rules = _github_get_file()
            if rules is not None:
                return rules
        except Exception:
            pass  # 云端读取失败时回退到本地文件。

    if not CONFIG_PATH.exists():
        write_config(DEFAULT_RULES)
        return DEFAULT_RULES.copy()

    try:
        with CONFIG_PATH.open("r", encoding="utf-8") as config_file:
            saved_config = json.load(config_file)
        rules = saved_config.get("price_rules", [])
        if not isinstance(rules, list):
            raise ValueError("price_rules 不是列表")
        return rules
    except (json.JSONDecodeError, OSError, ValueError, AttributeError):
        # 配置损坏时不让网页无法启动，恢复为一份可编辑的默认配置。
        write_config(DEFAULT_RULES)
        return DEFAULT_RULES.copy()


def clean_rules(edited_rules: pd.DataFrame) -> list[dict[str, Any]]:
    """校验并清理编辑器中的规则，返回适合保存和匹配的规则列表。"""
    rules: list[dict[str, Any]] = []
    seen_keywords: set[str] = set()

    for row_number, row in edited_rules.iterrows():
        keyword = "" if pd.isna(row["关键词"]) else str(row["关键词"]).strip()

        # data_editor 新增的一整行空白内容不作为规则保存。
        if not keyword and pd.isna(row["拼单价"]):
            continue
        if not keyword:
            raise ValueError(f"第 {row_number + 1} 条规则缺少关键词。")
        if keyword in seen_keywords:
            raise ValueError(f"关键词“{keyword}”重复，请保留一条即可。")

        try:
            group_price = float(row["拼单价"])
        except (TypeError, ValueError):
            raise ValueError(f"关键词“{keyword}”的价格必须是数字。") from None

        if not math.isfinite(group_price):
            raise ValueError(f"关键词“{keyword}”的价格必须是有限数字。")

        seen_keywords.add(keyword)
        rules.append(
            {
                "关键词": keyword,
                "拼单价": group_price,
            }
        )

    if not rules:
        raise ValueError("请至少保留一条有效的价格规则。")
    return rules


def find_matched_rule(product_name: object, spec_name: object, rules: list[dict[str, Any]]) -> dict[str, Any] | None:
    """按关键词长度从长到短匹配，并在命中第一条后立即停止。"""
    product_text = "" if pd.isna(product_name) else str(product_name)
    spec_text = "" if pd.isna(spec_name) else str(spec_name)

    # reverse=True 表示“猫耳10cm”优先于“10cm”。长度相同则保持配置中的原有顺序。
    sorted_rules = sorted(rules, key=lambda rule: len(rule["关键词"]), reverse=True)
    for rule in sorted_rules:
        if rule["关键词"] in product_text or rule["关键词"] in spec_text:
            return rule  # 命中最高优先级规则后立即结束本行匹配。
    return None


def process_excel(source_bytes: bytes, rules: list[dict[str, Any]]) -> tuple[bytes, int]:
    """用 pandas 读取/匹配，只用 openpyxl 回写 E/F，最大程度保留原格式。"""
    # 读取第一个工作表；它应是上传文件中的商品数据表。
    excel_file = pd.ExcelFile(BytesIO(source_bytes))
    sheet_name = excel_file.sheet_names[0]
    dataframe = pd.read_excel(excel_file, sheet_name=sheet_name, dtype=object)

    missing_columns = [column for column in REQUIRED_COLUMNS if column not in dataframe.columns]
    if missing_columns:
        raise ValueError(
            "第一个工作表缺少表头：" + "、".join(missing_columns)
            + "。请确认表头位于第 1 行，且名称与要求一致。"
        )

    # 打开同一份原始工作簿，仅修改第 5、6 列（E/F）第 2 行及之后的单元格。
    # A列商品ID和C列SKUID在此流程中完全不会被写入。
    workbook = load_workbook(BytesIO(source_bytes))
    worksheet = workbook[sheet_name]

    updated_count = 0
    for excel_row, (product_name, spec_name) in enumerate(
        zip(dataframe["商品名称"], dataframe["规格名称"]), start=2
    ):
        matched_rule = find_matched_rule(product_name, spec_name, rules)
        if matched_rule is None:
            continue  # 没有关键词命中时，原 E/F 价格保持不变。

        group_price = matched_rule["拼单价"]
        worksheet.cell(row=excel_row, column=5).value = group_price + 2  # E：单买价 = 拼单价 + 2
        worksheet.cell(row=excel_row, column=6).value = group_price  # F：拼单价 = 配置价格
        updated_count += 1

    result = BytesIO()
    workbook.save(result)
    return result.getvalue(), updated_count


def main() -> None:
    st.title("💰 拼多多商品 Excel 批量改价")
    st.caption("配置价格写入 F 列“拼单价”，E 列“单买价”自动为拼单价 + 2；商品 ID、SKUID 与原有表头样式不会被改动。")

    uploaded_file = st.file_uploader("上传拼多多商品 Excel 文件（.xlsx）", type=["xlsx"])

    st.subheader("价格规则")
    if cloud_mode():
        st.caption("当前规则存储：云端 GitHub 仓库（保存后永久生效，所有人共用同一套规则）。")
    else:
        st.caption("当前规则存储：本地 config.json（仅本机网页生效）。")

    current_rules = load_config()
    rules_dataframe = pd.DataFrame(current_rules, columns=["关键词", "拼单价"])
    edited_rules = st.data_editor(
        rules_dataframe,
        column_config={
            "关键词": st.column_config.TextColumn("关键词", required=True),
            "拼单价": st.column_config.NumberColumn("拼单价", min_value=0.0, format="%.2f", required=True),
        },
        num_rows="dynamic",
        hide_index=True,
        use_container_width=True,
        key="price_rules_editor",
    )

    if st.button("保存配置", use_container_width=True):
        try:
            valid_rules = clean_rules(edited_rules)
            saved = write_config(valid_rules)
            if saved and cloud_mode():
                st.success(f"已保存 {len(valid_rules)} 条规则到云端 GitHub 仓库，永久有效。")
            elif saved:
                st.success(f"已保存 {len(valid_rules)} 条规则到 {CONFIG_PATH.name}。")
            else:
                st.error("云端保存失败（可能网络不通或令牌无效），本次修改仅临时保存在云端容器，重启后会恢复。请检查设置。")
        except ValueError as error:
            st.error(f"配置未保存：{error}")

    st.divider()
    if st.button("开始处理", type="primary", use_container_width=True):
        if uploaded_file is None:
            st.warning("请先上传一个 .xlsx 文件。")
        else:
            try:
                # 即使尚未点击“保存配置”，处理也会使用编辑器内当前显示的规则。
                valid_rules = clean_rules(edited_rules)
                output_bytes, updated_count = process_excel(uploaded_file.getvalue(), valid_rules)
                st.session_state["output_excel"] = output_bytes
                st.session_state["output_filename"] = f"已改价_{uploaded_file.name}"
                st.success(f"处理完成，共更新 {updated_count} 行。")
            except Exception as error:
                st.error(f"处理失败：{error}")

    if "output_excel" in st.session_state:
        st.download_button(
            "下载处理后的 Excel 文件",
            data=st.session_state["output_excel"],
            file_name=st.session_state["output_filename"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )


if __name__ == "__main__":
    main()
