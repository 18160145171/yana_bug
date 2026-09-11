#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import datetime as dt
import re
from pathlib import Path

import streamlit as st

from generate_bug_report import DEFAULT_REPORT_PROMPT, build_html, load_rows, validate_columns


BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
ALLOWED_EXTENSIONS = {".csv", ".xlsx", ".xls"}
APP_VERSION = "2026.09.11-streamlit-ai-strict"


def safe_filename(filename):
    name = Path(filename or "bugs.xlsx").name
    cleaned = re.sub(r"[^A-Za-z0-9._\-\u4e00-\u9fff]+", "_", name).strip("._")
    return cleaned or "bugs.xlsx"


def save_uploaded_file(uploaded_file):
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    upload_path = UPLOAD_DIR / f"{timestamp}_{safe_filename(uploaded_file.name)}"
    upload_path.write_bytes(uploaded_file.getvalue())
    return timestamp, upload_path


def generate_report(uploaded_file, project_name, llm_config):
    suffix = Path(uploaded_file.name).suffix.lower()
    if suffix not in ALLOWED_EXTENSIONS:
        raise RuntimeError("仅支持 .csv / .xlsx / .xls 文件。")

    timestamp, upload_path = save_uploaded_file(uploaded_file)
    headers, rows = load_rows(str(upload_path))
    validate_columns(headers)

    report_project_name = project_name.strip() or Path(uploaded_file.name).stem
    rendered = build_html(report_project_name, rows, llm_config=llm_config)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_name = f"质量验收与缺陷分析报告_{timestamp}.doc"
    output_path = OUTPUT_DIR / output_name
    output_path.write_text(rendered, encoding="utf-8")
    return output_name, output_path, rendered


st.set_page_config(
    page_title="质量验收报告生成器",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.title("质量验收与缺陷分析报告生成器")
st.caption(f"版本：{APP_VERSION}")

project_name = st.text_input("项目名称", placeholder="例如：视频编辑器App")
bug_file = st.file_uploader(
    "上传 Bug 文档",
    type=["csv", "xlsx", "xls"],
    accept_multiple_files=False,
)
use_ai = st.toggle("启用 AI 增强分析", value=False)

api_base_url = ""
api_key = ""
model_name = ""
report_prompt = DEFAULT_REPORT_PROMPT
if use_ai:
    st.info("AI 模式已开启：生成时必须成功调用模型接口；接口失败不会回退生成固定模板报告。")
    api_base_url = st.text_input("API Base URL", placeholder="例如：https://api.openai.com/v1")
    api_key = st.text_input("API Key", type="password")
    model_name = st.text_input("模型名称", placeholder="例如：gpt-4.1 / glm-4-plus")
    st.subheader("报告生成规则")
    report_prompt = st.text_area(
        "可编辑提示词",
        value=DEFAULT_REPORT_PROMPT,
        height=190,
        help="每次启用 AI 生成报告时，模型都会根据这段规则重新分析上传的 Bug 数据。留空时使用推荐规则。",
    )

with st.form("report-form"):
    submitted = st.form_submit_button("生成质量验收报告")

if submitted:
    if bug_file is None:
        st.error("请选择要上传的 Bug 文档。")
    elif use_ai and (not api_base_url.strip() or not api_key.strip() or not model_name.strip()):
        st.error("启用 AI 时，请填写 API Base URL、API Key 和模型名称。")
    else:
        llm_config = None
        if use_ai:
            llm_config = {
                "enabled": True,
                "api_base_url": api_base_url.strip(),
                "api_key": api_key.strip(),
                "model_name": model_name.strip(),
                "report_prompt": report_prompt.strip() or DEFAULT_REPORT_PROMPT,
            }

        progress_text = "正在调用 AI 接口并生成报告..." if use_ai else "正在生成报告..."
        with st.spinner(progress_text):
            try:
                output_name, output_path, rendered = generate_report(
                    uploaded_file=bug_file,
                    project_name=project_name,
                    llm_config=llm_config,
                )
            except Exception as exc:  # pylint: disable=broad-except
                st.error(f"生成失败：{exc}")
            else:
                st.success("报告生成成功，AI 接口调用成功。")
                st.download_button(
                    label="下载报告（.doc）",
                    data=rendered.encode("utf-8"),
                    file_name=output_name,
                    mime="application/msword",
                )
                st.caption(f"云端临时保存路径：{output_path}")
