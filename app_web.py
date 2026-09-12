#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import datetime as dt
from pathlib import Path

from flask import Flask, jsonify, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

from generate_bug_report import (
    DEFAULT_API_ENDPOINT,
    DEFAULT_REPORT_PROMPT,
    build_html,
    load_rows,
    validate_columns,
)

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
ALLOWED_EXTENSIONS = {".csv", ".xlsx", ".xls"}
APP_VERSION = "2026.09.12-api-endpoint"

UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)


@app.after_request
def add_no_cache_headers(response):
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    return response


def allowed_file(filename):
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


@app.route("/", methods=["GET", "POST"])
def index():
    return render_template(
        "index.html",
        app_version=APP_VERSION,
        default_report_prompt=DEFAULT_REPORT_PROMPT,
    )


@app.route("/api/generate", methods=["POST"])
def api_generate():
    file = request.files.get("bug_file")
    project_name = (request.form.get("project_name") or "").strip()
    use_ai = (request.form.get("use_ai") or "0").strip() == "1"
    api_base_url = (request.form.get("api_base_url") or "").strip()
    api_endpoint = (request.form.get("api_endpoint") or "").strip() or DEFAULT_API_ENDPOINT
    api_key = (request.form.get("api_key") or "").strip()
    model_name = (request.form.get("model_name") or "").strip()
    report_prompt = (request.form.get("report_prompt") or "").strip() or DEFAULT_REPORT_PROMPT

    if not file or not file.filename:
        return jsonify({"ok": False, "error": "请选择要上传的 Bug 文档（CSV/Excel）。"}), 400

    if not allowed_file(file.filename):
        return jsonify({"ok": False, "error": "仅支持 .csv / .xlsx / .xls 文件。"}), 400

    safe_name = secure_filename(file.filename)
    timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    upload_path = UPLOAD_DIR / f"{timestamp}_{safe_name}"
    file.save(upload_path)

    try:
        headers, rows = load_rows(str(upload_path))
        validate_columns(headers)
        if not project_name:
            project_name = Path(file.filename).stem
        llm_config = None
        if use_ai:
            if not api_base_url or not api_endpoint or not api_key or not model_name:
                return (
                    jsonify({"ok": False, "error": "启用 AI 时，请填写接口地址、接口端点、API Key 和模型标识。"}),
                    400,
                )
            llm_config = {
                "enabled": True,
                "api_base_url": api_base_url,
                "api_endpoint": api_endpoint,
                "api_key": api_key,
                "model_name": model_name,
                "report_prompt": report_prompt,
            }
        rendered = build_html(project_name, rows, llm_config=llm_config)

        out_name = f"质量验收与缺陷分析报告_{timestamp}.doc"
        out_path = OUTPUT_DIR / out_name
        out_path.write_text(rendered, encoding="utf-8")

        return jsonify(
            {
                "ok": True,
                "message": "报告生成成功，AI 接口调用成功。" if use_ai else "报告生成成功。",
                "download_link": url_for("download_file", filename=out_name),
                "output_path": str(out_path),
            }
        )
    except Exception as exc:  # pylint: disable=broad-except
        return jsonify({"ok": False, "error": f"生成失败：{exc}"}), 400


@app.route("/download/<path:filename>", methods=["GET"])
def download_file(filename):
    file_path = OUTPUT_DIR / filename
    if not file_path.exists():
        return redirect(url_for("index"))
    return send_file(file_path, as_attachment=True, download_name=filename)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=8512, debug=False)
