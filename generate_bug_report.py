#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import csv
import datetime as dt
import html
import json
import os
import re
import sys
import urllib.error
import urllib.request
from collections import Counter, defaultdict
from pathlib import Path

REQUIRED_COLUMNS = []
OPTIONAL_COLUMNS = ["标题", "任务ID", "执行者", "父任务", "任务状态", "优先级", "参与者"]

COLUMN_ALIASES = {
    "标题": ["标题", "bug标题", "缺陷标题", "问题标题", "title", "summary", "主题"],
    "任务ID": ["任务id", "任务编号", "单号", "id", "bugid", "缺陷id", "ticketid"],
    "执行者": ["执行者", "处理人", "责任人", "指派给", "assignee", "owner", "执行人"],
    "父任务": ["父任务", "父任务id", "上级任务", "所属任务", "parent", "parentid"],
    "任务状态": ["任务状态", "状态", "bug状态", "缺陷状态", "当前状态", "status"],
    "优先级": ["优先级", "严重程度", "等级", "priority", "severity"],
    "参与者": ["参与者", "参与人", "抄送人", "协作者", "关注人", "watchers", "cc"],
}

UNCLOSED_STATUS_ORDER = ["重新打开", "待处理", "开发中", "暂不处理"]
CLOSED_STATUS_ORDER = ["已解决", "已完成", "开发完成"]

STATUS_GROUP_INDEX = {s: 0 for s in UNCLOSED_STATUS_ORDER}
STATUS_GROUP_INDEX.update({s: 1 for s in CLOSED_STATUS_ORDER})

ALL_STATUS_ORDER = UNCLOSED_STATUS_ORDER + CLOSED_STATUS_ORDER
STATUS_INDEX = {status: idx for idx, status in enumerate(ALL_STATUS_ORDER)}

HIGH_PRIORITY_KEYWORDS = ["高", "紧急", "严重", "critical", "p0", "p1", "blocker"]
UI_KEYWORDS = ["ui", "页面", "显示", "样式", "布局", "交互", "按钮", "弹窗", "对齐", "颜色", "字体"]
MODEL_MAX_ROWS = 500

DEFAULT_REPORT_PROMPT = """请以优秀测试软件工程师和测试经理的视角，重新分析输入的 Bug 数据，生成正式、客观、简洁、可执行的质量验收与缺陷分析结论。
严格以输入数据为依据，不臆造版本、环境、复现步骤、责任人或修复状态；输入缺失的信息请明确标注“数据未提供”。
重点分析 Bug 总量、有效记录、优先级分布、功能模块分布、执行人分布、任务状态、已解决率、未关闭问题和高优先级未闭环问题。
识别缺陷集中模块、发布风险、数据质量问题（例如标题不完整、状态与优先级异常或关键信息缺失），并给出有依据的发布建议和后续行动。
结论要说明判断依据；行动建议要具体到责任角色、优先级或验证动作。请使用中文。"""

MODULE_RULES = [
    ("去水印/去字幕", ["去水印", "去字幕", "logo"]),
    ("数据埋点", ["埋点"]),
    ("多语言适配", ["多语言"]),
    ("视频翻译", ["视频翻译"]),
    ("AI字幕", ["ai字幕", "ai 字幕", "智能字幕"]),
    ("设置页", ["设置页", "设置页面"]),
    ("编辑页", ["编辑页", "编辑页面"]),
]


def normalize_text(value):
    if value is None:
        return ""
    return str(value).strip()


def normalize_header_name(name):
    raw = normalize_text(name).lower()
    # 移除空格、下划线和常见标点，提升列名兼容性
    return re.sub(r"[\s\-_（）()\[\]【】:：]+", "", raw)


def build_header_mapping(headers):
    normalized_to_original = {normalize_header_name(h): h for h in headers}
    mapping = {}

    for canonical, alias_list in COLUMN_ALIASES.items():
        found = None
        for alias in alias_list:
            alias_norm = normalize_header_name(alias)
            if alias_norm in normalized_to_original:
                found = normalized_to_original[alias_norm]
                break

        if not found:
            # 包含匹配，兼容如“状态(当前)”这类表头
            for normalized_header, original_header in normalized_to_original.items():
                if any(normalize_header_name(alias) in normalized_header for alias in alias_list):
                    found = original_header
                    break

        if found:
            mapping[canonical] = found

    return mapping


def split_people(text):
    value = normalize_text(text)
    if not value:
        return []
    parts = re.split(r"[，,、;/\\\s]+", value)
    return [p.strip() for p in parts if p.strip()]


def is_high_priority(priority):
    p = normalize_text(priority).lower()
    if not p:
        return False
    return any(token in p for token in HIGH_PRIORITY_KEYWORDS)


def classify_module(title):
    text = normalize_text(title).lower()
    for module_name, keywords in MODULE_RULES:
        if any(k.lower() in text for k in keywords):
            return module_name

    if "设置" in text:
        return "设置页"
    if "编辑" in text:
        return "编辑页"
    if "字幕" in text and "ai" in text:
        return "AI字幕"
    if "翻译" in text and "视频" in text:
        return "视频翻译"
    return "其他模块"


def issue_nature_by_titles(titles):
    if not titles:
        return "功能逻辑问题"
    ui_hits = 0
    for title in titles:
        t = normalize_text(title).lower()
        if any(k in t for k in UI_KEYWORDS):
            ui_hits += 1
    ratio = ui_hits / len(titles)
    return "UI交互问题" if ratio >= 0.5 else "功能逻辑问题"


def parse_task_id(task_id):
    text = normalize_text(task_id)
    nums = re.findall(r"\d+", text)
    if nums:
        return int(nums[-1]), text
    return sys.maxsize, text


def load_rows(input_file):
    ext = Path(input_file).suffix.lower()
    if ext == ".csv":
        with open(input_file, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            rows = [dict(r) for r in reader]
            headers = reader.fieldnames or []
        return harmonize_columns(headers, rows)

    if ext in [".xlsx", ".xls"]:
        try:
            import pandas as pd  # pylint: disable=import-outside-toplevel
        except ImportError as exc:
            raise RuntimeError(
                "读取 Excel 需要 pandas/openpyxl。请先执行：pip install pandas openpyxl"
            ) from exc
        df = pd.read_excel(input_file, dtype=str).fillna("")
        headers = list(df.columns)
        rows = df.to_dict(orient="records")
        return harmonize_columns(headers, rows)

    raise RuntimeError("仅支持 .csv / .xlsx / .xls 文件。")


def validate_columns(headers):
    if not headers:
        raise RuntimeError("未识别到可用字段，请检查表头行是否存在。")


def harmonize_columns(headers, rows):
    mapping = build_header_mapping(headers)

    normalized_rows = []
    for row in rows:
        normalized_row = {}
        for canonical in OPTIONAL_COLUMNS:
            source_col = mapping.get(canonical)
            normalized_row[canonical] = normalize_text(row.get(source_col, "")) if source_col else ""

        # 参与者缺失时，默认沿用执行者
        if not normalized_row.get("参与者"):
            normalized_row["参与者"] = normalized_row.get("执行者", "")

        normalized_rows.append(normalized_row)

    detected_headers = list(mapping.keys())
    return detected_headers, normalized_rows


def density_weight(count):
    if count > 5:
        return 3
    if 2 <= count <= 5:
        return 2
    return 1


def format_percent(numerator, denominator):
    if denominator == 0:
        return "0.00%"
    return f"{(numerator / denominator) * 100:.2f}%"


def sort_rows_for_appendix(rows):
    def sort_key(row):
        status = normalize_text(row.get("任务状态"))
        group = STATUS_GROUP_INDEX.get(status, 2)
        status_idx = STATUS_INDEX.get(status, 999)
        high_rank = 0 if is_high_priority(row.get("优先级")) else 1
        task_num, task_raw = parse_task_id(row.get("任务ID"))
        return (group, status_idx, high_rank, task_num, task_raw)

    return sorted(rows, key=sort_key)


def ensure_f_drive_or_desktop():
    f_drive = Path("F:/")
    if f_drive.exists():
        return f_drive
    desktop = Path.home() / "Desktop"
    return desktop


def call_openai_compatible_chat(api_base_url, api_key, model_name, messages, timeout_sec=45):
    base = normalize_text(api_base_url).rstrip("/")
    if not base:
        raise RuntimeError("API Base URL 不能为空。")
    if base.endswith("/chat/completions"):
        endpoint = base
    else:
        endpoint = f"{base}/chat/completions"

    normalized_model = normalize_text(model_name)
    if "open.bigmodel.cn" in base.lower():
        alias_map = {
            "智谱": "glm-4-plus",
            "zhipu": "glm-4-plus",
            "zhipuai": "glm-4-plus",
            "glm4": "glm-4-plus",
        }
        normalized_model = alias_map.get(normalized_model.lower(), normalized_model) if normalized_model else "glm-4-plus"
        if normalized_model == "智谱":
            normalized_model = "glm-4-plus"
        if "." not in normalize_text(api_key):
            raise RuntimeError(
                "检测到智谱接口，但 API Key 格式疑似不正确。"
                "通常应为 'id.secret' 形式，请从智谱控制台重新复制。"
            )

    payload = {
        "model": normalized_model,
        "messages": messages,
        "temperature": 0.3,
    }
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        endpoint,
        data=body,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout_sec) as resp:
            raw = resp.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        if exc.code == 401:
            raise RuntimeError(
                "模型接口返回 401（鉴权失败）。请检查："
                "1) API Key 是否有效且完整；"
                "2) Base URL 是否为 https://open.bigmodel.cn/api/paas/v4；"
                "3) 模型名称是否为 glm-4-plus / glm-4-air / glm-4-flash。"
                f" 原始返回：{detail[:300]}"
            ) from exc
        raise RuntimeError(f"模型接口错误 HTTP {exc.code}: {detail[:500]}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"模型接口连接失败: {exc.reason}") from exc

    try:
        data = json.loads(raw)
        content = data["choices"][0]["message"]["content"]
    except Exception as exc:  # pylint: disable=broad-except
        raise RuntimeError(f"模型返回格式异常: {raw[:500]}") from exc

    if isinstance(content, list):
        return "".join(
            item.get("text", "") if isinstance(item, dict) else str(item)
            for item in content
        ).strip()
    return normalize_text(content)


def build_ai_insight(project_name, total, resolved_count, high_count, module_analysis, top1_module, top1_nature, rows, llm_config):
    unresolved = []
    for row in rows:
        status = normalize_text(row.get("任务状态"))
        if status in UNCLOSED_STATUS_ORDER:
            unresolved.append(
                {
                    "title": normalize_text(row.get("标题")),
                    "status": status,
                    "priority": normalize_text(row.get("优先级")),
                    "owner": normalize_text(row.get("执行者")),
                }
            )
    unresolved = unresolved[:15]

    module_stats = [
        {
            "module": item["module"],
            "count": item["count"],
            "density": round(item["density"], 2),
        }
        for item in module_analysis[:8]
    ]

    model_rows = sort_rows_for_appendix(rows)[:MODEL_MAX_ROWS]
    context = {
        "project_name": project_name,
        "total_bug_count": total,
        "resolved_count": resolved_count,
        "high_priority_count": high_count,
        "resolved_rate": format_percent(resolved_count, total),
        "high_priority_rate": format_percent(high_count, total),
        "top1_module": top1_module,
        "top1_issue_nature": top1_nature,
        "module_stats": module_stats,
        "sample_unresolved_bugs": unresolved,
        "bug_rows": model_rows,
        "bug_rows_total": total,
        "bug_rows_included": len(model_rows),
        "bug_rows_truncated": total > len(model_rows),
    }

    report_prompt = normalize_text(llm_config.get("report_prompt")) or DEFAULT_REPORT_PROMPT
    messages = [
        {
            "role": "system",
            "content": (
                "你是优秀测试软件工程师和资深测试经理。"
                "请根据用户提供的报告生成规则，重新分析 Bug 明细和统计信息，生成客观、可执行的发布评估。"
                "只输出 JSON，不要输出 Markdown 代码块。"
            ),
        },
        {
            "role": "user",
            "content": (
                "【报告生成规则提示词】\n"
                f"{report_prompt}\n\n"
                "【输出格式约束】\n"
                "请严格返回 JSON 对象，字段必须包含："
                "executive_summary, key_risks(数组), release_decision, actions(数组)。"
                "其中 release_decision 只能是 GO / NO-GO / CONDITIONAL-GO。"
                "每个 action 最多 30 字。不得输出 JSON 以外的内容。"
                "以下是本次上传文件经过字段标准化后的统计和 Bug 明细；"
                "如果 bug_rows_truncated 为 true，说明明细过多，仅展示优先级更高或未关闭的前 500 条，"
                "请结合总数和统计字段进行判断：\n"
                + json.dumps(context, ensure_ascii=False)
            ),
        },
    ]

    content = call_openai_compatible_chat(
        api_base_url=llm_config.get("api_base_url", ""),
        api_key=llm_config.get("api_key", ""),
        model_name=llm_config.get("model_name", ""),
        messages=messages,
    )

    cleaned = content.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
        cleaned = re.sub(r"\s*```$", "", cleaned)

    result = json.loads(cleaned)
    summary = normalize_text(result.get("executive_summary"))
    decision = normalize_text(result.get("release_decision")).upper()
    risks = result.get("key_risks", [])
    actions = result.get("actions", [])

    if decision not in {"GO", "NO-GO", "CONDITIONAL-GO"}:
        decision = "CONDITIONAL-GO"

    if not isinstance(risks, list):
        risks = []
    if not isinstance(actions, list):
        actions = []

    risk_lines = "".join([f"<li>{html.escape(normalize_text(r))}</li>" for r in risks if normalize_text(r)])
    action_lines = "".join([f"<li>{html.escape(normalize_text(a))}</li>" for a in actions if normalize_text(a)])
    if not risk_lines:
        risk_lines = "<li>暂无模型返回风险项。</li>"
    if not action_lines:
        action_lines = "<li>暂无模型返回行动项。</li>"

    return (
        "<h2>五、 AI 增强洞察（可选）</h2>"
        f"<div class='summary-box'><b>发布建议：</b>{html.escape(decision)}<br>"
        f"<b>结论摘要：</b>{html.escape(summary or '无')}</div>"
        "<div><b>关键风险：</b><ul>"
        f"{risk_lines}"
        "</ul></div>"
        "<div><b>行动建议：</b><ul>"
        f"{action_lines}"
        "</ul></div>"
    )


def build_html(project_name, rows, llm_config=None):
    total = len(rows)
    resolved_count = sum(1 for r in rows if normalize_text(r.get("任务状态")) == "已解决")
    high_count = sum(1 for r in rows if is_high_priority(r.get("优先级")))

    participants = set()
    module_counter = Counter()
    module_titles = defaultdict(list)

    for row in rows:
        for person in split_people(row.get("执行者")):
            participants.add(person)
        title = normalize_text(row.get("标题"))
        module = classify_module(title)
        module_counter[module] += 1
        module_titles[module].append(title)

    module_analysis = []
    for module_name, count in module_counter.items():
        weight = density_weight(count)
        density = count / weight
        module_analysis.append(
            {
                "module": module_name,
                "count": count,
                "weight": weight,
                "density": density,
            }
        )
    module_analysis.sort(key=lambda x: (-x["count"], x["module"]))

    top1_module = module_analysis[0]["module"] if module_analysis else "无"
    top1_nature = issue_nature_by_titles(module_titles.get(top1_module, []))

    sorted_rows = sort_rows_for_appendix(rows)
    people_text = "、".join(sorted(participants)) if participants else "无"

    summary_box = (
        f"Bug总数：<b>{total}</b>；"
        f"解决率（已解决/总数）：<b>{format_percent(resolved_count, total)}</b>；"
        f"高优先级占比：<b>{format_percent(high_count, total)}</b>；"
        f"参与人员：<b>{html.escape(people_text)}</b>"
    )

    top_risk = "当前未发现明显风险。"
    if total > 0 and module_analysis:
        top_risk = (
            f"当前缺陷主要集中在「{html.escape(top1_module)}」模块，"
            f"Top1 模块倾向于 <b>{top1_nature}</b>，建议优先投入专项回归资源。"
        )

    module_rows_html = "".join(
        [
            (
                "<tr>"
                f"<td>{idx + 1}</td>"
                f"<td>{html.escape(item['module'])}</td>"
                f"<td>{item['count']}</td>"
                f"<td>{item['weight']}</td>"
                f"<td>{item['density']:.2f}</td>"
                "</tr>"
            )
            for idx, item in enumerate(module_analysis)
        ]
    )

    if not module_rows_html:
        module_rows_html = "<tr><td colspan='5'>暂无数据</td></tr>"

    improvements = [
        (
            "Top模块专项治理",
            f"针对「{html.escape(top1_module)}」建立专项缺陷清单，按 {top1_nature} 维度拆分并逐项闭环。",
        ),
        (
            "高优先级前置拦截",
            "将高优先级问题纳入每日跟踪，新增提测门禁：高优先级未清零禁止发布。",
        ),
        (
            "回归用例增强",
            "沉淀高频问题场景为自动化回归用例，覆盖跨页面、跨状态、跨语言主流程。",
        ),
    ]

    improvements_html = "".join(
        [
            (
                "<div class='improvement-item'>"
                f"<span class='improvement-title'>{html.escape(title)}</span>"
                f"<span>{html.escape(content)}</span>"
                "</div>"
            )
            for title, content in improvements
        ]
    )

    appendix_rows_html = ""
    for row in sorted_rows:
        status = normalize_text(row.get("任务状态"))
        cls = "status-unclosed" if status in UNCLOSED_STATUS_ORDER else "status-closed"
        appendix_rows_html += (
            "<tr>"
            f"<td>{html.escape(normalize_text(row.get('标题')))}</td>"
            f"<td>{html.escape(normalize_text(row.get('任务ID')))}</td>"
            f"<td>{html.escape(normalize_text(row.get('执行者')))}</td>"
            f"<td>{html.escape(normalize_text(row.get('父任务')))}</td>"
            f"<td class='{cls}'>{html.escape(status)}</td>"
            f"<td>{html.escape(normalize_text(row.get('优先级')))}</td>"
            "</tr>"
        )
    if not appendix_rows_html:
        appendix_rows_html = "<tr><td colspan='6'>暂无缺陷数据</td></tr>"

    ai_section_html = ""
    if llm_config and llm_config.get("enabled"):
        try:
            ai_section_html = build_ai_insight(
                project_name=project_name,
                total=total,
                resolved_count=resolved_count,
                high_count=high_count,
                module_analysis=module_analysis,
                top1_module=top1_module,
                top1_nature=top1_nature,
                rows=rows,
                llm_config=llm_config,
            )
        except Exception as exc:  # pylint: disable=broad-except
            ai_section_html = (
                "<h2>五、 AI 增强洞察（可选）</h2>"
                f"<div class='risk-box'>AI 增强调用失败，已回退为规则模板输出。原因：{html.escape(str(exc))}</div>"
            )

    html_doc = f"""<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
<head>
    <meta charset="utf-8">
    <style>
        body {{ font-family: 'Microsoft YaHei', sans-serif; line-height: 1.6; color: #333; max-width: 800px; margin: 0 auto; }}
        h1 {{ font-size: 24pt; color: #2c3e50; text-align: center; border-bottom: 2px solid #3498db; margin: 30px 0; padding-bottom: 20px; }}
        h2 {{ font-size: 16pt; color: #2980b9; background-color: #ecf0f1; padding: 10px; border-left: 5px solid #2980b9; margin-top: 25px; }}
        table {{ width: 100%; border-collapse: collapse; margin-bottom: 20px; font-size: 10.5pt; }}
        th {{ background-color: #3498db; color: white; padding: 10px; text-align: left; border: 1px solid #2980b9; }}
        td {{ padding: 8px 10px; border: 1px solid #ddd; }}
        tr:nth-child(even) {{ background-color: #f9f9f9; }}
        .improvement-item {{ margin-bottom: 15px; padding-bottom: 15px; border-bottom: 1px dashed #eee; }}
        .improvement-title {{ font-weight: bold; color: #2c3e50; display: block; }}
        .summary-box {{ background-color: #fcf8e3; border: 1px solid #faebcc; padding: 15px; margin-bottom: 20px; }}
        .risk-box {{ border: 1px solid #ebccd1; background-color: #f2dede; color: #a94442; padding: 15px; }}
        .status-unclosed {{ color: #e74c3c; font-weight: bold; }}
        .status-closed {{ color: #27ae60; }}
    </style>
</head>
<body>
    <h1>{html.escape(project_name)} - 质量验收报告</h1>

    <h2>一、 验收概览</h2>
    <div class="summary-box">{summary_box}</div>
    <div class="risk-box">{top_risk}</div>

    <h2>二、 核心深度分析</h2>
    <table>
        <tr>
            <th>序号</th>
            <th>模块</th>
            <th>Bug数量</th>
            <th>权重</th>
            <th>模块密度(数量/权重)</th>
        </tr>
        {module_rows_html}
    </table>
    <div class="summary-box">
        Top1模块：<b>{html.escape(top1_module)}</b>；
        动态深度解读：该模块问题主要表现为 <b>{top1_nature}</b>。
    </div>

    <h2>三、 改进措施</h2>
    <div>
        {improvements_html}
    </div>

    <h2>四、 附录：详细缺陷列表</h2>
    <table>
        <tr>
            <th>标题</th>
            <th>任务ID</th>
            <th>执行者</th>
            <th>父任务</th>
            <th>任务状态</th>
            <th>优先级</th>
        </tr>
        {appendix_rows_html}
    </table>
    {ai_section_html}
</body>
</html>
"""
    return html_doc


def save_report(input_file, project_name):
    headers, rows = load_rows(input_file)
    validate_columns(headers)
    rendered = build_html(project_name, rows)

    output_dir = ensure_f_drive_or_desktop()
    timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    output_name = f"质量验收与缺陷分析报告_{timestamp}.doc"
    output_path = output_dir / output_name

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(rendered)
    return output_path


def main():
    from tkinter import Tk, filedialog, messagebox, simpledialog  # pylint: disable=import-outside-toplevel

    root = Tk()
    root.withdraw()
    root.update()

    file_path = filedialog.askopenfilename(
        title="请选择 Bug 数据文件 (CSV/Excel)",
        filetypes=[("Data Files", "*.csv *.xlsx *.xls"), ("All Files", "*.*")],
    )
    if not file_path:
        messagebox.showinfo("提示", "未选择文件，已取消。")
        return

    default_project = Path(file_path).stem
    project_name = simpledialog.askstring("项目名称", "请输入项目名称：", initialvalue=default_project)
    if not project_name:
        project_name = default_project

    try:
        output_path = save_report(file_path, project_name)
        target_tip = "F盘" if str(output_path).lower().startswith("f:") else "桌面(未检测到F盘)"
        messagebox.showinfo("生成成功", f"报告已生成到 {target_tip}\n{output_path}")
        print(f"生成成功: {output_path}")
    except Exception as exc:  # pylint: disable=broad-except
        messagebox.showerror("生成失败", str(exc))
        print(f"生成失败: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
