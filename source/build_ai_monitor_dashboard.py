from __future__ import annotations

import base64
import html
import json
import math
import os
from pathlib import Path
import re

import pandas as pd


ROOT = Path(__file__).resolve().parent


def env_path(name: str, default: Path | str) -> Path:
    raw = os.environ.get(name)
    if raw:
        return Path(raw).expanduser().resolve()
    return Path(default).expanduser().resolve()


OUTPUT = env_path("AI_MONITOR_OUTPUT_HTML", ROOT / "ai_monitor_dashboard_v1.html")
LOGO_PATH = env_path("AI_MONITOR_LOGO_PATH", ROOT / "assets" / "company_logo.jpg")
LOGIC_REFERENCE_PATH = env_path("AI_MONITOR_LOGIC_REFERENCE_PATH", ROOT / "ai_monitor_logic_reference.html")
STANDARD_PATH = env_path(
    "AI_MONITOR_STANDARD_PATH",
    "/Users/lilblackmac/Desktop/AI监课&北极星/海外AI监课质检标准_优化版.xlsx",
)

AI_FILES = {
    "W1": env_path("AI_MONITOR_AI_FILE_W1", "/Users/lilblackmac/Desktop/AI监课&北极星/海外监课指标6月W1.xlsx"),
    "W2": env_path("AI_MONITOR_AI_FILE_W2", "/Users/lilblackmac/Desktop/AI监课&北极星/海外监课指标6月W2.xlsx"),
    "W3": env_path("AI_MONITOR_AI_FILE_W3", "/Users/lilblackmac/Desktop/AI监课&北极星/海外监课指标6月W3.xlsx"),
    "W4": env_path("AI_MONITOR_AI_FILE_W4", "/Users/lilblackmac/Desktop/AI监课&北极星/海外监课指标6月W4.xlsx"),
    "W5": env_path("AI_MONITOR_AI_FILE_W5", "/Users/lilblackmac/Desktop/AI监课&北极星/海外监课指标6月W5.xlsx"),
    "7W1": env_path("AI_MONITOR_AI_FILE_7W1", "/Users/lilblackmac/Desktop/AI监课&北极星/海外监课指标7月W1 .xlsx"),
}
NS_FILES = {
    "W1": env_path("AI_MONITOR_NS_FILE_W1", "/Users/lilblackmac/Desktop/AI监课&北极星/北极星指标-海外团队6月W1（截至0604）.xlsx"),
    "W2": env_path("AI_MONITOR_NS_FILE_W2", "/Users/lilblackmac/Desktop/AI监课&北极星/北极星指标-海外团队6月W2（截至0611）.xlsx"),
    "W3": env_path("AI_MONITOR_NS_FILE_W3", "/Users/lilblackmac/Desktop/AI监课&北极星/北极星指标-海外团队6月W3（截至0618）.xlsx"),
    "W4": env_path("AI_MONITOR_NS_FILE_W4", "/Users/lilblackmac/Desktop/AI监课&北极星/北极星指标-海外团队6月W4（截至0625）.xlsx"),
    "W5": env_path("AI_MONITOR_NS_FILE_W5", "/Users/lilblackmac/Desktop/AI监课&北极星/北极星指标-海外团队6月W5（截至0630）.xlsx"),
}

CURRENT_AI_WEEK = os.environ.get("AI_MONITOR_CURRENT_AI_WEEK", "7W1")
PREVIOUS_AI_WEEK = os.environ.get("AI_MONITOR_PREVIOUS_AI_WEEK", "W5")
CURRENT_AI_SHEET = os.environ.get("AI_MONITOR_CURRENT_AI_SHEET", "W1老师达成情况")
CURRENT_AI_PERIOD = os.environ.get("AI_MONITOR_CURRENT_AI_PERIOD", "W1")
CURRENT_NS_WEEK = os.environ.get("AI_MONITOR_CURRENT_NS_WEEK") or None
PREVIOUS_NS_WEEK = os.environ.get("AI_MONITOR_PREVIOUS_NS_WEEK", "W5")
CURRENT_WEEK_LABEL = os.environ.get("AI_MONITOR_CURRENT_WEEK_LABEL", "7W1")
PREVIOUS_WEEK_LABEL = os.environ.get("AI_MONITOR_PREVIOUS_WEEK_LABEL", "W5")
CURRENT_MONTH_LABEL = os.environ.get("AI_MONITOR_CURRENT_MONTH_LABEL", "202607")
PREVIOUS_MONTH_LABEL = os.environ.get("AI_MONITOR_PREVIOUS_MONTH_LABEL", "202606")
MONTH_PERIOD_START = os.environ.get("AI_MONITOR_MONTH_PERIOD_START", "202601")

FOCUS_DEPARTMENTS = {
    "海外益智英语教学区",
    "海外益智粤语教学区",
    "海外益智台湾教学区",
    "海外益智外教教学区",
}

DEPARTMENT_OVERRIDES = {
    "Dainty老师": "海外益智外教教学区",
}

AI_METRICS = {
    "感染力通过率": {"column": "感染力", "target": None, "format": "pct1", "higher_better": True, "category": "课堂质量核心", "target_label": "感染力合格率"},
    "开口时长": {"column": "开口时长(s)", "target": None, "format": "num1", "higher_better": True, "category": "课堂质量核心", "target_label": "开口时长\n（s）"},
    "课中正确率": {"column": "课中正确率", "target": None, "format": "pct1", "higher_better": True, "category": "课堂质量核心", "target_label": "课中正确率"},
    "上台平均数": {"column": "上台平均数", "target": None, "format": "num1", "higher_better": True, "category": "课堂质量核心", "target_label": "上台次数（BI）"},
    "有效互动频次": {"column": "有效互动次数", "target": None, "format": "num1", "higher_better": True, "category": "课堂质量核心", "target_label": "有效互动频次"},
    "镜头感": {"column": "镜头感", "target": None, "format": "pct1", "higher_better": True, "category": "课堂质量核心", "target_label": None},
    "愉悦度": {"column": "愉悦度", "target": None, "format": "pct1", "higher_better": True, "category": "课堂质量核心", "target_label": None},
    "主动性": {"column": "主动性", "target": None, "format": "num1", "higher_better": True, "category": "教学动作", "target_label": "学员正面动作次数"},
    "课前作业提醒": {"column": "课前作业提醒", "target": None, "format": "pct1", "higher_better": True, "category": "课前课后链路", "target_label": "课前作业提醒合格率"},
    "课后作业布置": {"column": "课后作业布置", "target": None, "format": "pct1", "higher_better": True, "category": "课前课后链路", "target_label": "课后作业布置合格率"},
    "课程预告": {"column": "课程预告", "target": None, "format": "pct1", "higher_better": True, "category": "课前课后链路", "target_label": "课程预告通过率"},
    "读题审题": {"column": "读题审题", "target": None, "format": "pct1", "higher_better": True, "category": "教学动作", "target_label": "读题审题通过率"},
    "关键提问": {"column": "关键提问", "target": None, "format": "pct1", "higher_better": True, "category": "教学动作", "target_label": "关键提问通过率"},
    "情绪策略": {"column": "情绪策略", "target": None, "format": "pct1", "higher_better": True, "category": "教学动作", "target_label": "情绪策略使用通过率"},
    "知识点讲错通过率": {"column": "知识点\n讲错", "target": None, "format": "pct1", "higher_better": True, "category": "教学风险", "target_label": "知识点讲错通过率"},
    "知识点漏讲通过率": {"column": "知识点\n漏讲", "target": None, "format": "pct1", "higher_better": True, "category": "教学风险", "target_label": "知识点漏讲通过率"},
    "课堂规则提及通过率": {"column": "课堂规则\n提及", "target": None, "format": "pct1", "higher_better": True, "category": "基础规范", "target_label": "课堂规则提及通过率"},
    "背景布合格率": {"column": "背景布", "target": None, "format": "pct1", "higher_better": True, "category": "基础规范", "target_label": "背景布\n合格率"},
    "工服合格率": {"column": "工服", "target": None, "format": "pct1", "higher_better": True, "category": "基础规范", "target_label": "工服\n合格率"},
    "光线合格率": {"column": "光线", "target": None, "format": "pct1", "higher_better": True, "category": "基础规范", "target_label": "光线\n合格率"},
    "妆发合格率": {"column": "妆发", "target": None, "format": "pct1", "higher_better": True, "category": "基础规范", "target_label": "妆发\n合格率"},
    "坐姿合格率": {"column": "坐姿", "target": None, "format": "pct1", "higher_better": True, "category": "基础规范", "target_label": "坐姿\n合格率"},
    "行为合格率": {"column": "行为举止", "target": None, "format": "pct1", "higher_better": True, "category": "基础规范", "target_label": "行为\n合格率"},
}

STANDARD_TO_PANEL_METRIC = {
    "背景布": "背景布合格率",
    "工服": "工服合格率",
    "光线镜头": "光线合格率",
    "妆发": "妆发合格率",
    "坐姿": "坐姿合格率",
    "行为举止": "行为合格率",
    "面带微笑亲切温暖": "愉悦度",
    "热情洋溢感染力强": "感染力通过率",
    "知识点讲错": "知识点讲错通过率",
    "知识点漏讲": "知识点漏讲通过率",
    "课堂规则": "课堂规则提及通过率",
    "学习习惯检查": "课前作业提醒",
    "学员开口时长（均值）": "开口时长",
    "课堂正确率（均值）": "课中正确率",
    "学员参与度（均值）": "上台平均数",
    "镜头感（均值）": "镜头感",
    "主动性（均值 ）": "主动性",
    "愉悦度（均值）": "愉悦度",
    "读题审题": "读题审题",
    "关键提问（依据详案）": "关键提问",
    "使用情绪策略（依据详案）": "情绪策略",
    "有效学员互动": "有效互动频次",
    "布置作业": "课后作业布置",
    "下节预告（动画+知识点）": "课程预告",
}

NS_METRICS = {
    "愉悦度": "愉悦度",
    "镜头感": "镜头感",
    "课中首答正确率": "课中首答正确率",
    "课中末答正确率": "课中末答正确率",
    "课后首答正确率": "课后首答正确率",
    "课后末答正确率": "课后末答正确率",
    "作业完成率": "作业完成率",
    "预习率": "预习率",
    "小老师视频提交率": "小老师视频提交率",
}

MANUAL_TARGET_OVERRIDES = {}
DISPLAY_TARGET_OVERRIDES = {
    "镜头感": 0.65,
    "愉悦度": 0.03,
}
MANUAL_TARGET_OVERRIDES = {
    "上台平均数": 6.0,
}

STANDARD_RULES_CACHE: dict[str, dict] = {}

METRICS_WITHOUT_TARGET = {
    "课前作业提醒",
    "课后作业布置",
    "课程预告",
    "读题审题",
    "关键提问",
    "情绪策略",
}

DISPLAY_ONLY_TARGET_METRICS = {
    "镜头感",
    "愉悦度",
}


def build_ns_stage_metric_actual_map(raw: pd.DataFrame) -> dict[str, dict[str, int]]:
    stage_row = raw.iloc[1]
    metric_row = raw.iloc[2]
    value_type_row = raw.iloc[3]
    mapping: dict[str, dict[str, int]] = {}
    current_stage = None
    for col_idx in range(len(raw.columns)):
        stage_value = stage_row.iloc[col_idx]
        if pd.notna(stage_value):
            current_stage = str(stage_value).strip()
        metric_value = metric_row.iloc[col_idx]
        if not current_stage or pd.isna(metric_value):
            continue
        metric_name = str(metric_value).strip()
        value_type = value_type_row.iloc[col_idx]
        value_type_text = "" if pd.isna(value_type) else str(value_type).strip()
        if value_type_text == "实际值":
            mapping.setdefault(current_stage, {})[metric_name] = col_idx
            continue
        if col_idx + 1 < len(raw.columns):
            next_value_type = value_type_row.iloc[col_idx + 1]
            next_value_type_text = "" if pd.isna(next_value_type) else str(next_value_type).strip()
            if next_value_type_text == "实际值":
                mapping.setdefault(current_stage, {})[metric_name] = col_idx + 1
    return mapping

CATEGORY_MAP = {}
for metric_name, config in AI_METRICS.items():
    CATEGORY_MAP.setdefault(config["category"], []).append(metric_name)

CORE_DRIVER_METRICS = [
    "感染力通过率",
    "开口时长",
    "课中正确率",
    "上台平均数",
    "有效互动频次",
    "主动性",
    "背景布合格率",
    "工服合格率",
    "光线合格率",
    "妆发合格率",
    "坐姿合格率",
    "行为合格率",
]

ISSUE_POINTS_MAP = {
    "重度": 5.0,
    "中度": 2.0,
    "轻度": 0.5,
}

CATEGORY_WEIGHT_MAP = {
    "课堂质量核心": 1.5,
    "教学风险": 1.5,
    "教学动作": 1.0,
    "课前课后链路": 0.7,
    "基础规范": 0.5,
}

HIGH_RISK_HARD_GATE_METRICS = {
    "感染力通过率",
    "开口时长",
    "课中正确率",
    "上台平均数",
    "有效互动频次",
    "主动性",
}

RISK_LEVEL_THRESHOLDS_V2 = {
    "high_min": 30.0,
    "mid_min": 12.0,
    "high_quantile": 0.80,
    "mid_quantile": 0.50,
}


def safe_float(value) -> float | None:
    if value is None:
        return None
    if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
        return None
    try:
        numeric = float(value)
    except Exception:
        return value
    if math.isnan(numeric) or math.isinf(numeric):
        return None
    return numeric


def week_file(week: str) -> Path:
    return AI_FILES[week]


def week_teacher_sheet(week: str) -> str:
    if week == CURRENT_AI_WEEK:
        return CURRENT_AI_SHEET
    return f"{week}老师达成情况"


def week_period_label(week: str) -> str:
    if week == CURRENT_AI_WEEK:
        return CURRENT_AI_PERIOD
    return week


def normalize_period_label(value) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, (int, float)) and not math.isnan(float(value)):
        return str(int(value))
    text = str(value).strip()
    if re.fullmatch(r"\d+\.0", text):
        return text.split(".", 1)[0]
    return text


def load_ai_targets() -> dict[str, float | None]:
    df = pd.read_excel(week_file(CURRENT_AI_WEEK), sheet_name=week_teacher_sheet(CURRENT_AI_WEEK), header=None)
    target_row = df.iloc[0].tolist()
    label_row = df.iloc[1].tolist()
    label_to_target = {}
    for target, label in zip(target_row, label_row):
        if pd.isna(label):
            continue
        if isinstance(target, str) and target.strip() == "-":
            label_to_target[str(label)] = None
            continue
        label_to_target[str(label)] = safe_float(target)
    return label_to_target


def load_monthly_summary() -> pd.DataFrame:
    raw = pd.read_excel(week_file(CURRENT_AI_WEEK), sheet_name="月度汇总", header=None)
    section_titles = {"课堂互动", "仪容仪表", "授课状态&教学事故", "课前课中课后"}
    section_rows = [idx for idx in range(len(raw)) if pd.notna(raw.iloc[idx, 1]) and str(raw.iloc[idx, 1]).strip() == "老师属性" and pd.notna(raw.iloc[idx, 2]) and str(raw.iloc[idx, 2]).strip() == "小组"]
    section_rows.append(len(raw))

    combined_records: dict[tuple[str, str], dict[str, object]] = {}
    metric_category_map: dict[str, str] = {}

    for pos, header_row_idx in enumerate(section_rows[:-1]):
        section_title = ""
        for back in range(header_row_idx - 1, -1, -1):
            val = raw.iloc[back, 1]
            if pd.notna(val) and str(val).strip():
                section_title = str(val).strip()
                break
        metric_row = raw.iloc[header_row_idx]
        period_row = raw.iloc[header_row_idx + 1]
        next_header = section_rows[pos + 1]
        current_teacher_attr = ""

        metric_by_col: dict[int, str] = {}
        current_metric = None
        for col_idx in range(3, raw.shape[1]):
            metric_value = metric_row.iloc[col_idx]
            if pd.notna(metric_value):
                current_metric = str(metric_value).strip()
                if current_metric and current_metric != "趋势图":
                    metric_category_map[current_metric] = section_title
            if current_metric:
                metric_by_col[col_idx] = current_metric

        for row_idx in range(header_row_idx + 2, next_header):
            row = raw.iloc[row_idx]
            teacher_attr = "" if pd.isna(row.iloc[1]) else str(row.iloc[1]).strip()
            group_name = "" if pd.isna(row.iloc[2]) else str(row.iloc[2]).strip()
            if teacher_attr:
                current_teacher_attr = teacher_attr
            teacher_attr = current_teacher_attr
            if not teacher_attr and not group_name:
                continue
            if teacher_attr in section_titles and not group_name:
                continue
            key = (teacher_attr, group_name)
            record = combined_records.setdefault(key, {"老师属性": teacher_attr, "小组": group_name})
            for col_idx in range(3, raw.shape[1]):
                metric_name = metric_by_col.get(col_idx)
                if not metric_name or metric_name == "趋势图":
                    continue
                period_value = period_row.iloc[col_idx]
                if pd.isna(period_value):
                    continue
                period_text = normalize_period_label(period_value)
                if not re.fullmatch(r"\d{6}", period_text):
                    continue
                if period_text < MONTH_PERIOD_START or period_text > CURRENT_MONTH_LABEL:
                    continue
                record[f"{metric_name}_{period_text}"] = safe_float(pd.to_numeric(row.iloc[col_idx], errors="coerce"))

    monthly = pd.DataFrame(combined_records.values())
    month_periods = sorted({
        col.rsplit("_", 1)[1]
        for col in monthly.columns
        if re.search(r"_\d{6}$", col)
    })
    metric_names = sorted({
        col.rsplit("_", 1)[0]
        for col in monthly.columns
        if re.search(r"_\d{6}$", col)
    })
    for metric in metric_names:
        monthly[f"{metric}_变化"] = monthly.get(f"{metric}_{CURRENT_MONTH_LABEL}") - monthly.get(f"{metric}_{PREVIOUS_MONTH_LABEL}")
        monthly[f"{metric}_category"] = metric_category_map.get(metric)
    monthly.attrs["periods"] = month_periods
    return monthly


def load_team_monthly_total() -> dict[str, object]:
    raw = pd.read_excel(week_file(CURRENT_AI_WEEK), sheet_name="月度汇总", header=None)
    section_rows = [
        idx for idx in range(len(raw))
        if pd.notna(raw.iloc[idx, 1])
        and str(raw.iloc[idx, 1]).strip() == "老师属性"
        and pd.notna(raw.iloc[idx, 2])
        and str(raw.iloc[idx, 2]).strip() == "小组"
    ]
    section_rows.append(len(raw))

    result: dict[str, object] = {"老师属性": "总计", "小组": ""}
    periods: set[str] = set()
    metric_category_map: dict[str, str] = {}

    for pos, header_row_idx in enumerate(section_rows[:-1]):
        section_title = ""
        for back in range(header_row_idx - 1, -1, -1):
            val = raw.iloc[back, 1]
            if pd.notna(val) and str(val).strip():
                section_title = str(val).strip()
                break

        metric_row = raw.iloc[header_row_idx]
        period_row = raw.iloc[header_row_idx + 1]
        next_header_idx = section_rows[pos + 1]
        total_row_idx = len(raw) - 1 if next_header_idx == len(raw) else next_header_idx - 3
        total_row = raw.iloc[total_row_idx]

        current_metric = None
        for col_idx in range(3, raw.shape[1]):
            metric_value = metric_row.iloc[col_idx]
            if pd.notna(metric_value):
                current_metric = str(metric_value).strip()
                if current_metric and current_metric != "趋势图":
                    metric_category_map[current_metric] = section_title
            if not current_metric or current_metric == "趋势图":
                continue
            period_text = normalize_period_label(period_row.iloc[col_idx])
            if not re.fullmatch(r"\d{6}", period_text):
                continue
            if period_text < MONTH_PERIOD_START or period_text > CURRENT_MONTH_LABEL:
                continue
            periods.add(period_text)
            result[f"{current_metric}_{period_text}"] = safe_float(pd.to_numeric(total_row.iloc[col_idx], errors="coerce"))

    result["periods"] = sorted(periods)
    result["metric_category_map"] = metric_category_map
    return result


def load_standard_rules() -> dict[str, dict]:
    df = pd.read_excel(STANDARD_PATH)
    df["一级分类"] = df["一级分类"].ffill()
    df["二级分类"] = df["二级分类"].ffill()
    rules = {}
    for _, row in df.iterrows():
        indicator = str(row["指标"]).strip()
        panel_metric = STANDARD_TO_PANEL_METRIC.get(indicator)
        if not panel_metric:
            continue
        rules[panel_metric] = {
            "source_indicator": indicator,
            "level1": None if pd.isna(row["一级分类"]) else str(row["一级分类"]).strip(),
            "level2": None if pd.isna(row["二级分类"]) else str(row["二级分类"]).strip(),
            "current_standard": None if pd.isna(row["当前质检标准"]) else str(row["当前质检标准"]).strip(),
            "detail": None if pd.isna(row["质检细节"]) else str(row["质检细节"]).strip(),
            "adjusted_standard": None if pd.isna(row["质检标准调整"]) else str(row["质检标准调整"]).strip(),
        }
    return rules


def extract_a_target_from_adjusted_standard(text: str) -> float | None:
    if not text:
        return None
    compact = text.replace("建议调整为：", "").replace(" ", "")
    match = re.search(r"A：([^ABCD]+?)(?=([BCD])：|$)", compact)
    if not match:
        if "通过：" in compact and "不通过" in compact:
            return 0.5
        return None
    clause = match.group(1)
    ge_match = re.search(r"≥([0-9.]+分钟|[0-9.]+次|[0-9.]+秒|[0-9.]+%)", clause)
    if ge_match:
        return coerce_rule_number(ge_match.group(1))
    exact_match = re.search(r"=([0-9.]+分钟|[0-9.]+次|[0-9.]+秒|[0-9.]+%)", clause)
    if exact_match:
        return coerce_rule_number(exact_match.group(1))
    return None


def parse_numeric_band_rule(text: str) -> dict | None:
    if not text or "A：" not in text:
        return None
    compact = text.replace("建议调整为：", "").replace(" ", "")
    pattern = re.compile(r"([ABCD])：([^ABCD]+?)(?=([ABCD])：|$)")
    parsed = []
    for match in pattern.finditer(compact):
        level = match.group(1)
        clause = match.group(2)
        unit_hint = None
        if "分钟" in clause:
            unit_hint = "分钟"
        elif "%" in clause:
            unit_hint = "%"
        elif "次" in clause:
            unit_hint = "次"
        elif "秒" in clause:
            unit_hint = "秒"
        normalized = (
            clause.replace("均值", "")
            .replace("平均正确率", "")
            .replace("平均开口时长", "")
            .replace("平均次数", "")
            .replace("占比", "")
            .replace("互动学员", "")
            .replace("每节课所有学员课中习题", "")
            .replace("每节课学员", "")
            .replace("学员有效正脸帧", "")
            .replace("学员正面微笑表情", "")
            .replace("学员正面动作", "")
            .replace("有效引导频次", "")
            .replace("频次", "")
        )
        ge_match = re.search(r"≥([0-9.]+%?|[0-9.]+分钟|[0-9.]+次|[0-9.]+秒)", normalized)
        range_match = re.search(r"([0-9.]+%?|[0-9.]+分钟|[0-9.]+次|[0-9.]+秒)≤.*?<([0-9.]+%?|[0-9.]+分钟|[0-9.]+次|[0-9.]+秒)", normalized)
        lt_match = re.search(r"<([0-9.]+%?|[0-9.]+分钟|[0-9.]+次|[0-9.]+秒)", normalized)
        le_upper_match = re.search(r">([0-9.]+%?).*?≤([0-9.]+%?)", normalized)
        def apply_unit_hint(v: str | None) -> str | None:
            if v is None or unit_hint is None:
                return v
            if any(token in v for token in ["%", "分钟", "次", "秒"]):
                return v
            return f"{v}{unit_hint}"
        if ge_match and "≤" not in normalized and "<" not in normalized:
            parsed.append({"level": level, "min": apply_unit_hint(ge_match.group(1)), "max": None, "min_inclusive": True, "max_inclusive": False})
        elif range_match:
            parsed.append({"level": level, "min": apply_unit_hint(range_match.group(1)), "max": apply_unit_hint(range_match.group(2)), "min_inclusive": True, "max_inclusive": False})
        elif le_upper_match:
            parsed.append({"level": level, "min": apply_unit_hint(le_upper_match.group(1)), "max": apply_unit_hint(le_upper_match.group(2)), "min_inclusive": False, "max_inclusive": True})
        elif lt_match:
            parsed.append({"level": level, "min": None, "max": apply_unit_hint(lt_match.group(1)), "min_inclusive": False, "max_inclusive": False})
    if not parsed:
        return None
    return {"type": "band", "bands": parsed}


def coerce_rule_number(value: str) -> float | None:
    if value is None:
        return None
    v = value.replace("%", "").replace("分钟", "").replace("次", "").replace("秒", "").strip()
    try:
        num = float(v)
    except Exception:
        return None
    if "%" in value:
        return num / 100
    if "分钟" in value:
        return num * 60
    return num


def enrich_standard_rules(rules: dict[str, dict]) -> dict[str, dict]:
    for metric, rule in rules.items():
        adjusted = rule.get("adjusted_standard")
        if not adjusted:
            continue
        rule["a_target"] = extract_a_target_from_adjusted_standard(adjusted)
        if "通过：" in adjusted and "不通过" in adjusted:
            rule["status_rule"] = {"type": "pass_fail"}
        else:
            band_rule = parse_numeric_band_rule(adjusted)
            if band_rule:
                for band in band_rule["bands"]:
                    band["min"] = coerce_rule_number(band["min"])
                    band["max"] = coerce_rule_number(band["max"])
                rule["status_rule"] = band_rule
    return rules


def format_value(value: float | None, fmt: str) -> str:
    if value is None:
        return "-"
    if fmt == "pct1":
        return f"{value * 100:.1f}%"
    if fmt == "pct0":
        return f"{value * 100:.0f}%"
    if fmt == "num0":
        return f"{value:.0f}"
    return f"{value:.1f}"


def compute_gap(metric_name: str, value: float | None) -> tuple[float | None, float]:
    config = AI_METRICS[metric_name]
    target = config["target"]
    if value is None or target is None:
        return None, 0.0
    gap = value - target if config["higher_better"] else target - value
    if gap >= 0:
        severity = 0.0
    elif config["format"].startswith("pct"):
        # 百分类指标用“绝对百分点差距”计风险，避免目标值较小时被比例放大。
        severity = min(abs(gap), 1.0)
    else:
        severity = min(abs(gap) / target, 1.0) if target else min(abs(gap), 1.0)
    return gap, severity


def bucket_from_status_rule(metric_name: str, value: float | None) -> tuple[str | None, int, float]:
    if value is None:
        return None, 0, 0.0
    rule = STANDARD_RULES_CACHE.get(metric_name, {}).get("status_rule")
    if not rule:
        return None, 0, 0.0
    if rule.get("type") == "pass_fail":
        return (None, 0, 0.0) if value >= 0.5 else ("中度", 2, 0.5)
    if rule.get("type") != "band":
        return None, 0, 0.0
    for band in rule.get("bands", []):
        min_ok = band["min"] is None or value >= band["min"]
        max_ok = band["max"] is None or value < band["max"]
        if min_ok and max_ok:
            level = band.get("level")
            if level == "A":
                return None, 0, 0.0
            if level == "B":
                return "轻度", 1, 1.0
            if level == "C":
                return "中度", 2, 2.0
            if level == "D":
                return "重度", 3, 3.0
    return None, 0, 0.0


def issue_bucket(metric_name: str, current_value: float | None, previous_value: float | None) -> tuple[str | None, int, float]:
    config = AI_METRICS[metric_name]
    target = config["target"]
    if current_value is None:
        return None, 0, 0.0

    if metric_name in DISPLAY_ONLY_TARGET_METRICS:
        return None, 0, 0.0

    if target is None:
        return None, 0, 0.0

    rule_bucket, rule_points, rule_score = bucket_from_status_rule(metric_name, current_value)
    if rule_points > 0 or rule_bucket is None and STANDARD_RULES_CACHE.get(metric_name, {}).get("status_rule"):
        return rule_bucket, rule_points, rule_score
    gap, severity = compute_gap(metric_name, current_value)
    if gap is None or gap >= 0:
        return None, 0, 0.0
    if config["format"].startswith("pct"):
        shortfall = abs(gap)
        if shortfall >= 0.10:
            return "重度", 3, severity
        if shortfall >= 0.05:
            return "中度", 2, severity
        return "轻度", 1, severity
    shortfall_ratio = severity
    if shortfall_ratio >= 0.30:
        return "重度", 3, severity
    if shortfall_ratio >= 0.15:
        return "中度", 2, severity
    return "轻度", 1, severity


def issue_bucket_counts(row: pd.Series, week_suffix: str = "W2", compare_suffix: str = "W1") -> dict[str, int]:
    counts = {"重度": 0, "中度": 0, "轻度": 0}
    for metric_name in CORE_DRIVER_METRICS:
        bucket, _, _ = issue_bucket(metric_name, row.get(f"{metric_name}_{week_suffix}"), row.get(f"{metric_name}_{compare_suffix}"))
        if bucket in counts:
            counts[bucket] += 1
    return counts


def risk_score_components(row: pd.Series, week_suffix: str = "W2", compare_suffix: str = "W1") -> tuple[float, dict[str, int]]:
    total = 0.0
    counts = {"重度": 0, "中度": 0, "轻度": 0}
    for metric_name in CORE_DRIVER_METRICS:
        bucket, _, _ = issue_bucket(metric_name, row.get(f"{metric_name}_{week_suffix}"), row.get(f"{metric_name}_{compare_suffix}"))
        if not bucket:
            continue
        counts[bucket] += 1
        category = AI_METRICS[metric_name]["category"]
        weight = CATEGORY_WEIGHT_MAP.get(category, 1.0)
        total += ISSUE_POINTS_MAP[bucket] * weight
    return round(total, 1), counts


def risk_hard_gate_flags(row: pd.Series, week_suffix: str = "W2", compare_suffix: str = "W1") -> dict[str, int]:
    core_severe = 0
    core_medium = 0
    teaching_risk_severe = 0
    for metric_name in CORE_DRIVER_METRICS:
        bucket, _, _ = issue_bucket(metric_name, row.get(f"{metric_name}_{week_suffix}"), row.get(f"{metric_name}_{compare_suffix}"))
        if metric_name in HIGH_RISK_HARD_GATE_METRICS:
            if bucket == "重度":
                core_severe += 1
            elif bucket == "中度":
                core_medium += 1
        if AI_METRICS[metric_name]["category"] == "教学风险" and bucket == "重度":
            teaching_risk_severe += 1
    return {
        "核心重度问题数": core_severe,
        "核心中度问题数": core_medium,
        "教学风险重度问题数": teaching_risk_severe,
    }


def driver_snapshot(metric_name: str, current_value: float | None, previous_value: float | None) -> dict:
    gap, severity = compute_gap(metric_name, current_value)
    wow = None if current_value is None or previous_value is None else current_value - previous_value
    bucket, points, order_score = issue_bucket(metric_name, current_value, previous_value)
    order_score = safe_float(order_score) or 0.0
    return {
        "metric": metric_name,
        "current": safe_float(current_value),
        "previous": safe_float(previous_value),
        "target": AI_METRICS[metric_name]["target"],
        "gap": safe_float(gap),
        "wow": safe_float(wow),
        "severity": round(order_score, 4),
        "bucket": bucket,
        "points": safe_float(ISSUE_POINTS_MAP.get(bucket, 0.0) * CATEGORY_WEIGHT_MAP.get(AI_METRICS[metric_name]["category"], 1.0) if bucket else 0.0),
        "category": AI_METRICS[metric_name]["category"],
        "format": AI_METRICS[metric_name]["format"],
        "reason": "目标差距" if AI_METRICS[metric_name]["target"] is not None else "周环比走弱",
    }


def load_ai_week(week: str) -> pd.DataFrame:
    summary = pd.read_excel(week_file(week), sheet_name=week_teacher_sheet(week), header=1)
    summary = summary[summary["昵称"].notna()].copy()
    summary = summary[summary["昵称"].astype(str).str.strip() != ""].copy()
    for config in AI_METRICS.values():
        if config["target_label"] and config["target_label"] in summary.columns:
            summary[config["target_label"]] = pd.to_numeric(summary[config["target_label"]], errors="coerce")

    detail = pd.read_excel(week_file(week), sheet_name="质检明细")
    if "周期" in detail.columns:
        detail = detail[detail["周期"].astype(str).str.strip() == week_period_label(week)].copy()
    detail = detail[detail["教师艺名"].notna()].copy()
    sample_counts = (
        detail.groupby(["教师艺名", "教师部门5", "教师部门6"], as_index=False)["id"]
        .count()
        .rename(columns={"教师艺名": "老师", "教师部门5": "业务部门", "教师部门6": "小组", "id": "样本课节数"})
    )

    teacher = pd.DataFrame({
        "老师": summary["昵称"].astype(str).str.strip(),
        "老师工号": summary["id"].astype(str).str.strip(),
        "业务部门": summary["大组"].astype(str).str.strip(),
        "小组": summary["小组"].astype(str).str.strip(),
    })
    teacher["业务部门"] = teacher.apply(
        lambda row: DEPARTMENT_OVERRIDES.get(row["老师"], row["业务部门"]),
        axis=1,
    )
    for metric_name, config in AI_METRICS.items():
        source_col = config["target_label"]
        teacher[metric_name] = summary[source_col] if source_col in summary.columns else None

    teacher = teacher.merge(sample_counts, on=["老师", "业务部门", "小组"], how="left")
    teacher["样本课节数"] = teacher["样本课节数"].fillna(0)
    teacher["周次"] = CURRENT_WEEK_LABEL if week == CURRENT_AI_WEEK else week
    return teacher


def load_ns_week(week: str) -> pd.DataFrame:
    if not week or week not in NS_FILES or not NS_FILES[week].exists():
        return pd.DataFrame(columns=["老师", "业务部门", "小组", *NS_METRICS.keys()])
    raw = pd.read_excel(NS_FILES[week], sheet_name="老师达成情况", header=None)
    stage_metric_actual_map = build_ns_stage_metric_actual_map(raw)

    df = pd.read_excel(NS_FILES[week], sheet_name="老师达成情况", header=2)
    df = df[df["老师艺名"].notna()].copy()
    df = df[df["老师艺名"].astype(str).str.strip() != ""].copy()
    df = df[df["老师艺名"].astype(str).str.strip() != "老师艺名"].copy()

    records = []
    for _, row in df.iterrows():
        stage = None if pd.isna(row["课堂主学段"]) else str(row["课堂主学段"]).strip()
        mapping = stage_metric_actual_map.get(stage, {})
        record = {
            "老师": None if pd.isna(row["老师艺名"]) else str(row["老师艺名"]).strip(),
            "业务部门": None if pd.isna(row["教学老师五级部门"]) else str(row["教学老师五级部门"]).strip(),
            "小组": None if pd.isna(row["教学老师六级部门"]) else str(row["教学老师六级部门"]).strip(),
        }
        for metric in NS_METRICS:
            col_idx = mapping.get(metric)
            value = None if col_idx is None else safe_float(pd.to_numeric(row.iloc[col_idx], errors="coerce"))
            record[metric] = value
        records.append(record)

    return pd.DataFrame(records)


def collapse_duplicate_teachers(teacher: pd.DataFrame) -> pd.DataFrame:
    metric_columns = [
        col for col in teacher.columns
        if col not in {"老师", "老师工号", "业务部门", "小组"}
    ]

    records = []
    for name, sub in teacher.groupby("老师", dropna=False, sort=False):
        if len(sub) == 1:
            records.append(sub.iloc[0].to_dict())
            continue

        # Prefer the row with current-week samples; fall back to the row with more previous samples.
        ranked = sub.assign(
            _w2=sub["样本课节数_W2"].fillna(0),
            _w1=sub["样本课节数_W1"].fillna(0),
        ).sort_values(["_w2", "_w1"], ascending=[False, False], kind="stable")
        base = ranked.iloc[0].to_dict()

        for col in metric_columns:
            if col in {"样本课节数_W2", "样本课节数_W1"}:
                base[col] = safe_float(sub[col].fillna(0).sum())
                continue
            if col.endswith("_W2"):
                values = sub[col].dropna()
                base[col] = values.iloc[0] if not values.empty else None
                continue
            if col.endswith("_W1"):
                values = sub[col].dropna()
                base[col] = values.iloc[0] if not values.empty else None
                continue
            values = sub[col].dropna()
            base[col] = values.iloc[0] if not values.empty else base.get(col)

        teacher_ids = [str(v).strip() for v in sub["老师工号"].tolist() if str(v).strip()]
        base["老师工号"] = teacher_ids[0] if teacher_ids else ""
        records.append(base)

    return pd.DataFrame(records)


def merge_weeks(current: pd.DataFrame, previous: pd.DataFrame, ns_current: pd.DataFrame, ns_previous: pd.DataFrame) -> pd.DataFrame:
    keys = ["老师", "业务部门", "小组"]
    previous = previous.rename(columns={col: f"{col}_W1" for col in previous.columns if col not in keys})
    current = current.rename(columns={col: f"{col}_W2" for col in current.columns if col not in keys})
    teacher = current.merge(previous, on=keys, how="outer")
    teacher["老师"] = teacher["老师"].fillna("")
    teacher["业务部门"] = teacher["业务部门"].fillna("未分组")
    teacher["小组"] = teacher["小组"].fillna("未分组")
    teacher["老师工号_W2"] = teacher.get("老师工号_W2", "").fillna("")
    teacher["老师工号_W1"] = teacher.get("老师工号_W1", "").fillna("")
    teacher["老师工号"] = teacher["老师工号_W2"].where(teacher["老师工号_W2"] != "", teacher["老师工号_W1"])
    teacher = teacher.drop(columns=["老师工号_W2", "老师工号_W1"])

    ns_current = ns_current.rename(columns={col: f"{col}_W2_NS" for col in ns_current.columns if col not in ["老师", "业务部门", "小组"]})
    ns_previous = ns_previous.rename(columns={col: f"{col}_W1_NS" for col in ns_previous.columns if col not in ["老师", "业务部门", "小组"]})
    teacher = teacher.merge(ns_current, on=["老师", "业务部门", "小组"], how="left")
    teacher = teacher.merge(ns_previous, on=["老师", "业务部门", "小组"], how="left")

    for metric_name in ["镜头感", "愉悦度"]:
        teacher[f"{metric_name}_W2"] = teacher[f"{metric_name}_W2"].combine_first(teacher[f"{metric_name}_W2_NS"])
        teacher[f"{metric_name}_W1"] = teacher[f"{metric_name}_W1"].combine_first(teacher[f"{metric_name}_W1_NS"])
        teacher = teacher.drop(columns=[f"{metric_name}_W2_NS", f"{metric_name}_W1_NS"])

    for metric_name in ["课中首答正确率", "课中末答正确率", "课后首答正确率", "课后末答正确率", "作业完成率", "预习率", "小老师视频提交率"]:
        teacher[f"{metric_name}_W2"] = teacher[f"{metric_name}_W2_NS"]
        teacher[f"{metric_name}_W1"] = teacher[f"{metric_name}_W1_NS"]
        teacher = teacher.drop(columns=[f"{metric_name}_W2_NS", f"{metric_name}_W1_NS"])

    for metric_name in AI_METRICS:
        teacher[f"{metric_name}_变化"] = teacher[f"{metric_name}_W2"] - teacher[f"{metric_name}_W1"]
    teacher["样本课节数_W2"] = teacher["样本课节数_W2"].fillna(0)
    teacher["样本课节数_W1"] = teacher["样本课节数_W1"].fillna(0)
    return collapse_duplicate_teachers(teacher)


def extract_top_drivers(row: pd.Series, limit: int = 3) -> list[dict]:
    issues = []
    for metric_name in CORE_DRIVER_METRICS:
        snapshot = driver_snapshot(metric_name, row.get(f"{metric_name}_W2"), row.get(f"{metric_name}_W1"))
        if snapshot["points"] > 0:
            issues.append(snapshot)
    issues.sort(key=lambda item: (item["points"], item["severity"]), reverse=True)
    return issues[:limit]


def split_top_drivers_by_category(issues: list[dict], limit: int = 3) -> dict[str, list[dict]]:
    categorized: dict[str, list[dict]] = {}
    for category in CATEGORY_MAP:
        items = [issue for issue in issues if issue.get("category") == category]
        items.sort(key=lambda item: (item["points"], item["severity"]), reverse=True)
        categorized[category] = items[:limit]
    return categorized


def risk_score(row: pd.Series) -> float:
    score, _ = risk_score_components(row, "W2", "W1")
    return score


def risk_level(score: float) -> str:
    if score >= RISK_LEVEL_THRESHOLDS_V2["high_min"]:
        return "高"
    if score >= RISK_LEVEL_THRESHOLDS_V2["mid_min"]:
        return "中"
    return "低"


def risk_level_v2(row: pd.Series, high_threshold: float, mid_threshold: float) -> str:
    flags = row.get("风险硬门槛") or {}
    core_severe = flags.get("核心重度问题数", 0)
    core_medium = flags.get("核心中度问题数", 0)
    score = row.get("风险分") or 0
    if (
        core_severe >= 2
        or (core_severe >= 1 and core_medium >= 3)
        or score >= high_threshold
    ):
        return "高"
    if score >= mid_threshold:
        return "中"
    return "低"


def baseline_risk_score(row: pd.Series) -> float:
    score, _ = risk_score_components(row, "W1", "W2")
    return score


def prepare_teacher_records(teacher: pd.DataFrame) -> pd.DataFrame:
    teacher = teacher.copy()
    teacher["风险问题分档统计"] = teacher.apply(lambda row: issue_bucket_counts(row, "W2", "W1"), axis=1)
    teacher["风险分"] = teacher.apply(risk_score, axis=1)
    teacher["风险硬门槛"] = teacher.apply(lambda row: risk_hard_gate_flags(row, "W2", "W1"), axis=1)
    high_threshold = max(
        RISK_LEVEL_THRESHOLDS_V2["high_min"],
        safe_float(teacher["风险分"].quantile(RISK_LEVEL_THRESHOLDS_V2["high_quantile"])) or 0.0,
    )
    mid_threshold = max(
        RISK_LEVEL_THRESHOLDS_V2["mid_min"],
        safe_float(teacher["风险分"].quantile(RISK_LEVEL_THRESHOLDS_V2["mid_quantile"])) or 0.0,
    )
    teacher["风险等级"] = teacher.apply(
        lambda row: risk_level_v2(
            row,
            high_threshold,
            mid_threshold,
        ),
        axis=1,
    )
    teacher["风险等级阈值"] = teacher.apply(
        lambda _: {"high": round(high_threshold, 1), "mid": round(mid_threshold, 1)},
        axis=1,
    )
    teacher["重点问题明细"] = teacher.apply(extract_top_drivers, axis=1)
    teacher["分类重点问题明细"] = teacher["重点问题明细"].apply(split_top_drivers_by_category)
    teacher["重点问题"] = teacher["重点问题明细"].apply(lambda items: " / ".join(item["metric"] for item in items) if items else "整体稳定")
    teacher["风险变化"] = teacher["风险分"] - teacher.apply(baseline_risk_score, axis=1)
    return teacher


def aggregate_level(teacher: pd.DataFrame, level_keys: list[str]) -> pd.DataFrame:
    agg_dict = {
        "老师": "count",
        "风险分": "mean",
        "样本课节数_W2": "sum",
        "样本课节数_W1": "sum",
    }
    for metric_name in AI_METRICS:
        agg_dict[f"{metric_name}_W2"] = "mean"
        agg_dict[f"{metric_name}_W1"] = "mean"
    for metric_name in ["作业完成率", "预习率", "小老师视频提交率", "课中首答正确率", "课中末答正确率", "课后首答正确率", "课后末答正确率"]:
        agg_dict[f"{metric_name}_W2"] = "mean"
        agg_dict[f"{metric_name}_W1"] = "mean"
    result = teacher.groupby(level_keys, as_index=False).agg(agg_dict).rename(columns={"老师": "老师数"})
    risk_group = teacher.groupby(level_keys)
    result["高风险老师数"] = risk_group["风险等级"].apply(lambda s: int((s == "高").sum())).values
    result["中高风险老师数"] = risk_group["风险等级"].apply(lambda s: int(s.isin(["高", "中"]).sum())).values
    result["高风险占比"] = result["高风险老师数"] / result["老师数"]
    result["风险变化"] = risk_group["风险变化"].mean().values
    result["覆盖老师数_W2"] = risk_group["样本课节数_W2"].apply(lambda s: int((s > 0).sum())).values
    result["覆盖老师数_W1"] = risk_group["样本课节数_W1"].apply(lambda s: int((s > 0).sum())).values
    return result


def attach_top_drivers(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    top_driver_records = []
    for _, row in df.iterrows():
        issues = []
        for metric_name in CORE_DRIVER_METRICS:
            issues.append(driver_snapshot(metric_name, row.get(f"{metric_name}_W2"), row.get(f"{metric_name}_W1")))
        issues = [item for item in issues if item["points"] > 0]
        issues.sort(key=lambda item: (item["points"], item["severity"]), reverse=True)
        issues = issues[:3]
        top_driver_records.append(issues)
    df["重点抓手明细"] = top_driver_records
    df["分类重点抓手明细"] = df["重点抓手明细"].apply(split_top_drivers_by_category)
    df["重点抓手"] = df["重点抓手明细"].apply(lambda items: " / ".join(item["metric"] for item in items) if items else "整体稳定")
    return df


def attainment_summary(row: pd.Series, suffix: str = "W2") -> tuple[int, int, float | None]:
    met = 0
    total = 0
    for metric_name, config in AI_METRICS.items():
        target = config.get("target")
        if target is None:
            continue
        value = row.get(f"{metric_name}_{suffix}")
        if value is None or pd.isna(value):
            continue
        total += 1
        if value >= target:
            met += 1
    rate = (met / total) if total else None
    return met, total, rate


def trend_label(delta: float | None) -> str:
    if delta is None or pd.isna(delta):
        return "暂无趋势"
    if delta >= 0.03:
        return "明显改善"
    if delta >= 0.005:
        return "小幅改善"
    if delta <= -0.03:
        return "明显回落"
    if delta <= -0.005:
        return "小幅回落"
    return "基本持平"


def add_attainment_fields(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    met_w2 = []
    total_w2 = []
    rate_w2 = []
    met_w1 = []
    total_w1 = []
    rate_w1 = []
    for _, row in df.iterrows():
        m2, t2, r2 = attainment_summary(row, "W2")
        m1, t1, r1 = attainment_summary(row, "W1")
        met_w2.append(m2)
        total_w2.append(t2)
        rate_w2.append(r2)
        met_w1.append(m1)
        total_w1.append(t1)
        rate_w1.append(r1)
    df["达标指标数_W2"] = met_w2
    df["应评指标数_W2"] = total_w2
    df["指标达成率_W2"] = rate_w2
    df["达标指标数_W1"] = met_w1
    df["应评指标数_W1"] = total_w1
    df["指标达成率_W1"] = rate_w1
    df["指标达成率变化"] = df["指标达成率_W2"] - df["指标达成率_W1"]
    df["指标达成趋势"] = df["指标达成率变化"].apply(trend_label)
    return df


def clean_records(df: pd.DataFrame) -> list[dict]:
    records = []
    for record in df.to_dict(orient="records"):
        cleaned = {}
        for key, value in record.items():
            if isinstance(value, list):
                if value and isinstance(value[0], dict):
                    cleaned[key] = [{sub_k: safe_float(sub_v) for sub_k, sub_v in item.items()} for item in value]
                else:
                    cleaned[key] = [safe_float(item) for item in value]
            elif isinstance(value, dict):
                cleaned[key] = {sub_k: safe_float(sub_v) for sub_k, sub_v in value.items()}
            else:
                cleaned[key] = safe_float(value)
        records.append(cleaned)
    return records


def natural_group_sort_key(name: str) -> tuple:
    if not isinstance(name, str):
        return ("", 0, "")
    match = re.match(r"^(.*?)(\d+)组$", name)
    if match:
        return (match.group(1), int(match.group(2)), name)
    return (name, 0, name)


def build_payload() -> dict:
    global STANDARD_RULES_CACHE
    ai_targets = load_ai_targets()
    standard_rules = enrich_standard_rules(load_standard_rules())
    STANDARD_RULES_CACHE = standard_rules
    for metric_name, config in AI_METRICS.items():
        if metric_name in METRICS_WITHOUT_TARGET:
            config["target"] = None
            continue
        if metric_name in MANUAL_TARGET_OVERRIDES:
            config["target"] = MANUAL_TARGET_OVERRIDES[metric_name]
            continue
        if metric_name in DISPLAY_TARGET_OVERRIDES:
            config["target"] = DISPLAY_TARGET_OVERRIDES[metric_name]
            continue
        standard_target = standard_rules.get(metric_name, {}).get("a_target")
        if standard_target is not None:
            config["target"] = standard_target
        elif config["target_label"]:
            config["target"] = ai_targets.get(config["target_label"])
        else:
            config["target"] = None

    ai_w1 = load_ai_week(PREVIOUS_AI_WEEK)
    ai_w2 = load_ai_week(CURRENT_AI_WEEK)
    ns_w1 = load_ns_week(PREVIOUS_NS_WEEK)
    ns_w2 = load_ns_week(CURRENT_NS_WEEK)
    monthly = load_monthly_summary()
    team_monthly_total = load_team_monthly_total()

    teacher = merge_weeks(ai_w2, ai_w1, ns_w2, ns_w1)
    teacher = prepare_teacher_records(teacher)
    teacher = add_attainment_fields(teacher)

    group = add_attainment_fields(attach_top_drivers(aggregate_level(teacher, ["业务部门", "小组"])))
    department_all = add_attainment_fields(attach_top_drivers(aggregate_level(teacher, ["业务部门"])))
    department = department_all[department_all["业务部门"].isin(FOCUS_DEPARTMENTS)].copy()
    focus_department = department.copy()
    if focus_department.empty:
        focus_department = department_all.copy()

    overall = add_attainment_fields(pd.DataFrame([teacher.mean(numeric_only=True)]).pipe(attach_top_drivers)).iloc[0]
    summary = {
        "业务部门数": int(department["业务部门"].nunique()),
        "小组数": int(group["小组"].nunique()),
        "老师数": int(teacher["老师"].nunique()),
        f"{CURRENT_WEEK_LABEL}监课量": int(teacher["样本课节数_W2"].sum()),
        f"{PREVIOUS_WEEK_LABEL}监课量": int(teacher["样本课节数_W1"].sum()),
        f"{CURRENT_WEEK_LABEL}覆盖老师数": int((teacher["样本课节数_W2"] > 0).sum()),
        f"{PREVIOUS_WEEK_LABEL}覆盖老师数": int((teacher["样本课节数_W1"] > 0).sum()),
        "高风险老师数": int((teacher["风险等级"] == "高").sum()),
        "高风险占比": safe_float((teacher["风险等级"] == "高").mean()),
        "风险变化": safe_float(teacher["风险变化"].mean()),
        "风险最集中业务部门": focus_department.sort_values("高风险占比", ascending=False).iloc[0]["业务部门"],
        "风险最集中小组": group.sort_values("高风险占比", ascending=False).iloc[0]["小组"],
        "TOP问题抓手": overall["重点抓手"],
        "分类TOP问题抓手明细": overall["分类重点抓手明细"],
        "新增高风险老师数": int(((teacher["风险等级"] == "高") & (teacher["风险变化"] > 0)).sum()),
        "已改善老师数": int((teacher["风险变化"] < 0).sum()),
        "持续高风险老师数": int(((teacher["风险等级"] == "高") & (teacher["风险变化"] >= 0)).sum()),
        "本周重点关注范围": focus_department.sort_values("高风险占比", ascending=False).iloc[0]["业务部门"],
        "月度重点波动指标": "感染力合格率",
        "整体指标达成率": safe_float(overall["指标达成率_W2"]),
        "整体指标达成率变化": safe_float(overall["指标达成率变化"]),
        "整体达成趋势": overall["指标达成趋势"],
    }

    teacher = teacher.sort_values(["风险分", "样本课节数_W2"], ascending=[False, False])
    group = group.sort_values(["业务部门", "小组"], key=lambda s: s.map(natural_group_sort_key) if s.name == "小组" else s)
    group = group.sort_values(["高风险占比", "风险分"], ascending=[False, False], kind="stable")
    department = department.sort_values(["高风险占比", "风险分"], ascending=[False, False])

    monthly_periods = monthly.attrs.get("periods", [])

    return {
        "summary": summary,
        "departments": clean_records(department),
        "groups": clean_records(group),
        "teachers": clean_records(teacher),
        "categories": CATEGORY_MAP,
        "metric_meta": {metric: {"format": config["format"], "target": safe_float(config["target"]), "category": config["category"]} for metric, config in AI_METRICS.items()},
        "monthly_target_map": {
            config["target_label"]: safe_float(config["target"])
            for _, config in AI_METRICS.items()
            if config.get("target_label")
        },
        "standard_rules": standard_rules,
        "weeks": {"current": CURRENT_WEEK_LABEL, "previous": PREVIOUS_WEEK_LABEL},
        "monthly": clean_records(monthly),
        "team_monthly_total": clean_records(pd.DataFrame([team_monthly_total]))[0],
        "monthly_meta": {
            "current": CURRENT_MONTH_LABEL,
            "previous": PREVIOUS_MONTH_LABEL,
            "periods": team_monthly_total.get("periods", monthly_periods),
            "metrics": sorted({
                col.rsplit("_", 1)[0]
                for col in team_monthly_total.keys()
                if re.search(r"_\d{6}$", col)
            }),
        },
    }


def build_html(payload: dict, logo_src: str | None = None) -> str:
    data_json = json.dumps(payload, ensure_ascii=False, allow_nan=False, indent=2).replace("</", "<\\/")
    logic_srcdoc = html.escape(LOGIC_REFERENCE_PATH.read_text(encoding="utf-8"), quote=True)
    if logo_src is None:
        logo_base64 = base64.b64encode(LOGO_PATH.read_bytes()).decode("ascii")
        logo_src = f"data:image/jpeg;base64,{logo_base64}"
    template = """<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>AI监课周度管理面板 V2</title>
  <style>
    :root {
      --bg: #f4efe7;
      --panel: #fffaf3;
      --card: #fff;
      --ink: #16324f;
      --muted: #68788c;
      --line: #d9dfeb;
      --accent: #cb5f2b;
      --accent-soft: #f7d8c8;
      --good: #2f7d4f;
      --warn: #c98600;
      --bad: #b8442c;
      --shadow: 0 20px 50px rgba(22,50,79,.08);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "PingFang SC","Microsoft YaHei",sans-serif;
      background:
        radial-gradient(circle at top left, rgba(203,95,43,.18), transparent 24%),
        linear-gradient(180deg, #f8f4ed 0%, var(--bg) 100%);
      color: var(--ink);
    }
    .shell { max-width: 1500px; margin: 0 auto; padding: 28px 24px 60px; }
    .hero {
      display: block;
      margin-bottom: 22px;
    }
    .hero-main, .card, .table-wrap, .drill {
      background: linear-gradient(160deg, rgba(255,250,243,.94), rgba(255,255,255,.9));
      border: 1px solid rgba(203,95,43,.14);
      border-radius: 24px;
      box-shadow: var(--shadow);
    }
    .hero-main { padding: 24px; }
    .brand-lockup {
      display: flex;
      align-items: center;
      gap: 18px;
      margin-bottom: 14px;
      flex-wrap: wrap;
    }
    .brand-mark {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      padding: 10px 14px;
      border-radius: 18px;
      background: rgba(255,255,255,.82);
      border: 1px solid rgba(22,50,79,.08);
      box-shadow: 0 10px 22px rgba(22,50,79,.06);
    }
    .brand-mark img {
      display: block;
      height: 40px;
      width: auto;
    }
    .hero-copy {
      min-width: 320px;
      flex: 1;
    }
    .eyebrow {
      display: inline-block;
      margin-bottom: 8px;
      padding: 5px 10px;
      border-radius: 999px;
      background: rgba(22,50,79,.08);
      color: var(--ink);
      font-size: 11px;
      font-weight: 700;
      letter-spacing: .06em;
    }
    h1 { margin: 0 0 10px; font-size: 34px; line-height: 1.1; }
    .hero-main p, .hero-side ul, .section-head p { color: var(--muted); }
    .hero-main p { margin: 0; font-size: 14px; line-height: 1.7; max-width: 760px; }
    .tag-row { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 16px; }
    .tag { padding: 7px 12px; border-radius: 999px; font-size: 12px; background: var(--accent-soft); color: var(--accent); font-weight: 600; }
    .summary-grid { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 14px; margin-bottom: 22px; }
    .card { padding: 18px; }
    button.card {
      width: 100%;
      text-align: left;
      cursor: pointer;
      background: linear-gradient(160deg, rgba(255,250,243,.94), rgba(255,255,255,.9));
    }
    .team-metric-card.active {
      border-color: rgba(203,95,43,.5);
      box-shadow: 0 18px 36px rgba(203,95,43,.16);
    }
    .team-metric-card.status-good {
      border-color: rgba(47,125,79,.62);
      background: linear-gradient(160deg, rgba(208,237,217,.98), rgba(241,251,245,.94));
    }
    .team-metric-card.status-warn {
      border-color: rgba(201,134,0,.58);
      background: linear-gradient(160deg, rgba(255,231,173,.98), rgba(255,247,224,.94));
    }
    .team-metric-card.status-bad {
      border-color: rgba(184,68,44,.56);
      background: linear-gradient(160deg, rgba(247,203,194,.98), rgba(255,240,236,.94));
    }
    .team-sparkline {
      display: block;
      width: 100%;
      height: 46px;
      margin-top: 10px;
    }
    .summary-label { color: var(--muted); font-size: 12px; margin-bottom: 10px; }
    .summary-value { font-size: 28px; font-weight: 700; }
    .summary-sub { margin-top: 8px; color: var(--muted); font-size: 12px; }
    .filters {
      display: flex; flex-wrap: wrap; gap: 12px;
      background: rgba(255,255,255,.78);
      border: 1px solid var(--line);
      border-radius: 20px;
      padding: 14px;
      margin-bottom: 18px;
      position: sticky; top: 0; backdrop-filter: blur(8px); z-index: 4;
    }
    .filters label { display: block; font-size: 12px; color: var(--muted); margin-bottom: 6px; }
    .filters select, .filters input {
      width: 220px; padding: 10px 12px; border-radius: 12px;
      border: 1px solid var(--line); background: #fff; color: var(--ink);
    }
    .tab-row { display: flex; gap: 10px; margin-bottom: 18px; }
    .tab {
      border: 1px solid var(--line); background: rgba(255,255,255,.75);
      color: var(--ink); border-radius: 999px; padding: 10px 16px; cursor: pointer; font-weight: 600;
    }
    .tab.active { background: var(--ink); color: #fff; border-color: var(--ink); }
    .panel { display: none; }
    .panel.active { display: block; }
    .section-head { display: flex; justify-content: space-between; align-items: end; margin: 8px 0 12px; }
    .section-head h3 { margin: 0; font-size: 22px; }
    .focus-grid { display: grid; grid-template-columns: 1.1fr .9fr; gap: 14px; margin-bottom: 16px; }
    .focus-card, .category-board, .category-card, .table-wrap { background: #fff; border: 1px solid var(--line); border-radius: 20px; box-shadow: var(--shadow); }
    .focus-card { padding: 18px; }
    .focus-card h4, .category-card h4 { margin: 0 0 10px; font-size: 15px; }
    .focus-list { margin: 0; padding-left: 18px; color: var(--muted); line-height: 1.7; font-size: 13px; }
    .driver-category-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 12px; }
    .driver-category-card {
      border: 1px solid var(--line);
      border-radius: 16px;
      padding: 12px;
      background: linear-gradient(180deg, #fffdf9 0%, #fff 100%);
    }
    .driver-category-card h5 {
      margin: 0 0 10px;
      font-size: 13px;
      color: var(--ink);
    }
    .category-board { padding: 16px; margin-bottom: 18px; }
    .category-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 14px; }
    .category-card { padding: 14px; }
    .metric-table { width: 100%; border-collapse: collapse; font-size: 12px; }
    .metric-table th, .metric-table td { padding: 8px 6px; text-align: left; white-space: nowrap; border-bottom: 1px solid #edf1f6; }
    .table-wrap { overflow: auto; }
    table { width: 100%; border-collapse: collapse; font-size: 13px; }
    th, td { padding: 12px 12px; border-bottom: 1px solid #edf1f6; text-align: left; white-space: nowrap; }
    th { position: sticky; top: 0; background: #f8fafc; z-index: 1; color: var(--muted); font-weight: 700; }
    tr:hover td { background: #fff9f4; }
    .risk-high, .delta-bad { color: var(--bad); font-weight: 700; }
    .risk-mid { color: var(--warn); font-weight: 700; }
    .risk-low, .delta-good { color: var(--good); font-weight: 700; }
    .delta-flat { color: var(--muted); font-weight: 700; }
    .teacher-name { color: var(--accent); font-weight: 700; cursor: pointer; }
    .drill { display: none; margin-top: 14px; padding: 18px; }
    .drill.active { display: block; }
    .drill-grid { display: grid; grid-template-columns: 1.2fr 1fr 1fr; gap: 14px; }
    .mini-card { background: #fff; border: 1px solid var(--line); border-radius: 16px; padding: 14px; }
    .mini-card h4 { margin: 0 0 10px; font-size: 14px; }
    .mini-card ul { margin: 0; padding-left: 18px; color: var(--muted); line-height: 1.7; font-size: 13px; }
    .issue-table { width: 100%; font-size: 12px; border-collapse: collapse; }
    .issue-table th, .issue-table td { padding: 8px 6px; white-space: normal; }
    .teacher-issue-grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 10px; }
    .issue-pill { border: 1px solid var(--line); border-radius: 14px; padding: 10px 12px; background: #fffdf8; }
    .issue-pill strong { display: block; margin-bottom: 6px; font-size: 13px; }
    .issue-pill div { color: var(--muted); font-size: 12px; line-height: 1.6; }
    .problem-cell { min-width: 260px; }
    .problem-card {
      background: linear-gradient(180deg, #fffdf9 0%, #fff 100%);
      border: 1px solid #ece3d7;
      border-radius: 16px;
      padding: 10px 12px;
      box-shadow: inset 0 1px 0 rgba(255,255,255,.8);
    }
    .problem-card.empty {
      background: #fafafa;
      border-style: dashed;
      color: var(--muted);
      text-align: center;
    }
    .problem-name {
      font-size: 13px;
      font-weight: 700;
      color: var(--ink);
      margin-bottom: 8px;
      line-height: 1.4;
    }
    .problem-meta {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 6px;
    }
    .problem-stat {
      background: #f8f4ee;
      border-radius: 10px;
      padding: 6px 8px;
    }
    .problem-label {
      display: block;
      font-size: 11px;
      color: var(--muted);
      margin-bottom: 2px;
    }
    .problem-value {
      font-size: 13px;
      font-weight: 700;
      color: var(--ink);
    }
    .status-chip {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-width: 34px;
      padding: 2px 8px;
      border-radius: 999px;
      font-size: 11px;
      font-weight: 700;
      margin-left: 6px;
      vertical-align: middle;
    }
    .status-a, .status-pass { background: #e8f6ed; color: #2f7d4f; }
    .status-b, .status-watch { background: #fff4d6; color: #b57600; }
    .status-c, .status-d, .status-risk, .status-fail { background: #fde9e4; color: #b8442c; }
    .formula-note {
      background: #fff;
      border: 1px solid var(--line);
      border-radius: 16px;
      padding: 12px 14px;
      font-size: 12px;
      color: var(--muted);
      line-height: 1.7;
      margin-top: 10px;
    }
    .rule-note {
      margin-top: 8px;
      padding: 10px 12px;
      border-radius: 12px;
      background: #f8f4ee;
      color: var(--muted);
      font-size: 12px;
      line-height: 1.7;
    }
    .rule-title {
      font-weight: 700;
      color: var(--ink);
      margin-bottom: 4px;
      font-size: 12px;
    }
    .logic-note {
      background: #fff;
      border: 1px solid var(--line);
      border-radius: 16px;
      padding: 12px 14px;
      font-size: 12px;
      color: var(--muted);
      line-height: 1.7;
      margin-bottom: 14px;
    }
    .content-layout { display: block; }
    .main-column, .side-column { min-width: 0; }
    .toolbar-row {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 14px;
      margin-bottom: 14px;
      flex-wrap: wrap;
    }
    .logic-sidebar {
      display: flex;
      justify-content: flex-end;
      margin-bottom: 0;
      margin-left: auto;
    }
    .logic-toggle {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      width: min(360px, 100%);
      padding: 14px 16px;
      border: 1px solid var(--line);
      border-radius: 16px;
      background: linear-gradient(180deg, #fffdf9 0%, #fff 100%);
      color: var(--ink);
      cursor: pointer;
      font-size: 14px;
      font-weight: 700;
      box-shadow: var(--shadow);
    }
    .logic-toggle span:last-child {
      color: var(--muted);
      font-size: 12px;
      font-weight: 600;
    }
    .logic-overlay {
      position: fixed;
      inset: 0;
      background: rgba(15, 23, 42, 0.28);
      z-index: 60;
      display: none;
    }
    .logic-overlay.active {
      display: block;
    }
    .logic-frame-wrap {
      position: fixed;
      top: 18px;
      right: 18px;
      width: min(1200px, calc(100vw - 36px));
      height: calc(100vh - 36px);
      display: none;
      background: #fff;
      border: 1px solid var(--line);
      border-radius: 20px;
      box-shadow: var(--shadow);
      overflow: hidden;
      z-index: 70;
    }
    .logic-frame-wrap.active {
      display: flex;
      flex-direction: column;
    }
    .logic-frame-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      padding: 14px 16px;
      border-bottom: 1px solid var(--line);
      background: #fffdf9;
    }
    .logic-frame-head h3 {
      margin: 0;
      font-size: 16px;
    }
    .logic-close {
      border: 1px solid var(--line);
      background: #fff;
      color: var(--ink);
      border-radius: 10px;
      padding: 8px 12px;
      font-size: 12px;
      font-weight: 700;
      cursor: pointer;
    }
    .logic-frame {
      width: 100%;
      height: 100%;
      border: 0;
      background: #fff;
    }
    .logic-card {
      background: #fff;
      border: 1px solid var(--line);
      border-radius: 20px;
      box-shadow: var(--shadow);
      padding: 16px;
    }
    .logic-card h3, .logic-card h4 {
      margin: 0 0 8px;
    }
    .logic-card h3 { font-size: 17px; }
    .logic-card h4 { font-size: 14px; }
    .logic-card p {
      margin: 0;
      font-size: 12px;
      color: var(--muted);
      line-height: 1.7;
    }
    .logic-card ul {
      margin: 0;
      padding-left: 18px;
      color: var(--muted);
      line-height: 1.8;
      font-size: 12px;
    }
    .logic-kv {
      display: grid;
      grid-template-columns: 96px 1fr;
      gap: 6px 10px;
      font-size: 12px;
      color: var(--muted);
      line-height: 1.7;
    }
    .logic-kv strong {
      color: var(--ink);
      font-weight: 700;
    }
    .legend { display: flex; gap: 10px; flex-wrap: wrap; color: var(--muted); font-size: 12px; margin-bottom: 10px; }
    @media (max-width: 1120px) {
      .hero, .summary-grid, .drill-grid, .focus-grid, .category-grid, .driver-category-grid, .content-layout { grid-template-columns: 1fr; }
      .filters select, .filters input { width: 100%; }
      .toolbar-row { display: block; }
      .logic-sidebar { justify-content: stretch; margin-top: 10px; }
      .logic-toggle { width: 100%; }
      .logic-frame-wrap {
        top: 0;
        right: 0;
        width: 100vw;
        height: 100vh;
        border-radius: 0;
      }
    }
  </style>
</head>
<body>
  <div class="shell">
    <section class="hero">
      <div class="hero-main">
        <div class="brand-lockup">
          <div class="brand-mark">
            <img src="__LOGO_SRC__" alt="VIPTHINK">
          </div>
          <div class="hero-copy">
            <span class="eyebrow">VIPTHINK · Overseas Teaching QA</span>
            <h1>AI监课周度管理面板 V2</h1>
            <p>这版把 AI监课 更新成了 __PREVIOUS_WEEK__ 对 __CURRENT_WEEK__ 的周对比结构，并补充了 __CURRENT_MONTH__ 对 __PREVIOUS_MONTH__ 的月度汇总对比。现在既能看本周抓手，也能看月度变化趋势，方便业务部门和组长做周度和月度联动管理。</p>
          </div>
        </div>
        <div class="tag-row">
          <span class="tag">__CURRENT_WEEK__ 对比 __PREVIOUS_WEEK__</span>
          <span class="tag">业务部门 / 小组 / 老师下钻</span>
          <span class="tag">重点问题展示 当前值 / 目标值 / Gap</span>
          <span class="tag">__CURRENT_MONTH__ 对比 __PREVIOUS_MONTH__ 月度汇总</span>
        </div>
      </div>
    </section>

    <section class="summary-grid" id="summaryGrid"></section>
    <section class="logic-note" id="logicNote"></section>

    <section class="filters">
      <div>
        <label>业务部门</label>
        <select id="deptFilter"></select>
      </div>
      <div>
        <label>小组</label>
        <select id="groupFilter"></select>
      </div>
      <div>
        <label>指标分类</label>
        <select id="categoryFilter"></select>
      </div>
      <div>
        <label>老师搜索</label>
        <input id="teacherFilter" placeholder="输入老师姓名 / 工号">
      </div>
    </section>

    <div class="content-layout">
      <div class="main-column">
        <div class="toolbar-row">
          <div class="tab-row">
            <button class="tab active" data-panel="team">团队总览</button>
            <button class="tab" data-panel="dept">业务部门</button>
            <button class="tab" data-panel="group">小组</button>
            <button class="tab" data-panel="teacher">老师</button>
            <button class="tab" data-panel="month">月度汇总</button>
          </div>

          <aside class="side-column">
            <div class="logic-sidebar" id="logicSidebar"></div>
          </aside>
        </div>

        <section class="panel active" id="panel-team">
          <div class="section-head"><div><h3>海外益智团队总览</h3><p>展示 202601-202607 整体趋势，先看整个海外益智团队，再点开看单指标明细和环比。</p></div></div>
          <div class="focus-grid">
            <div class="focus-card" id="teamFocus"></div>
            <div class="focus-card" id="teamMeta"></div>
          </div>
          <div class="category-board" id="teamMetricBoard"></div>
          <div class="drill" id="teamMetricDetail"></div>
        </section>

        <section class="panel" id="panel-dept">
          <div class="section-head"><div><h3>业务部门看板</h3><p>看部门本周风险、较上周变化，以及最该抓的 3 个问题。</p></div></div>
          <div class="table-wrap"><table id="deptTable"></table></div>
          <div class="focus-grid">
            <div class="focus-card" id="deptFocus"></div>
            <div class="focus-card" id="deptMeta"></div>
          </div>
          <div class="category-board" id="deptCategories"></div>
        </section>

        <section class="panel" id="panel-group">
          <div class="section-head"><div><h3>小组看板</h3><p>看组长本周抓手是否清晰，重点关注风险占比和重点抓手变化。</p></div></div>
          <div class="table-wrap"><table id="groupTable"></table></div>
          <div class="focus-grid">
            <div class="focus-card" id="groupFocus"></div>
            <div class="focus-card" id="groupMeta"></div>
          </div>
          <div class="category-board" id="groupCategories"></div>
        </section>

        <section class="panel" id="panel-teacher">
          <div class="section-head"><div><h3>老师看板</h3><p>点击老师姓名，下钻看本周值、上周值、目标值与 Gap。</p></div></div>
          <div class="legend">
            <span>高风险：本周优先复盘</span>
            <span>中风险：组长跟进</span>
            <span>低风险：持续观察</span>
          </div>
          <div class="table-wrap"><table id="teacherTable"></table></div>
          <div class="drill" id="teacherDrill"></div>
        </section>

        <section class="panel" id="panel-month">
          <div class="section-head"><div><h3>AI监课月度汇总</h3><p>看 __CURRENT_MONTH__ 相比 __PREVIOUS_MONTH__ 的月度变化，便于和周度结果一起判断整体改善趋势。</p></div></div>
          <div class="focus-grid">
            <div class="focus-card" id="monthFocus"></div>
            <div class="focus-card" id="monthMeta"></div>
          </div>
          <div class="table-wrap"><table id="monthTable"></table></div>
        </section>
      </aside>
    </div>
  </div>

  <script id="payload-data" type="application/json">__DATA_JSON__</script>
  <script>
    const payload = JSON.parse(document.getElementById("payload-data").textContent);
    const currentWeek = payload.weeks?.current || "__CURRENT_WEEK__";
    const previousWeek = payload.weeks?.previous || "__PREVIOUS_WEEK__";
    const currentMonth = payload.monthly_meta?.current || "__CURRENT_MONTH__";
    const previousMonth = payload.monthly_meta?.previous || "__PREVIOUS_MONTH__";
    const state = { dept: "全部", group: "全部", category: "全部指标", teacherKeyword: "", teamMetric: null };
    const metricMeta = payload.metric_meta;
    const monthlyTargetMap = payload.monthly_target_map || {};
    const standardRules = payload.standard_rules || {};
    const metricCategoryMap = Object.fromEntries(Object.entries(payload.categories).flatMap(([category, metrics]) => metrics.map(metric => [metric, category])));
    const allMetrics = Object.keys(metricMeta);
    const monthlyPeriods = (payload.monthly_meta?.periods || []).filter(period => period >= "202601" && period <= currentMonth);

    function pct(v, digits=1) {
      if (v === null || v === undefined || Number.isNaN(v)) return "-";
      return (v * 100).toFixed(digits) + "%";
    }
    function num(v, digits=1) {
      if (v === null || v === undefined || Number.isNaN(v)) return "-";
      return Number(v).toFixed(digits);
    }
    function formatByMeta(metric, value) {
      const fmt = metricMeta[metric]?.format || "num1";
      if (fmt.startsWith("pct")) return pct(value, fmt === "pct0" ? 0 : 1);
      return num(value, fmt === "num0" ? 0 : 1);
    }
    function evaluateMetricStatus(metric, value) {
      const rule = standardRules[metric]?.status_rule;
      if (!rule || value === null || value === undefined || Number.isNaN(value)) return null;
      if (rule.type === "pass_fail") {
        return value >= 0.5 ? { label: "通过", cls: "status-pass" } : { label: "风险", cls: "status-risk" };
      }
      if (rule.type === "band") {
        for (const band of rule.bands) {
          const minOk = band.min === null || value >= band.min;
          const maxOk = band.max === null || value < band.max;
          if (minOk && maxOk) {
            return { label: band.level, cls: `status-${band.level.toLowerCase()}` };
          }
        }
      }
      return null;
    }
    function statusChip(metric, value) {
      const status = evaluateMetricStatus(metric, value);
      if (!status) return "";
      return `<span class="status-chip ${status.cls}">${status.label}</span>`;
    }
    function deltaClass(value) {
      if (value === null || value === undefined || Number.isNaN(value) || Math.abs(value) < 0.00001) return "delta-flat";
      return value > 0 ? "delta-good" : "delta-bad";
    }
    function deltaText(metric, value) {
      if (value === null || value === undefined || Number.isNaN(value)) return "-";
      const prefix = value > 0 ? "+" : "";
      const fmt = metricMeta[metric]?.format || "num1";
      return prefix + (fmt.startsWith("pct") ? (value * 100).toFixed(1) + "pp" : Number(value).toFixed(1));
    }
    function riskClass(level) {
      return level === "高" ? "risk-high" : level === "中" ? "risk-mid" : "risk-low";
    }

    function buildSummary() {
      const s = payload.summary;
      const cards = [
        [`${currentWeek}质检样本量`, s[`${currentWeek}监课量`], `上周 ${s[`${previousWeek}监课量`] ?? "-"}`],
        [`${currentWeek}覆盖老师数`, s[`${currentWeek}覆盖老师数`], `上周 ${s[`${previousWeek}覆盖老师数`] ?? "-"}`],
        ["高风险老师数", s["高风险老师数"], "建议本周优先复盘"],
        ["高风险占比", pct(s["高风险占比"]), "高风险老师 / 全部老师"],
        ["整体指标达成率", pct(s["整体指标达成率"]), `较上周 ${deltaText("感染力通过率", s["整体指标达成率变化"])}`],
        ["整体达成趋势", s["整体达成趋势"], "按有目标值指标汇总"],
        ["新增高风险老师数", s["新增高风险老师数"], "本周进入高风险范围的老师"],
        ["已改善老师数", s["已改善老师数"], "本周风险分较上周下降的老师"],
        ["持续高风险老师数", s["持续高风险老师数"], "本周仍需持续重点跟进"],
        ["平均风险变化", num(s["风险变化"], 1), `${currentWeek} 相比 ${previousWeek} 的平均风险变化`],
        ["本周重点关注范围", s["本周重点关注范围"], "建议优先关注和支持"],
        ["月度重点波动指标", s["月度重点波动指标"], `结合 ${currentMonth} 对 ${previousMonth} 观察`],
      ];
      document.getElementById("summaryGrid").innerHTML = cards.map(([label, value, sub]) => `
        <article class="card">
          <div class="summary-label">${label}</div>
          <div class="summary-value">${value}</div>
          <div class="summary-sub">${sub}</div>
        </article>
      `).join("");
      document.getElementById("logicNote").innerHTML = `
        <strong>当前判定口径说明：</strong>
        风险分只按“有目标值”的指标计算；
        没有目标值的指标只做展示和周度观察，不进入风险分和风险等级。
        风险分 = 目标值指标的轻/中/重度问题累计计分，其中重度 5 分、中度 2 分、轻度 0.5 分，再叠加指标分类权重。
        高风险优先看核心课堂质量问题组合，其次再看风险分分层。
        当前 Top抓手 = 基于当前范围先按指标分类拆开，再在各分类内按“未达标/走弱程度”排序。
      `;
    }

    function buildLogicSidebar() {
      document.getElementById("logicSidebar").innerHTML = `
        <button class="logic-toggle" id="logicToggle" type="button">
          <span>面板口径说明</span>
          <span id="logicToggleText">点击展开完整说明</span>
        </button>
        <div class="logic-overlay" id="logicOverlay"></div>
        <div class="logic-frame-wrap" id="logicFrameWrap">
          <div class="logic-frame-head">
            <h3>AI监课面板口径说明</h3>
            <button class="logic-close" id="logicClose" type="button">关闭</button>
          </div>
          <iframe class="logic-frame" srcdoc="__LOGIC_SRCDOC__" title="AI监课面板口径说明"></iframe>
        </div>
      `;
      const toggle = document.getElementById("logicToggle");
      const wrap = document.getElementById("logicFrameWrap");
      const overlay = document.getElementById("logicOverlay");
      const close = document.getElementById("logicClose");
      const text = document.getElementById("logicToggleText");
      const setOpen = (open) => {
        wrap.classList.toggle("active", open);
        overlay.classList.toggle("active", open);
        text.textContent = open ? "点击收起完整说明" : "点击展开完整说明";
      };
      toggle.addEventListener("click", () => setOpen(!wrap.classList.contains("active")));
      close.addEventListener("click", () => setOpen(false));
      overlay.addEventListener("click", () => setOpen(false));
    }

    function initFilters() {
      const departments = ["全部", ...new Set(payload.departments.map(d => d["业务部门"]))];
      const categories = ["全部指标", ...Object.keys(payload.categories)];
      document.getElementById("deptFilter").innerHTML = departments.map(v => `<option value="${v}">${v}</option>`).join("");
      document.getElementById("categoryFilter").innerHTML = categories.map(v => `<option value="${v}">${v}</option>`).join("");
      refreshGroupFilterOptions();
    }

    function refreshGroupFilterOptions() {
      let source = payload.groups.slice();
      if (state.dept !== "全部") source = source.filter(g => g["业务部门"] === state.dept);
      const naturalGroupSort = (a, b) => {
        const ax = a.match(/^(.*?)(\\d+)组$/);
        const bx = b.match(/^(.*?)(\\d+)组$/);
        if (ax && bx) {
          if (ax[1] !== bx[1]) return ax[1].localeCompare(bx[1], "zh-Hans-CN");
          return Number(ax[2]) - Number(bx[2]);
        }
        return a.localeCompare(b, "zh-Hans-CN");
      };
      const groups = ["全部", ...new Set(source.map(g => g["小组"]).sort(naturalGroupSort))];
      if (!groups.includes(state.group)) state.group = "全部";
      document.getElementById("groupFilter").innerHTML = groups.map(v => `<option value="${v}" ${v === state.group ? "selected" : ""}>${v}</option>`).join("");
    }

    function applyTeacherFilter(rows) {
      return rows.filter(r => {
        const deptOk = state.dept === "全部" || r["业务部门"] === state.dept;
        const groupOk = state.group === "全部" || r["小组"] === state.group;
        const text = `${r["老师"] || ""} ${r["老师工号"] || ""}`.toLowerCase();
        const keywordOk = !state.teacherKeyword || text.includes(state.teacherKeyword.toLowerCase());
        return deptOk && groupOk && keywordOk;
      });
    }

    function averageRows(rows) {
      if (!rows.length) return null;
      const base = {};
      allMetrics.forEach(metric => {
        const w2Key = `${metric}_W2`;
        const w1Key = `${metric}_W1`;
        const valuesW2 = rows.map(r => r[w2Key]).filter(v => v !== null && v !== undefined && !Number.isNaN(v));
        const valuesW1 = rows.map(r => r[w1Key]).filter(v => v !== null && v !== undefined && !Number.isNaN(v));
        base[w2Key] = valuesW2.length ? valuesW2.reduce((a, b) => a + b, 0) / valuesW2.length : null;
        base[w1Key] = valuesW1.length ? valuesW1.reduce((a, b) => a + b, 0) / valuesW1.length : null;
      });
      base["风险分"] = rows.map(r => r["风险分"] || 0).reduce((a, b) => a + b, 0) / rows.length;
      base["高风险占比"] = rows.filter(r => r["风险等级"] === "高").length / rows.length;
      return base;
    }

    function topDriverHtml(items) {
      if (!items || !items.length) return `<ul class="focus-list"><li>当前无优先预警项。</li></ul>`;
      return `<div class="teacher-issue-grid">${items.map(item => `
        <div class="issue-pill">
          <strong>${item["metric"]}</strong>
          <div>${currentWeek}：${formatByMeta(item["metric"], item["current"])}</div>
          <div>${previousWeek}：${formatByMeta(item["metric"], item["previous"])}</div>
          <div>目标：${item["target"] !== null && item["target"] !== undefined ? formatByMeta(item["metric"], item["target"]) : "-"}</div>
        </div>
      `).join("")}</div>`;
    }

    function categorizedTopDriverHtml(itemsByCategory) {
      const cards = Object.keys(payload.categories).map(category => `
        <div class="driver-category-card">
          <h5>${category}</h5>
          ${topDriverHtml((itemsByCategory && itemsByCategory[category]) || [])}
        </div>
      `).join("");
      return `<div class="driver-category-grid">${cards}</div>`;
    }

    function metricWeaknessScore(metric, current, previous) {
      if (current === null || current === undefined || Number.isNaN(current)) return -1;
      if (metricMeta[metric].target !== null && metricMeta[metric].target !== undefined) {
        const target = metricMeta[metric].target;
        if (target === 0) return current < target ? Math.abs(current - target) : 0;
        return Math.max(0, (target - current) / Math.abs(target));
      }
      if (previous === null || previous === undefined || Number.isNaN(previous)) return 0;
      const baseline = Math.abs(previous) > 1e-6 ? Math.abs(previous) : 1;
      return Math.max(0, (previous - current) / baseline);
    }

    function metricGap(metric, current) {
      const target = metricMeta[metric].target;
      if (
        target === null || target === undefined ||
        current === null || current === undefined || Number.isNaN(current)
      ) {
        return null;
      }
      return current - target;
    }

    function renderCategoryBoard(containerId, row, title) {
      const board = document.getElementById(containerId);
      if (!row) {
        board.innerHTML = `<h4>${title}</h4><div class="summary-sub">当前筛选范围暂无数据。</div>`;
        return;
      }
      const cards = Object.entries(payload.categories).map(([category, metrics]) => {
        const visibleMetrics = state.category === "全部指标" ? metrics : metrics.filter(metric => metricCategoryMap[metric] === state.category);
        if (!visibleMetrics.length) return "";
        const rows = visibleMetrics.map(metric => `
          <tr>
            <td>${metric}${statusChip(metric, row[`${metric}_W2`])}</td>
            <td>${formatByMeta(metric, row[`${metric}_W2`])}</td>
            <td>${formatByMeta(metric, row[`${metric}_W1`])}</td>
            <td class="${deltaClass((row[`${metric}_W2`] ?? 0) - (row[`${metric}_W1`] ?? 0))}">${deltaText(metric, (row[`${metric}_W2`] ?? NaN) - (row[`${metric}_W1`] ?? NaN))}</td>
            <td>${metricMeta[metric].target === null || metricMeta[metric].target === undefined ? "-" : formatByMeta(metric, metricMeta[metric].target)}</td>
            <td class="${deltaClass(metricGap(metric, row[`${metric}_W2`]))}">${metricGap(metric, row[`${metric}_W2`]) === null ? "-" : formatByMeta(metric, metricGap(metric, row[`${metric}_W2`]))}</td>
          </tr>
        `).join("");
        const firstRuleMetric = visibleMetrics.find(metric => standardRules[metric]);
        const ruleBlock = firstRuleMetric ? `
          <div class="rule-note">
            <div class="rule-title">${firstRuleMetric} 口径说明</div>
            <div>${standardRules[firstRuleMetric]["adjusted_standard"] || "-"}</div>
          </div>
        ` : "";
        return `
          <div class="category-card">
            <h4>${category}</h4>
            <table class="metric-table">
              <thead><tr><th>指标</th><th>${currentWeek}</th><th>${previousWeek}</th><th>变化</th><th>目标</th><th>Gap</th></tr></thead>
              <tbody>${rows}</tbody>
            </table>
            ${ruleBlock}
          </div>
        `;
      }).join("");
      board.innerHTML = `<div class="section-head"><div><h3>${title}</h3><p>Top抓手单独突出，完整指标在下方按分类展开，同时补充最新周相对目标值的 Gap。</p></div></div><div class="category-grid">${cards}</div>`;
    }

    function currentDeptScope() {
      if (state.dept === "全部") return { label: "全部业务部门", rows: payload.teachers.slice() };
      return { label: state.dept, rows: payload.teachers.filter(r => r["业务部门"] === state.dept) };
    }

    function currentGroupScope() {
      let rows = payload.teachers.slice();
      let label = "全部小组";
      if (state.dept !== "全部") {
        rows = rows.filter(r => r["业务部门"] === state.dept);
        label = state.dept;
      }
      if (state.group !== "全部") {
        rows = rows.filter(r => r["小组"] === state.group);
        label = state.group;
      }
      return { label, rows };
    }

    function buildScopeTopDrivers(rows) {
      const avg = averageRows(rows);
      if (!avg) return [];
      return allMetrics.map(metric => ({
        metric,
        current: avg[`${metric}_W2`],
        previous: avg[`${metric}_W1`],
        target: metricMeta[metric].target,
        category: metricMeta[metric].category,
        severity: (() => {
          if (avg[`${metric}_W2`] === null || avg[`${metric}_W2`] === undefined || Number.isNaN(avg[`${metric}_W2`])) return 0;
          if (metricMeta[metric].target === null || metricMeta[metric].target === undefined || avg[`${metric}_W2`] === null || avg[`${metric}_W2`] === undefined) {
            const wow = (avg[`${metric}_W2`] ?? 0) - (avg[`${metric}_W1`] ?? 0);
            return wow < 0 ? Math.abs(wow) : 0;
          }
          const gap = avg[`${metric}_W2`] - metricMeta[metric].target;
          return gap < 0 ? Math.abs(gap) : 0;
        })()
      })).filter(item => item.severity > 0).sort((a,b) => b.severity - a.severity).slice(0,3);
    }

    function buildScopeTopDriversByCategory(rows) {
      const avg = averageRows(rows);
      if (!avg) return {};
      const items = allMetrics.map(metric => ({
        metric,
        current: avg[`${metric}_W2`],
        previous: avg[`${metric}_W1`],
        target: metricMeta[metric].target,
        category: metricMeta[metric].category,
        severity: (() => {
          if (avg[`${metric}_W2`] === null || avg[`${metric}_W2`] === undefined || Number.isNaN(avg[`${metric}_W2`])) return 0;
          if (metricMeta[metric].target === null || metricMeta[metric].target === undefined || avg[`${metric}_W2`] === null || avg[`${metric}_W2`] === undefined) {
            const wow = (avg[`${metric}_W2`] ?? 0) - (avg[`${metric}_W1`] ?? 0);
            return wow < 0 ? Math.abs(wow) : 0;
          }
          const gap = avg[`${metric}_W2`] - metricMeta[metric].target;
          return gap < 0 ? Math.abs(gap) : 0;
        })()
      }));
      const grouped = {};
      Object.keys(payload.categories).forEach(category => {
        const categoryItems = items.filter(item => item.category === category);
        const riskItems = categoryItems.filter(item => item.severity > 0).sort((a, b) => b.severity - a.severity).slice(0, 3);
        if (riskItems.length) {
          grouped[category] = riskItems;
          return;
        }
        grouped[category] = categoryItems
          .map(item => ({ ...item, weakness: metricWeaknessScore(item.metric, item.current, item.previous) }))
          .filter(item => item.weakness >= 0)
          .sort((a, b) => b.weakness - a.weakness)
          .slice(0, 2);
      });
      return grouped;
    }

    function renderDeptInsights() {
      const scope = currentDeptScope();
      const topDriversByCategory = state.dept === "全部"
        ? (payload.summary["分类TOP问题抓手明细"] || {})
        : buildScopeTopDriversByCategory(scope.rows);
      const avg = averageRows(scope.rows);
      document.getElementById("deptFocus").innerHTML = `<h4>${scope.label} · 分类Top级抓手</h4>${categorizedTopDriverHtml(topDriversByCategory)}`;
      document.getElementById("deptMeta").innerHTML = avg ? `
        <h4>${scope.label} · 当前概览</h4>
        <ul class="focus-list">
          <li>${currentWeek} 老师数：${scope.rows.length}</li>
          <li>${currentWeek} 质检样本量：${scope.rows.map(r => r["样本课节数_W2"] || 0).reduce((a, b) => a + b, 0)}</li>
          <li>${currentWeek} 覆盖老师数：${scope.rows.filter(r => (r["样本课节数_W2"] || 0) > 0).length}</li>
          <li>${currentWeek} 平均风险分：${num(avg["风险分"], 1)}</li>
          <li>${currentWeek} 高风险占比：${pct(avg["高风险占比"])}</li>
          <li>${currentWeek} 指标达成率：${pct(avg["指标达成率_W2"])}</li>
          <li>达成趋势：${avg["指标达成趋势"] || "暂无趋势"}</li>
          <li>风险集中小组：${state.dept === "全部" ? payload.summary["风险最集中小组"] : (payload.groups.filter(g => g["业务部门"] === state.dept).sort((a,b) => (b["高风险占比"] || 0) - (a["高风险占比"] || 0))[0]?.["小组"] || "-")}</li>
        </ul>
      ` : `<h4>${scope.label} · 当前概览</h4><div class="summary-sub">暂无数据</div>`;
      renderCategoryBoard("deptCategories", avg, `${scope.label} · 分类指标`);
    }

    function renderGroupInsights() {
      const scope = currentGroupScope();
      const topDriversByCategory = buildScopeTopDriversByCategory(scope.rows);
      const avg = averageRows(scope.rows);
      document.getElementById("groupFocus").innerHTML = `<h4>${scope.label} · 分类Top级抓手</h4>${categorizedTopDriverHtml(topDriversByCategory)}`;
      document.getElementById("groupMeta").innerHTML = avg ? `
        <h4>${scope.label} · 当前概览</h4>
        <ul class="focus-list">
          <li>${currentWeek} 老师数：${scope.rows.length}</li>
          <li>${currentWeek} 质检样本量：${scope.rows.map(r => r["样本课节数_W2"] || 0).reduce((a, b) => a + b, 0)}</li>
          <li>${currentWeek} 覆盖老师数：${scope.rows.filter(r => (r["样本课节数_W2"] || 0) > 0).length}</li>
          <li>${currentWeek} 平均风险分：${num(avg["风险分"], 1)}</li>
          <li>${currentWeek} 高风险占比：${pct(avg["高风险占比"])}</li>
          <li>${currentWeek} 指标达成率：${pct(avg["指标达成率_W2"])}</li>
          <li>达成趋势：${avg["指标达成趋势"] || "暂无趋势"}</li>
          <li>当前业务部门：${state.dept === "全部" ? "全部" : state.dept}</li>
        </ul>
      ` : `<h4>${scope.label} · 当前概览</h4><div class="summary-sub">暂无数据</div>`;
      renderCategoryBoard("groupCategories", avg, `${scope.label} · 分类指标`);
    }

    function renderDeptTable() {
      let rows = payload.departments.slice();
      if (state.dept !== "全部") rows = rows.filter(r => r["业务部门"] === state.dept);
      const headers = ["业务部门", `质检样本量${currentWeek}`, `覆盖老师数${currentWeek}`, "老师数", `指标达成率${currentWeek}`, "达成趋势", "高风险老师数", "高风险占比", "风险变化", "重点抓手", `感染力通过率${currentWeek}`, "感染力变化", `课中正确率${currentWeek}`, "课中正确率变化", `课程预告${currentWeek}`, "课程预告变化"];
      document.getElementById("deptTable").innerHTML = `
        <thead><tr>${headers.map(h => `<th>${h}</th>`).join("")}</tr></thead>
        <tbody>${rows.map(r => `
          <tr>
            <td>${r["业务部门"]}</td>
            <td>${r["样本课节数_W2"] || 0}</td>
            <td>${r["覆盖老师数_W2"] || 0}</td>
            <td>${r["老师数"]}</td>
            <td>${pct(r["指标达成率_W2"])}</td>
            <td class="${deltaClass(r["指标达成率变化"])}">${r["指标达成趋势"] || "-"}</td>
            <td class="risk-high">${r["高风险老师数"]}</td>
            <td>${pct(r["高风险占比"])}</td>
            <td class="${deltaClass(r["风险变化"])}">${num(r["风险变化"], 1)}</td>
            <td>${r["重点抓手"]}</td>
            <td>${pct(r["感染力通过率_W2"])}</td>
            <td class="${deltaClass(r["感染力通过率_W2"] - r["感染力通过率_W1"])}">${deltaText("感染力通过率", r["感染力通过率_W2"] - r["感染力通过率_W1"])}</td>
            <td>${pct(r["课中正确率_W2"])}</td>
            <td class="${deltaClass(r["课中正确率_W2"] - r["课中正确率_W1"])}">${deltaText("课中正确率", r["课中正确率_W2"] - r["课中正确率_W1"])}</td>
            <td>${pct(r["课程预告_W2"])}</td>
            <td class="${deltaClass(r["课程预告_W2"] - r["课程预告_W1"])}">${deltaText("课程预告", r["课程预告_W2"] - r["课程预告_W1"])}</td>
          </tr>`).join("")}
        </tbody>`;
    }

    function renderGroupTable() {
      let rows = payload.groups.slice();
      if (state.dept !== "全部") rows = rows.filter(r => r["业务部门"] === state.dept);
      if (state.group !== "全部") rows = rows.filter(r => r["小组"] === state.group);
      const headers = ["业务部门", "小组", `质检样本量${currentWeek}`, `覆盖老师数${currentWeek}`, "老师数", `指标达成率${currentWeek}`, "达成趋势", "高风险占比", "风险变化", "重点抓手", `有效互动频次${currentWeek}`, "有效互动变化", `读题审题${currentWeek}`, "读题审题变化", `作业完成率${currentWeek}`];
      document.getElementById("groupTable").innerHTML = `
        <thead><tr>${headers.map(h => `<th>${h}</th>`).join("")}</tr></thead>
        <tbody>${rows.map(r => `
          <tr>
            <td>${r["业务部门"]}</td>
            <td>${r["小组"]}</td>
            <td>${r["样本课节数_W2"] || 0}</td>
            <td>${r["覆盖老师数_W2"] || 0}</td>
            <td>${r["老师数"]}</td>
            <td>${pct(r["指标达成率_W2"])}</td>
            <td class="${deltaClass(r["指标达成率变化"])}">${r["指标达成趋势"] || "-"}</td>
            <td>${pct(r["高风险占比"])}</td>
            <td class="${deltaClass(r["风险变化"])}">${num(r["风险变化"], 1)}</td>
            <td>${r["重点抓手"]}</td>
            <td>${num(r["有效互动频次_W2"], 1)}</td>
            <td class="${deltaClass(r["有效互动频次_W2"] - r["有效互动频次_W1"])}">${deltaText("有效互动频次", r["有效互动频次_W2"] - r["有效互动频次_W1"])}</td>
            <td>${pct(r["读题审题_W2"])}</td>
            <td class="${deltaClass(r["读题审题_W2"] - r["读题审题_W1"])}">${deltaText("读题审题", r["读题审题_W2"] - r["读题审题_W1"])}</td>
            <td>${pct(r["作业完成率_W2"])}</td>
          </tr>`).join("")}
        </tbody>`;
    }

    function issueSummaryCard(issue) {
      if (!issue) return `<div class="problem-card empty">-</div>`;
      return `
        <div class="problem-card">
          <div class="problem-name">${issue["metric"]}${statusChip(issue["metric"], issue["current"])}</div>
          <div class="problem-meta">
            <div class="problem-stat">
              <span class="problem-label">${currentWeek}</span>
              <span class="problem-value">${formatByMeta(issue["metric"], issue["current"])}</span>
            </div>
            <div class="problem-stat">
              <span class="problem-label">${previousWeek}</span>
              <span class="problem-value">${formatByMeta(issue["metric"], issue["previous"])}</span>
            </div>
            <div class="problem-stat">
              <span class="problem-label">目标</span>
              <span class="problem-value">${issue["target"] === null || issue["target"] === undefined ? "-" : formatByMeta(issue["metric"], issue["target"])}</span>
            </div>
          </div>
        </div>
      `;
    }

    function renderTeacherTable() {
      let rows = applyTeacherFilter(payload.teachers);
      const headers = ["老师", "业务部门", "小组", "风险等级", "风险分", "风险变化", "重点问题1", "重点问题2", "重点问题3", `${currentWeek}样本课节`];
      document.getElementById("teacherTable").innerHTML = `
        <thead><tr>${headers.map(h => `<th>${h}</th>`).join("")}</tr></thead>
        <tbody>${rows.map((r, idx) => `
          <tr>
            <td class="teacher-name" data-index="${idx}">${r["老师"]}</td>
            <td>${r["业务部门"]}</td>
            <td>${r["小组"]}</td>
            <td class="${riskClass(r["风险等级"])}">${r["风险等级"]}</td>
            <td>${num(r["风险分"], 1)}</td>
            <td class="${deltaClass(r["风险变化"])}">${num(r["风险变化"], 1)}</td>
            <td class="problem-cell">${issueSummaryCard((r["重点问题明细"] || [])[0])}</td>
            <td class="problem-cell">${issueSummaryCard((r["重点问题明细"] || [])[1])}</td>
            <td class="problem-cell">${issueSummaryCard((r["重点问题明细"] || [])[2])}</td>
            <td>${r["样本课节数_W2"] || 0}</td>
          </tr>`).join("")}
        </tbody>`;
      document.querySelectorAll(".teacher-name").forEach(cell => {
        cell.addEventListener("click", () => renderTeacherDrill(rows[Number(cell.dataset.index)]));
      });
    }

    function renderTeacherDrill(row) {
      const issues = (row["重点问题明细"] || []).filter(issue => state.category === "全部指标" || issue["category"] === state.category);
      const issueRows = issues.length ? issues.map(issue => `
        <tr>
          <td>${issue["metric"]}${issue["bucket"] ? `（${issue["bucket"]}）` : ""}</td>
          <td>${formatByMeta(issue["metric"], issue["current"])}</td>
          <td>${formatByMeta(issue["metric"], issue["previous"])}</td>
          <td>${issue["target"] === null || issue["target"] === undefined ? "-" : formatByMeta(issue["metric"], issue["target"])}</td>
          <td class="${deltaClass(issue["gap"])}">${issue["gap"] === null || issue["gap"] === undefined ? "-" : formatByMeta(issue["metric"], issue["gap"])}</td>
          <td>${issue["reason"]}</td>
        </tr>
      `).join("") : `<tr><td colspan="6">当前无优先预警项，可优先关注相对最弱项。</td></tr>`;

      const northstar = [
        ["课中首答正确率", row["课中首答正确率_W2"], row["课中首答正确率_W1"]],
        ["课中末答正确率", row["课中末答正确率_W2"], row["课中末答正确率_W1"]],
        ["课后首答正确率", row["课后首答正确率_W2"], row["课后首答正确率_W1"]],
        ["课后末答正确率", row["课后末答正确率_W2"], row["课后末答正确率_W1"]],
        ["作业完成率", row["作业完成率_W2"], row["作业完成率_W1"]],
        ["预习率", row["预习率_W2"], row["预习率_W1"]],
        ["小老师视频提交率", row["小老师视频提交率_W2"], row["小老师视频提交率_W1"]],
      ].filter(([, current, previous]) => {
        const hasCurrent = current !== null && current !== undefined && !Number.isNaN(current);
        const hasPrevious = previous !== null && previous !== undefined && !Number.isNaN(previous);
        return hasCurrent || hasPrevious;
      });
      const issueRuleNotes = issues.map(issue => {
        const rule = standardRules[issue["metric"]];
        if (!rule) return "";
        return `
          <div class="rule-note">
            <div class="rule-title">${issue["metric"]} 为什么看它</div>
            <div>${rule["detail"] || "-"}</div>
            <div style="margin-top:6px;"><strong>当前建议标准：</strong>${rule["adjusted_standard"] || "-"}</div>
          </div>
        `;
      }).join("");
      document.getElementById("teacherDrill").innerHTML = `
        <div class="section-head">
          <div>
            <h3>${row["老师"]} · 老师下钻</h3>
            <p>${row["业务部门"]} / ${row["小组"]} · 风险等级：<span class="${riskClass(row["风险等级"])}">${row["风险等级"]}</span></p>
          </div>
        </div>
        <div class="drill-grid">
          <div class="mini-card">
            <h4>重点问题明细</h4>
            <table class="issue-table">
              <thead><tr><th>指标</th><th>${currentWeek}</th><th>${previousWeek}</th><th>目标</th><th>Gap</th><th>判定依据</th></tr></thead>
              <tbody>${issueRows}</tbody>
            </table>
          </div>
          <div class="mini-card">
            <h4>AI监课补充观察</h4>
            <ul>
              <li>${currentWeek}样本课节数：${row["样本课节数_W2"] || 0}</li>
              <li>${previousWeek}样本课节数：${row["样本课节数_W1"] || 0}</li>
              <li>风险分变化：${num(row["风险变化"], 1)}</li>
              <li>当前重点问题：${row["重点问题"]}</li>
              <li>风险分口径：卡片看 A/B/C/D，风险等级按风险分分层</li>
            </ul>
            <div class="formula-note">
              风险分：按全部核心指标累计计分，重度 5 分、中度 2 分、轻度 0.5 分，
              再乘对应指标分类权重。只有有目标值的指标进入风险分；没有目标值的指标只做展示，不参与风险识别。
              有目标值的指标，看 ${currentWeek} 距目标差多少。
              当前高风险优先看核心课堂质量问题组合，再结合风险分分层判断。
              风险变化：当前这位老师的 ${currentWeek} 风险分，减去按 ${previousWeek} 同样口径计算的基线风险分。
            </div>
            ${issueRuleNotes}
          </div>
          <div class="mini-card">
            <h4>北极星辅助结果</h4>
            ${northstar.length
              ? `<ul>${northstar.map(([name, current, previous]) => `<li>${name}：${pct(current)} / 上周 ${pct(previous)}</li>`).join("")}</ul>`
              : `<div class="summary-sub">当前老师暂无可展示的北极星辅助结果。</div>`}
          </div>
        </div>
      `;
      document.getElementById("teacherDrill").classList.add("active");
      document.getElementById("teacherDrill").scrollIntoView({ behavior: "smooth", block: "nearest" });
    }

    function formatMonthlyMetric(metric, value) {
      if (value === null || value === undefined || Number.isNaN(value)) return "-";
      return metric.includes("率") ? pct(value) : num(value, 1);
    }

    function formatMonthlyDelta(metric, value) {
      if (value === null || value === undefined || Number.isNaN(value)) return "-";
      return metric.includes("率")
        ? `${value > 0 ? "+" : ""}${(value * 100).toFixed(1)}pp`
        : `${value > 0 ? "+" : ""}${Number(value).toFixed(1)}`;
    }

    function formatMonthLabel(period) {
      if (!period || period.length !== 6) return period || "-";
      return `${period.slice(0,4)}.${period.slice(4,6)}`;
    }

    function getTeamMonthlyRow() {
      return payload.team_monthly_total || null;
    }

    function getMetricSeries(row, metric) {
      return monthlyPeriods.map(period => ({
        period,
        value: row ? row[`${metric}_${period}`] : null,
      }));
    }

    function getMonthlyMetricTarget(metric) {
      if (monthlyTargetMap[metric] !== undefined) return monthlyTargetMap[metric];
      const compactMetric = metric.replace(/\\n/g, "");
      const matchedKey = Object.keys(monthlyTargetMap).find(key => key.replace(/\\n/g, "") === compactMetric);
      return matchedKey ? monthlyTargetMap[matchedKey] : null;
    }

    function teamMetricStatus(metric, currentValue) {
      const target = getMonthlyMetricTarget(metric);
      if (target === null || target === undefined || currentValue === null || currentValue === undefined || Number.isNaN(currentValue)) {
        return "";
      }
      const gap = currentValue - target;
      if (gap >= 0) return "status-good";
      const baseline = Math.abs(target) > 1e-6 ? Math.abs(target) : 1;
      const gapRatio = Math.abs(gap) / baseline;
      return gapRatio <= 0.1 ? "status-warn" : "status-bad";
    }

    function sparklineSvg(series, metric) {
      const values = series.map(item => item.value).filter(v => v !== null && v !== undefined && !Number.isNaN(v));
      if (!values.length) return `<div class="summary-sub">暂无趋势</div>`;
      const width = 180;
      const height = 46;
      const min = Math.min(...values);
      const max = Math.max(...values);
      const span = max - min || 1;
      const points = series.map((item, idx) => {
        const x = series.length === 1 ? width / 2 : (idx * (width - 8) / (series.length - 1)) + 4;
        const baseValue = item.value === null || item.value === undefined || Number.isNaN(item.value) ? min : item.value;
        const y = height - 6 - ((baseValue - min) / span) * (height - 12);
        return `${x.toFixed(1)},${y.toFixed(1)}`;
      }).join(" ");
      const last = series[series.length - 1];
      return `
        <svg viewBox="0 0 ${width} ${height}" class="team-sparkline" aria-hidden="true">
          <polyline fill="none" stroke="#cb5f2b" stroke-width="3" points="${points}" />
          <circle cx="${(series.length === 1 ? width / 2 : (series.length - 1) * (width - 8) / (series.length - 1) + 4).toFixed(1)}" cy="${points.split(' ').slice(-1)[0].split(',')[1]}" r="3.5" fill="#16324f"></circle>
        </svg>
      `;
    }

    function renderTeamPanel() {
      const row = getTeamMonthlyRow();
      const metrics = monthlyPeriods.length
        ? (payload.monthly_meta?.metrics || []).filter(metric => row && monthlyPeriods.some(period => row[`${metric}_${period}`] !== null && row[`${metric}_${period}`] !== undefined))
        : [];
      if (!state.teamMetric || !metrics.includes(state.teamMetric)) state.teamMetric = metrics[0] || null;

      document.getElementById("teamFocus").innerHTML = row ? `
        <h4>团队月度趋势</h4>
        <ul class="focus-list">
          <li>展示周期：${monthlyPeriods.map(formatMonthLabel).join(" / ")}</li>
          <li>当前最新月：${currentMonth}</li>
          <li>上一个月：${previousMonth}</li>
          <li>点击下方指标卡片，可看单指标详情和环比</li>
        </ul>
      ` : `<h4>团队月度趋势</h4><div class="summary-sub">当前暂无团队总计数据。</div>`;

      document.getElementById("teamMeta").innerHTML = row ? `
        <h4>团队关键概览</h4>
        <ul class="focus-list">
          <li>上台次数（BI）：${formatMonthlyMetric("上台次数（BI）", row[`上台次数（BI）_${currentMonth}`])}</li>
          <li>仪容仪表合格率：${formatMonthlyMetric("仪容仪表合格率", row[`仪容仪表合格率_${currentMonth}`])}</li>
          <li>感染力合格率：${formatMonthlyMetric("感染力合格率", row[`感染力合格率_${currentMonth}`])}</li>
          <li>作业提醒通过率：${formatMonthlyMetric("作业提醒通过率", row[`作业提醒通过率_${currentMonth}`])}</li>
        </ul>
      ` : `<h4>团队关键概览</h4><div class="summary-sub">暂无可展示数据。</div>`;

      const metricCards = metrics.map(metric => {
        const series = getMetricSeries(row, metric);
        const currentValue = row[`${metric}_${currentMonth}`];
        const previousValue = row[`${metric}_${previousMonth}`];
        const delta = currentValue === null || previousValue === null || currentValue === undefined || previousValue === undefined ? null : currentValue - previousValue;
        const active = state.teamMetric === metric ? " active" : "";
        const statusCls = teamMetricStatus(metric, currentValue);
        return `
          <button type="button" class="card team-metric-card ${statusCls}${active}" data-team-metric="${metric}">
            <div class="summary-label">${metricCategoryMap[metric] || "月度指标"} · ${metric}</div>
            <div class="summary-value">${formatMonthlyMetric(metric, currentValue)}</div>
            <div class="summary-sub">环比 ${formatMonthlyDelta(metric, delta)}</div>
            ${sparklineSvg(series, metric)}
          </button>
        `;
      }).join("");
      document.getElementById("teamMetricBoard").innerHTML = `<div class="summary-grid">${metricCards}</div>`;
      document.querySelectorAll("[data-team-metric]").forEach(btn => {
        btn.addEventListener("click", () => {
          state.teamMetric = btn.dataset.teamMetric;
          renderTeamPanel();
        });
      });

      renderTeamMetricDetail(row, state.teamMetric);
    }

    function renderTeamMetricDetail(row, metric) {
      const target = document.getElementById("teamMetricDetail");
      if (!row || !metric) {
        target.innerHTML = `<div class="summary-sub">当前暂无可展开的团队指标详情。</div>`;
        return;
      }
      const series = getMetricSeries(row, metric);
      const rows = series.map((item, idx) => {
        const prev = idx === 0 ? null : series[idx - 1].value;
        const wow = item.value === null || prev === null || item.value === undefined || prev === undefined ? null : item.value - prev;
        return `
          <tr>
            <td>${formatMonthLabel(item.period)}</td>
            <td>${formatMonthlyMetric(metric, item.value)}</td>
            <td class="${deltaClass(wow)}">${formatMonthlyDelta(metric, wow)}</td>
          </tr>
        `;
      }).join("");
      target.innerHTML = `
        <div class="section-head">
          <div>
            <h3>${metric} · 指标详情</h3>
            <p>查看 ${formatMonthLabel(monthlyPeriods[0])} 到 ${formatMonthLabel(monthlyPeriods[monthlyPeriods.length - 1])} 的连续变化，以及每月环比。</p>
          </div>
        </div>
        <div class="drill-grid">
          <div class="mini-card">
            <h4>趋势明细</h4>
            <table class="issue-table">
              <thead><tr><th>月份</th><th>数值</th><th>环比</th></tr></thead>
              <tbody>${rows}</tbody>
            </table>
          </div>
          <div class="mini-card">
            <h4>最新判断</h4>
            <ul>
              <li>当前值：${formatMonthlyMetric(metric, row[`${metric}_${currentMonth}`])}</li>
              <li>上月值：${formatMonthlyMetric(metric, row[`${metric}_${previousMonth}`])}</li>
              <li>最新环比：${formatMonthlyDelta(metric, (row[`${metric}_${currentMonth}`] ?? NaN) - (row[`${metric}_${previousMonth}`] ?? NaN))}</li>
              <li>指标分类：${metricCategoryMap[metric] || "其他"}</li>
            </ul>
            <div class="formula-note">
              这里展示的是整个海外益智团队在月度汇总“总计”行的指标结果。
              环比 = 当月数值 - 上月数值；百分比类展示为 pp，数值类展示为绝对值变化。
            </div>
          </div>
        </div>
      `;
      target.classList.add("active");
    }

    function renderMonthPanel() {
      const rows = (payload.monthly || []).filter(r => {
        const deptOk = state.dept === "全部" || r["老师属性"] === state.dept || r["老师属性"] === state.dept.replace("海外益智", "").replace("教学区", "");
        const groupOk = state.group === "全部" || r["小组"] === state.group;
        return deptOk && groupOk;
      });
      const metrics = payload.monthly_meta?.metrics || [];
      const metricCategories = {};
      metrics.forEach(metric => {
        const sample = rows.find(r => r[`${metric}_category`]);
        metricCategories[metric] = sample?.[`${metric}_category`] || "其他";
      });
      document.getElementById("monthFocus").innerHTML = `
        <h4>月度对比重点</h4>
        <ul class="focus-list">
          <li>月度口径：${currentMonth} 对比 ${previousMonth}</li>
          <li>数据来源：${currentWeek} 对应周表内“月度汇总”</li>
          <li>当前筛选后记录数：${rows.length}</li>
        </ul>
      `;
      document.getElementById("monthMeta").innerHTML = `
        <h4>使用说明</h4>
        <ul class="focus-list">
          <li>这一页只补月度变化，不改原有周度风险判定</li>
          <li>建议结合周度高风险老师名单一起看</li>
          <li>变化列 = ${currentMonth} - ${previousMonth}</li>
        </ul>
      `;
      const headers = ["老师属性", "小组", ...metrics.flatMap(metric => [`${metricCategories[metric]}·${metric} ${previousMonth}`, `${metric} ${currentMonth}`, `${metric} 变化`])];
      document.getElementById("monthTable").innerHTML = `
        <thead><tr>${headers.map(h => `<th>${h}</th>`).join("")}</tr></thead>
        <tbody>${rows.map(r => `
          <tr>
            <td>${r["老师属性"] || "-"}</td>
            <td>${r["小组"] || "-"}</td>
            ${metrics.map(metric => `
              <td>${formatMonthlyMetric(metric, r[`${metric}_${previousMonth}`])}</td>
              <td>${formatMonthlyMetric(metric, r[`${metric}_${currentMonth}`])}</td>
              <td class="${deltaClass(r[`${metric}_变化`])}">${formatMonthlyDelta(metric, r[`${metric}_变化`])}</td>
            `).join("")}
          </tr>
        `).join("")}</tbody>
      `;
    }

    function renderAll() {
      refreshGroupFilterOptions();
      renderTeamPanel();
      renderDeptInsights();
      renderGroupInsights();
      renderDeptTable();
      renderGroupTable();
      renderTeacherTable();
      renderMonthPanel();
    }

    buildSummary();
    buildLogicSidebar();
    initFilters();
    renderAll();

    document.getElementById("deptFilter").addEventListener("change", e => { state.dept = e.target.value; state.group = "全部"; renderAll(); });
    document.getElementById("groupFilter").addEventListener("change", e => { state.group = e.target.value; renderAll(); });
    document.getElementById("categoryFilter").addEventListener("change", e => { state.category = e.target.value; renderAll(); });
    document.getElementById("teacherFilter").addEventListener("input", e => { state.teacherKeyword = e.target.value.trim(); renderTeacherTable(); });
    document.querySelectorAll(".tab").forEach(btn => {
      btn.addEventListener("click", () => {
        document.querySelectorAll(".tab").forEach(x => x.classList.remove("active"));
        document.querySelectorAll(".panel").forEach(x => x.classList.remove("active"));
        btn.classList.add("active");
        document.getElementById(`panel-${btn.dataset.panel}`).classList.add("active");
      });
    });
  </script>
</body>
</html>"""
    return (
        template
        .replace("__DATA_JSON__", data_json)
        .replace("__LOGIC_SRCDOC__", logic_srcdoc)
        .replace("__LOGO_SRC__", logo_src)
        .replace("__CURRENT_WEEK__", CURRENT_WEEK_LABEL)
        .replace("__PREVIOUS_WEEK__", PREVIOUS_WEEK_LABEL)
        .replace("__CURRENT_MONTH__", CURRENT_MONTH_LABEL)
        .replace("__PREVIOUS_MONTH__", PREVIOUS_MONTH_LABEL)
    )


def main() -> None:
    payload = build_payload()
    OUTPUT.write_text(build_html(payload), encoding="utf-8")
    print(OUTPUT)


if __name__ == "__main__":
    main()
