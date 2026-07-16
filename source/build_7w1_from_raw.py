from __future__ import annotations

import math
import re
from pathlib import Path

import pandas as pd
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Alignment, Font, PatternFill


RAW_PARTS = [
    Path("/Users/lilblackmac/Desktop/质检记录列表7.1-7.3.xlsx"),
    Path("/Users/lilblackmac/Desktop/质检记录列表7.4-7.5.xlsx"),
    Path("/Users/lilblackmac/Desktop/质检记录列表7.6-7.7.xlsx"),
]
EFFECTIVE_PATH = Path("/Users/lilblackmac/Desktop/AI监课有效课节数7.1-7.7.xlsx")
BI_PATH = Path("/Users/lilblackmac/Desktop/AI监课上台次数7.1-7.7.xlsx")
STANDARD_PATH = Path("/Users/lilblackmac/Desktop/AI监课&北极星/海外AI监课质检标准_优化版.xlsx")
REFERENCE_W5_PATH = Path("/Users/lilblackmac/Desktop/AI监课&北极星/海外监课指标6月W5.xlsx")
OUTPUT_DIR = Path("/Users/lilblackmac/Documents/New project/outputs/7w1_rebuild")
PERIOD_LABEL = "W1"


RESULT_COLUMNS = {
    "背景布": "AI评测-背景布-结果",
    "工服": "AI评测-工服-结果",
    "光线": "AI评测-光线镜头-结果",
    "妆发": "AI评测-妆发-结果",
    "坐姿": "AI评测-坐姿-结果",
    "行为举止": "AI评测-行为举止-结果",
    "感染力": "AI评测-热情洋溢感染力强-结果",
    "课前作业提醒": "AI评测-学习习惯检查-结果",
    "课后作业布置": "AI评测-布置作业-结果",
    "知识点讲错": "AI评测-知识点讲错-结果",
    "知识点漏讲": "AI评测-知识点漏讲-结果",
    "课堂规则提及": "AI评测-课堂规则-结果",
    "读题审题": "AI评测-读题审题-结果",
    "关键提问": "AI评测-关键提问（依据详案）-结果",
    "情绪策略": "AI评测-使用情绪策略（依据详案）-结果",
    "课程预告": "AI评测-下节预告（动画+知识点）-结果",
}

DETAIL_COLUMNS = {
    "感染力分数": "AI评测-热情洋溢感染力强-明细",
    "声音洪亮分数": "AI评测-声音洪亮活力四射-明细",
    "开口时长(s)": "AI评测-学员开口时长（均值）-明细",
    "课中正确率": "AI评测-课堂正确率（均值）-明细",
    "上台平均数": "AI评测-学员参与度（均值）-明细",
    "主动性": "AI评测-主动性（均值 ）-明细",
    "学员数": "AI评测-学员参与度（学生占比）-明细",
    "镜头感": "AI评测-镜头感（均值）-明细",
    "愉悦度": "AI评测-愉悦度（均值）-明细",
    "有效互动次数": "AI评测-有效学员互动-明细",
}

TEACHER_OUTPUT_COLUMNS = [
    "标签分组",
    "老师属性",
    "大组",
    "小组",
    "id",
    "昵称",
    "上台次数（BI）",
    "有效互动频次",
    "感染力合格率",
    "仪容仪表合格率",
    "背景布合格率",
    "工服合格率",
    "光线合格率",
    "妆发合格率",
    "坐姿合格率",
    "行为合格率",
    "课前作业提醒合格率",
    "课后作业布置合格率",
    "感染力分数",
    "声音洪亮分数",
    "知识点讲错通过率",
    "知识点漏讲通过率",
    "课堂规则提及通过率",
    "开口时长（s）",
    "课中正确率",
    "学员正面动作次数",
    "读题审题通过率",
    "关键提问通过率",
    "情绪策略使用通过率",
    "课程预告通过率",
]

GROUP_OUTPUT_COLUMNS = [
    "团队",
    "样本数",
    "老师数",
    "感染力分数",
    "声音洪亮分数",
    "知识点讲错",
    "知识点漏讲",
    "课堂规则提及",
    "开口时长",
    "课中正确率",
    "上台次数",
    "学员正面动作次数",
    "读题审题",
    "关键提问",
    "情绪策略",
    "课程预告",
]

SD_LOWER_COLUMNS = [
    "团队",
    "样本数 ",
    "老师数",
    "背景布合格率",
    "工服合格率",
    "光线合格率",
    "妆发合格率",
    "坐姿合格率",
    "行为合格率",
    "感染力合格率",
    "有效互动频次",
    "课前作业提醒合格率",
    "课后作业布置合格率",
]

LABEL_MAP = {
    "海外益智英语教学区": "英语",
    "海外益智粤语教学区": "粤语",
    "海外益智台湾教学区": "台教",
    "海外益智外教教学区": "菲教",
}

MONTH_DETAIL_METRICS = [
    ("仪容仪表合格率", "仪容仪表合格率"),
    ("背景布合格率", "背景布\n合格率"),
    ("工服合格率", "工服\n合格率"),
    ("光线合格率", "光线\n合格率"),
    ("妆发合格率", "妆发\n合格率"),
    ("坐姿合格率", "坐姿\n合格率"),
    ("行为合格率", "行为\n合格率"),
    ("课前作业提醒合格率", "课前作业提醒合格率"),
    ("课后作业布置合格率", "课后作业布置合格率"),
    ("感染力分数", "感染力\n分数"),
    ("声音洪亮分数", "声音洪亮分数"),
    ("知识点讲错通过率", "知识点讲错通过率"),
    ("知识点漏讲通过率", "知识点漏讲通过率"),
    ("课堂规则提及通过率", "课堂规则提及通过率"),
    ("开口时长（s）", "开口时长\n（s）"),
    ("课中正确率", "课中正确率"),
    ("学员正面动作次数", "学员正面动作次数"),
    ("读题审题通过率", "读题审题通过率"),
    ("关键提问通过率", "关键提问通过率"),
    ("情绪策略使用通过率", "情绪策略使用通过率"),
    ("课程预告通过率", "课程预告通过率"),
]

MONTH_TARGETS = {
    "上台次数（BI）": 6.0,
    "感染力合格率": 0.5,
    "有效互动频次": 25.0,
}

TEACHER_PERCENT_COLUMNS = {
    "感染力合格率",
    "仪容仪表合格率",
    "背景布\n合格率",
    "工服\n合格率",
    "光线\n合格率",
    "妆发\n合格率",
    "坐姿\n合格率",
    "行为\n合格率",
    "课前作业提醒合格率",
    "课后作业布置合格率",
    "感染力\n分数",
    "声音洪亮分数",
    "知识点讲错通过率",
    "知识点漏讲通过率",
    "课堂规则提及通过率",
    "课中正确率",
    "读题审题通过率",
    "关键提问通过率",
    "情绪策略使用通过率",
    "课程预告通过率",
}

SD_PERCENT_COLUMNS = {
    "感染力分数",
    "声音洪亮分数 ",
    "知识点讲错",
    "知识点漏讲",
    "课堂规则提及",
    "课中正确率 ",
    "读题审题 ",
    "关键提问 ",
    "情绪策略 ",
    "课程预告 ",
}

TEACHER_PERCENT_2_COLUMNS = {
    "感染力合格率",
    "感染力\n分数",
    "声音洪亮分数",
    "课中正确率",
}

TEACHER_PERCENT_0_COLUMNS = TEACHER_PERCENT_COLUMNS - TEACHER_PERCENT_2_COLUMNS

SD_PERCENT_2_COLUMNS = {
    "感染力分数",
    "声音洪亮分数 ",
    "知识点讲错",
    "知识点漏讲",
    "课堂规则提及",
    "课中正确率 ",
    "读题审题 ",
    "关键提问 ",
    "情绪策略 ",
    "课程预告 ",
}

TEACHER_GROOMING_COLUMNS = {
    "仪容仪表合格率",
    "背景布\n合格率",
    "工服\n合格率",
    "光线\n合格率",
    "妆发\n合格率",
    "坐姿\n合格率",
    "行为\n合格率",
}

TEACHER_SCALE_COLUMNS = {
    "上台次数（BI）",
    "有效互动频次",
    "感染力合格率",
    "课前作业提醒合格率",
    "课后作业布置合格率",
    "感染力\n分数",
    "声音洪亮分数",
    "知识点讲错通过率",
    "知识点漏讲通过率",
    "课堂规则提及通过率",
    "开口时长\n（s）",
    "课中正确率",
    "学员正面动作次数",
    "读题审题通过率",
    "关键提问通过率",
    "情绪策略使用通过率",
    "课程预告通过率",
}

SD_SCALE_COLUMNS = {
    "感染力分数",
    "声音洪亮分数 ",
    "知识点讲错",
    "知识点漏讲",
    "课堂规则提及",
    "开口时长",
    "课中正确率 ",
    "上台次数",
    "学员正面动作次数",
    "读题审题 ",
    "关键提问 ",
    "情绪策略 ",
    "课程预告 ",
}


def safe_float(value: object) -> float | None:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def parse_pct_from_text(text: object) -> float | None:
    if text is None or (isinstance(text, float) and math.isnan(text)):
        return None
    match = re.search(r"([0-9]+(?:\.[0-9]+)?)%", str(text))
    return None if not match else float(match.group(1)) / 100


def parse_number_after_colon(text: object) -> float | None:
    if text is None or (isinstance(text, float) and math.isnan(text)):
        return None
    match = re.search(r":\s*([0-9]+(?:\.[0-9]+)?)", str(text))
    return None if not match else float(match.group(1))


def parse_active_student_count(text: object) -> float | None:
    if text is None or (isinstance(text, float) and math.isnan(text)):
        return None
    for pattern in [r"上下台≥1的学生数[:：]\s*([0-9]+)", r"学生数[:：]\s*([0-9]+)", r"掌握学生数[:：]\s*([0-9]+)"]:
        match = re.search(pattern, str(text))
        if match:
            return float(match.group(1))
    return None


def parse_total_students_from_participation(text: object) -> float | None:
    if text is None or (isinstance(text, float) and math.isnan(text)):
        return None
    raw = str(text)
    count_match = re.search(r"上下台≥1的学生数[:：]\s*([0-9]+)", raw)
    ratio_match = re.search(r"上下台≥1的学生占比[:：]\s*([0-9]+(?:\.[0-9]+)?)%", raw)
    if not count_match or not ratio_match:
        return None
    count = float(count_match.group(1))
    ratio = float(ratio_match.group(1)) / 100
    if ratio <= 0:
        return None
    return round(count / ratio, 0)


def parse_interaction_frequency(text: object) -> float | None:
    if text is None or (isinstance(text, float) and math.isnan(text)):
        return None
    total_match = re.search(r"老师互动动作总数[:：]\s*([0-9]+(?:\.[0-9]+)?)", str(text))
    return None if not total_match else float(total_match.group(1))


def teacher_alias(name: object) -> str:
    if name is None:
        return ""
    text = str(name).strip()
    return re.sub(r"[（(].*?[）)]", "", text).strip()


def teacher_alias_loose(name: object) -> str:
    text = teacher_alias(name)
    return text.replace("老師", "老师").replace("　", " ").strip().lower()


def make_trend_sparkline(values: list[object]) -> str:
    nums = [safe_float(v) for v in values]
    nums = [v for v in nums if v is not None]
    if not nums:
        return "-"
    if len(set(nums)) == 1:
        return "▅" * len(nums)
    chars = "▁▂▃▄▅▆▇█"
    lo, hi = min(nums), max(nums)
    out = []
    for v in values:
        n = safe_float(v)
        if n is None:
            out.append("·")
            continue
        idx = int(round((n - lo) / (hi - lo) * (len(chars) - 1)))
        out.append(chars[max(0, min(len(chars) - 1, idx))])
    return "".join(out)


def sanitize_excel_text(df: pd.DataFrame) -> pd.DataFrame:
    safe = df.copy()
    for col in safe.columns:
        if safe[col].dtype == "object":
            safe[col] = safe[col].map(
                lambda v: (
                    str(v).replace("https://", "hxxps://")
                    if isinstance(v, str) and "https://" in v and len(v) > 500
                    else v
                )
            )
    return safe


def extract_department_parts(value: object) -> tuple[str, str]:
    if value is None:
        return "", ""
    parts = [part.strip() for part in str(value).split("/") if part and str(part).strip()]
    if len(parts) >= 2:
        return parts[-2], parts[-1]
    if len(parts) == 1:
        return parts[0], ""
    return "", ""


def load_raw() -> pd.DataFrame:
    frames = [pd.read_excel(path) for path in RAW_PARTS]
    raw = pd.concat(frames, ignore_index=True)
    raw = raw.drop_duplicates(subset=["id"], keep="first").reset_index(drop=True)
    return raw


def load_invalid_keys() -> set[tuple[str, str, str, str, str]]:
    invalid = pd.read_excel(EFFECTIVE_PATH)
    invalid = invalid.copy()
    invalid["主讲昵称"] = invalid["主讲昵称"].astype(str).str.strip()
    invalid["主讲六级部门"] = invalid["主讲六级部门"].astype(str).str.strip()
    invalid["主讲七级部门"] = invalid["主讲七级部门"].astype(str).str.strip()
    invalid["课程阶段"] = invalid["课程阶段"].astype(str).str.strip().str.upper()
    invalid["上课日期"] = pd.to_datetime(invalid["上课日期"], errors="coerce").dt.strftime("%Y-%m-%d")
    invalid = invalid[pd.to_numeric(invalid["有效课节数"], errors="coerce").fillna(0) == 0].copy()
    keys = set(
        zip(
            invalid["主讲昵称"],
            invalid["主讲六级部门"],
            invalid["主讲七级部门"],
            invalid["课程阶段"],
            invalid["上课日期"],
        )
    )
    return keys


def build_detail(raw: pd.DataFrame) -> pd.DataFrame:
    detail = raw.copy()
    invalid_keys = load_invalid_keys()
    detail.insert(0, "周期", PERIOD_LABEL)
    detail["教师艺名"] = detail["教师姓名"].map(teacher_alias)
    depts = detail["教师部门"].map(extract_department_parts)
    detail["教师部门5"] = [item[0] for item in depts]
    detail["教师部门6"] = [item[1] for item in depts]
    detail["课程阶段"] = detail["课程编码"].astype(str).str.extract(r"^(s\d+)", expand=False).str.upper()
    detail["上课日期"] = pd.to_datetime(detail["上课时间"].astype(str).str.extract(r"(^\d{4}-\d{2}-\d{2})")[0], errors="coerce").dt.strftime("%Y-%m-%d")
    detail.insert(
        1,
        "样本数",
        detail.apply(
            lambda row: 0
            if (
                row["教师艺名"],
                row["教师部门5"],
                row["教师部门6"],
                row["课程阶段"],
                row["上课日期"],
            )
            in invalid_keys
            else 1,
            axis=1,
        ),
    )
    detail.insert(2, "老师计数", 1)
    detail["老师属性"] = detail["教师部门5"].map(LABEL_MAP).fillna("")

    for output_name, source_col in RESULT_COLUMNS.items():
        detail[output_name] = (detail[source_col] == "通过").astype(float)

    detail["感染力\n分数"] = detail[DETAIL_COLUMNS["感染力分数"]].map(parse_pct_from_text)
    detail["声音洪亮\n分数"] = detail[DETAIL_COLUMNS["声音洪亮分数"]].map(parse_pct_from_text)
    detail["开口时长(s)"] = detail[DETAIL_COLUMNS["开口时长(s)"]].map(parse_number_after_colon)
    detail["课中正确率"] = detail[DETAIL_COLUMNS["课中正确率"]].map(parse_pct_from_text)
    detail["上台平均数"] = detail[DETAIL_COLUMNS["上台平均数"]].map(parse_number_after_colon)
    detail["主动性"] = detail[DETAIL_COLUMNS["主动性"]].map(parse_number_after_colon)
    detail["上台学生数"] = detail[DETAIL_COLUMNS["学员数"]].map(parse_active_student_count)
    detail["学员数"] = detail[DETAIL_COLUMNS["学员数"]].map(parse_total_students_from_participation)
    detail["镜头感"] = detail[DETAIL_COLUMNS["镜头感"]].map(parse_pct_from_text)
    detail["愉悦度"] = detail[DETAIL_COLUMNS["愉悦度"]].map(parse_pct_from_text)
    detail["有效互动次数"] = detail[DETAIL_COLUMNS["有效互动次数"]].map(parse_interaction_frequency)
    detail["有效互动频次"] = detail.apply(
        lambda row: None
        if row["学员数"] is None or pd.isna(row["学员数"]) or row["学员数"] <= 0
        else safe_float(row["有效互动次数"] / row["学员数"]) if row["有效互动次数"] is not None and not pd.isna(row["有效互动次数"]) else None,
        axis=1,
    )
    return detail


def ratio_metric(group: pd.DataFrame, col: str) -> float | None:
    valid_group = group[group["样本数"] == 1]
    values = pd.to_numeric(valid_group[col], errors="coerce").dropna()
    if values.empty:
        return None
    return safe_float(values.mean())


def average_metric(group: pd.DataFrame, col: str) -> float | None:
    valid_group = group[group["样本数"] == 1]
    values = pd.to_numeric(valid_group[col], errors="coerce").dropna()
    if values.empty:
        return None
    return safe_float(values.mean())


def load_bi() -> pd.DataFrame:
    bi = pd.read_excel(BI_PATH)
    bi = bi.rename(
        columns={
            "主讲ID": "id",
            "主讲昵称": "昵称",
            "主讲六级部门": "大组",
            "主讲七级部门": "小组",
            "上台次数": "上台次数（BI）",
        }
    )
    bi["昵称"] = bi["昵称"].astype(str).str.strip()
    bi["昵称_alias"] = bi["昵称"].map(teacher_alias)
    bi["昵称_alias_loose"] = bi["昵称"].map(teacher_alias_loose)
    bi["id"] = pd.to_numeric(bi["id"], errors="coerce")
    return bi[["id", "昵称", "昵称_alias", "昵称_alias_loose", "大组", "小组", "上台次数（BI）", "有效课节数", "应出勤人数", "出勤人数", "总上台次数"]]


def build_teacher_summary(detail: pd.DataFrame, bi: pd.DataFrame) -> pd.DataFrame:
    records: list[dict[str, object]] = []
    grouped = detail.groupby(["教师艺名", "教师工号", "教师部门5", "教师部门6", "老师属性"], dropna=False)
    for (teacher, teacher_id, dept, group_name, label), group in grouped:
        display_label = "兼职" if str(group_name).strip() == "兼职老师组" else label
        row: dict[str, object] = {
            "标签分组": display_label,
            "老师属性": display_label,
            "大组": dept,
            "小组": group_name,
            "id": pd.to_numeric(teacher_id, errors="coerce"),
            "昵称": teacher,
            "上台次数（BI）": None,
            "有效互动频次": average_metric(group, "有效互动频次"),
            "感染力合格率": ratio_metric(group, "感染力"),
            "背景布合格率": ratio_metric(group, "背景布"),
            "工服合格率": ratio_metric(group, "工服"),
            "光线合格率": ratio_metric(group, "光线"),
            "妆发合格率": ratio_metric(group, "妆发"),
            "坐姿合格率": ratio_metric(group, "坐姿"),
            "行为合格率": ratio_metric(group, "行为举止"),
            "课前作业提醒合格率": ratio_metric(group, "课前作业提醒"),
            "课后作业布置合格率": ratio_metric(group, "课后作业布置"),
            "感染力分数": average_metric(group, "感染力\n分数"),
            "声音洪亮分数": average_metric(group, "声音洪亮\n分数"),
            "知识点讲错通过率": ratio_metric(group, "知识点讲错"),
            "知识点漏讲通过率": ratio_metric(group, "知识点漏讲"),
            "课堂规则提及通过率": ratio_metric(group, "课堂规则提及"),
            "开口时长（s）": average_metric(group, "开口时长(s)"),
            "课中正确率": average_metric(group, "课中正确率"),
            "学员正面动作次数": average_metric(group, "主动性"),
            "读题审题通过率": ratio_metric(group, "读题审题"),
            "关键提问通过率": ratio_metric(group, "关键提问"),
            "情绪策略使用通过率": ratio_metric(group, "情绪策略"),
            "课程预告通过率": ratio_metric(group, "课程预告"),
        }
        grooming = [
            row["背景布合格率"],
            row["工服合格率"],
            row["光线合格率"],
            row["妆发合格率"],
            row["坐姿合格率"],
            row["行为合格率"],
        ]
        grooming_values = [value for value in grooming if value is not None]
        row["仪容仪表合格率"] = safe_float(sum(grooming_values) / len(grooming_values)) if grooming_values else None
        records.append(row)

    teacher_df = pd.DataFrame(records)
    teacher_df["id"] = pd.to_numeric(teacher_df["id"], errors="coerce")
    teacher_df["昵称_alias"] = teacher_df["昵称"].map(teacher_alias)
    teacher_df["昵称_alias_loose"] = teacher_df["昵称"].map(teacher_alias_loose)
    teacher_df = teacher_df.merge(
        bi[["id", "昵称", "大组", "小组", "上台次数（BI）"]],
        on=["id", "昵称", "大组", "小组"],
        how="left",
        suffixes=("", "_bi"),
    )
    teacher_df["上台次数（BI）"] = teacher_df["上台次数（BI）"].combine_first(teacher_df["上台次数（BI）_bi"])
    teacher_df = teacher_df.drop(columns=["上台次数（BI）_bi"])
    bi_alias = bi.rename(columns={"昵称": "昵称_bi"})
    teacher_df = teacher_df.merge(
        bi_alias[["昵称_alias", "大组", "小组", "上台次数（BI）"]].rename(columns={"上台次数（BI）": "上台次数（BI）_alias"}),
        on=["昵称_alias", "大组", "小组"],
        how="left",
    )
    teacher_df["上台次数（BI）"] = teacher_df["上台次数（BI）"].combine_first(teacher_df["上台次数（BI）_alias"])
    teacher_df = teacher_df.drop(columns=["上台次数（BI）_alias"])
    teacher_df = teacher_df.merge(
        bi_alias[["昵称_alias", "大组", "上台次数（BI）"]].rename(columns={"上台次数（BI）": "上台次数（BI）_dept_only"}).drop_duplicates(),
        on=["昵称_alias", "大组"],
        how="left",
    )
    teacher_df["上台次数（BI）"] = teacher_df["上台次数（BI）"].combine_first(teacher_df["上台次数（BI）_dept_only"])
    teacher_df = teacher_df.drop(columns=["上台次数（BI）_dept_only"])
    teacher_df = teacher_df.merge(
        bi_alias[["昵称_alias", "上台次数（BI）"]].rename(columns={"上台次数（BI）": "上台次数（BI）_name_only"}).drop_duplicates(),
        on=["昵称_alias"],
        how="left",
    )
    teacher_df["上台次数（BI）"] = teacher_df["上台次数（BI）"].combine_first(teacher_df["上台次数（BI）_name_only"])
    teacher_df = teacher_df.drop(columns=["上台次数（BI）_name_only"])
    teacher_df = teacher_df.merge(
        bi_alias[["昵称_alias_loose", "上台次数（BI）"]].rename(columns={"上台次数（BI）": "上台次数（BI）_loose"}).drop_duplicates(),
        on=["昵称_alias_loose"],
        how="left",
    )
    teacher_df["上台次数（BI）"] = teacher_df["上台次数（BI）"].combine_first(teacher_df["上台次数（BI）_loose"])
    teacher_df = teacher_df.drop(columns=["上台次数（BI）_loose", "昵称_alias", "昵称_alias_loose"])
    teacher_df = teacher_df.sort_values(["大组", "小组", "昵称"], kind="stable").reset_index(drop=True)
    teacher_df = teacher_df[TEACHER_OUTPUT_COLUMNS]
    return teacher_df.rename(columns={
        "背景布合格率": "背景布\n合格率",
        "工服合格率": "工服\n合格率",
        "光线合格率": "光线\n合格率",
        "妆发合格率": "妆发\n合格率",
        "坐姿合格率": "坐姿\n合格率",
        "行为合格率": "行为\n合格率",
        "感染力分数": "感染力\n分数",
        "开口时长（s）": "开口时长\n（s）",
    })


def build_sd_sheet(detail: pd.DataFrame, teacher_summary: pd.DataFrame) -> pd.DataFrame:
    detail = detail.copy()
    label_order = ["英语", "粤语", "台教", "外教"]
    team_rows: list[dict[str, object]] = []
    group_rows: list[dict[str, object]] = []

    for dept in ["海外益智英语教学区", "海外益智粤语教学区", "海外益智台湾教学区", "海外益智外教教学区"]:
        group = detail[detail["教师部门5"] == dept]
        if group.empty:
            continue
        team_rows.append({
            "团队": LABEL_MAP.get(dept, dept),
            "样本数": int(len(group)),
            "老师数": int(teacher_summary.loc[teacher_summary["大组"] == dept, "昵称"].nunique()),
            "感染力分数": average_metric(group, "感染力\n分数"),
            "声音洪亮分数": average_metric(group, "声音洪亮\n分数"),
            "知识点讲错": ratio_metric(group, "知识点讲错"),
            "知识点漏讲": ratio_metric(group, "知识点漏讲"),
            "课堂规则提及": ratio_metric(group, "课堂规则提及"),
            "开口时长": average_metric(group, "开口时长(s)"),
            "课中正确率": average_metric(group, "课中正确率"),
            "上台次数": average_metric(group, "上台平均数"),
            "学员正面动作次数": average_metric(group, "主动性"),
            "读题审题": ratio_metric(group, "读题审题"),
            "关键提问": ratio_metric(group, "关键提问"),
            "情绪策略": ratio_metric(group, "情绪策略"),
            "课程预告": ratio_metric(group, "课程预告"),
        })
        for team_name, subgroup in group.groupby("教师部门6", dropna=False):
            group_rows.append({
                "团队": team_name,
                "样本数": int(len(subgroup)),
                "老师数": int(teacher_summary.loc[(teacher_summary["大组"] == dept) & (teacher_summary["小组"] == team_name), "昵称"].nunique()),
                "感染力分数": average_metric(subgroup, "感染力\n分数"),
                "声音洪亮分数": average_metric(subgroup, "声音洪亮\n分数"),
                "知识点讲错": ratio_metric(subgroup, "知识点讲错"),
                "知识点漏讲": ratio_metric(subgroup, "知识点漏讲"),
                "课堂规则提及": ratio_metric(subgroup, "课堂规则提及"),
                "开口时长": average_metric(subgroup, "开口时长(s)"),
                "课中正确率": average_metric(subgroup, "课中正确率"),
                "上台次数": average_metric(subgroup, "上台平均数"),
                "学员正面动作次数": average_metric(subgroup, "主动性"),
                "读题审题": ratio_metric(subgroup, "读题审题"),
                "关键提问": ratio_metric(subgroup, "关键提问"),
                "情绪策略": ratio_metric(subgroup, "情绪策略"),
                "课程预告": ratio_metric(subgroup, "课程预告"),
            })

    summary_rows = []
    for row in team_rows:
        summary_rows.append({"周次": PERIOD_LABEL, "当周": None, "当月": None, **row})
    for row in group_rows:
        summary_rows.append({"周次": None, "当周": row["团队"], "当月": None, **row})
    return pd.DataFrame(summary_rows)


def load_standard_targets() -> dict[str, float | None]:
    df = pd.read_excel(STANDARD_PATH)
    metric_map = {
        "学员参与度（均值）": "上台次数（BI）",
        "热情洋溢感染力强": "感染力合格率",
        "有效互动": "有效互动频次",
        "背景布": "背景布合格率",
        "工服": "工服合格率",
        "光线镜头": "光线合格率",
        "妆发": "妆发合格率",
        "坐姿": "坐姿合格率",
        "行为举止": "行为合格率",
        "学习习惯检查": "课前作业提醒合格率",
        "布置作业": "课后作业布置合格率",
        "知识点讲错": "知识点讲错通过率",
        "知识点漏讲": "知识点漏讲通过率",
        "课堂规则": "课堂规则提及通过率",
        "学员开口时长（均值）": "开口时长（s）",
        "课堂正确率（均值）": "课中正确率",
        "主动性（均值 ）": "学员正面动作次数",
        "下节预告（动画+知识点）": "课程预告通过率",
    }
    targets: dict[str, float | None] = {}
    for _, row in df.iterrows():
        indicator = str(row.get("指标", "")).strip()
        target_col = metric_map.get(indicator)
        if not target_col:
            continue
        text = str(row.get("质检标准调整", "") or "").replace(" ", "")
        pct_match = re.search(r"A：.*?([0-9]+(?:\.[0-9]+)?)%", text)
        num_match = re.search(r"A：.*?([0-9]+(?:\.[0-9]+)?)", text)
        target = None
        if pct_match:
            target = float(pct_match.group(1)) / 100
        elif num_match:
            target = float(num_match.group(1))
        if "分钟" in text and target is not None:
            target *= 60
        targets[target_col] = target
    targets.update(MONTH_TARGETS)
    return targets


def build_month_sheet(teacher_summary: pd.DataFrame) -> pd.DataFrame:
    targets = load_standard_targets()
    rows = []
    for dept in ["海外益智英语教学区", "海外益智粤语教学区", "海外益智台湾教学区", "海外益智外教教学区"]:
        sub = teacher_summary[teacher_summary["大组"] == dept]
        if sub.empty:
            continue
        label = LABEL_MAP.get(dept, dept)
        for group_name, g in sub.groupby("小组", dropna=False):
            rows.append({
                "老师属性": label,
                "小组": group_name,
                "老师数": int(g["昵称"].nunique()),
                "上台次数（BI）": safe_float(pd.to_numeric(g["上台次数（BI）"], errors="coerce").mean()),
                "感染力合格率": safe_float(pd.to_numeric(g["感染力合格率"], errors="coerce").mean()),
                "有效互动频次": safe_float(pd.to_numeric(g["有效互动频次"], errors="coerce").mean()),
            })
    month = pd.DataFrame(rows)
    for metric, target in targets.items():
        if metric in month.columns:
            month[f"{metric}_目标值"] = target
            month[f"{metric}_GAP"] = pd.to_numeric(month[metric], errors="coerce") - target if target is not None else None
    return month


def build_sd_display(teacher_summary: pd.DataFrame, detail: pd.DataFrame, bi: pd.DataFrame) -> pd.DataFrame:
    label_order = ["英语", "粤语", "台教", "菲教", "督导", "职能岗"]
    team_dept_order = [
        "海外益智英语教学区",
        "海外益智粤语教学区",
        "海外益智台湾教学区",
        "海外益智外教教学区",
        "海外益智教学部",
        "海外益智教学运营部",
        "海外益智教学质量管理部",
    ]
    current_rows = []
    month_rows = []
    right_rows = []
    valid_detail = detail[detail["样本数"] == 1].copy()

    for dept in team_dept_order:
        label = LABEL_MAP.get(dept, "督导" if "教学部" in dept and "质量" not in dept and "运营" not in dept else ("职能岗" if "运营" in dept or "质量" in dept else dept))
        tsub = teacher_summary[teacher_summary["大组"] == dept].copy()
        if tsub.empty:
            continue
        bsub = bi[bi["大组"] == dept].copy()
        current_rows.append({"行标签": label, "上台次数 ": safe_float(pd.to_numeric(bsub["上台次数（BI）"], errors="coerce").mean()), "去重老师数": int(tsub["昵称"].nunique())})
        month_rows.append({"行标签": label, "上台次数 ": safe_float(pd.to_numeric(bsub["上台次数（BI）"], errors="coerce").mean()), "去重老师数": int(tsub["昵称"].nunique())})
        right_rows.append({
            "团队": label,
            "样本数 ": int(len(valid_detail[valid_detail["教师部门5"] == dept])),
            "老师数": int(tsub["昵称"].nunique()),
            "感染力分数": safe_float(pd.to_numeric(tsub["感染力\n分数"], errors="coerce").mean()),
            "声音洪亮分数 ": safe_float(pd.to_numeric(tsub["声音洪亮分数"], errors="coerce").mean()),
            "知识点讲错": safe_float(pd.to_numeric(tsub["知识点讲错通过率"], errors="coerce").mean()),
            "知识点漏讲": safe_float(pd.to_numeric(tsub["知识点漏讲通过率"], errors="coerce").mean()),
            "课堂规则提及": safe_float(pd.to_numeric(tsub["课堂规则提及通过率"], errors="coerce").mean()),
            "开口时长": safe_float(pd.to_numeric(tsub["开口时长\n（s）"], errors="coerce").mean()),
            "课中正确率 ": safe_float(pd.to_numeric(tsub["课中正确率"], errors="coerce").mean()),
            "上台次数": safe_float(pd.to_numeric(tsub["上台次数（BI）"], errors="coerce").mean()),
            "学员正面动作次数": safe_float(pd.to_numeric(tsub["学员正面动作次数"], errors="coerce").mean()),
            "读题审题 ": safe_float(pd.to_numeric(tsub["读题审题通过率"], errors="coerce").mean()),
            "关键提问 ": safe_float(pd.to_numeric(tsub["关键提问通过率"], errors="coerce").mean()),
            "情绪策略 ": safe_float(pd.to_numeric(tsub["情绪策略使用通过率"], errors="coerce").mean()),
            "课程预告 ": safe_float(pd.to_numeric(tsub["课程预告通过率"], errors="coerce").mean()),
        })
        for group_name, g in tsub.groupby("小组", dropna=False):
            group_label = "兼职" if str(group_name).strip() == "兼职老师组" else ("" if pd.isna(group_name) else group_name)
            gb = bi[(bi["大组"] == dept) & (bi["小组"] == group_name)].copy()
            current_rows.append({"行标签": group_label, "上台次数 ": safe_float(pd.to_numeric(gb["上台次数（BI）"], errors="coerce").mean()), "去重老师数": int(g["昵称"].nunique())})
            month_rows.append({"行标签": group_label, "上台次数 ": safe_float(pd.to_numeric(gb["上台次数（BI）"], errors="coerce").mean()), "去重老师数": int(g["昵称"].nunique())})
            gd = valid_detail[(valid_detail["教师部门5"] == dept) & (valid_detail["教师部门6"] == group_name)]
            right_rows.append({
                "团队": group_label,
                "样本数 ": int(len(gd)),
                "老师数": int(g["昵称"].nunique()),
                "感染力分数": safe_float(pd.to_numeric(g["感染力\n分数"], errors="coerce").mean()),
                "声音洪亮分数 ": safe_float(pd.to_numeric(g["声音洪亮分数"], errors="coerce").mean()),
                "知识点讲错": safe_float(pd.to_numeric(g["知识点讲错通过率"], errors="coerce").mean()),
                "知识点漏讲": safe_float(pd.to_numeric(g["知识点漏讲通过率"], errors="coerce").mean()),
                "课堂规则提及": safe_float(pd.to_numeric(g["课堂规则提及通过率"], errors="coerce").mean()),
                "开口时长": safe_float(pd.to_numeric(g["开口时长\n（s）"], errors="coerce").mean()),
                "课中正确率 ": safe_float(pd.to_numeric(g["课中正确率"], errors="coerce").mean()),
                "上台次数": safe_float(pd.to_numeric(g["上台次数（BI）"], errors="coerce").mean()),
                "学员正面动作次数": safe_float(pd.to_numeric(g["学员正面动作次数"], errors="coerce").mean()),
                "读题审题 ": safe_float(pd.to_numeric(g["读题审题通过率"], errors="coerce").mean()),
                "关键提问 ": safe_float(pd.to_numeric(g["关键提问通过率"], errors="coerce").mean()),
                "情绪策略 ": safe_float(pd.to_numeric(g["情绪策略使用通过率"], errors="coerce").mean()),
                "课程预告 ": safe_float(pd.to_numeric(g["课程预告通过率"], errors="coerce").mean()),
            })
    left = pd.DataFrame(current_rows)
    left_month = pd.DataFrame(month_rows)
    right = pd.DataFrame(right_rows)
    n = max(len(left), len(left_month), len(right))
    left = left.reindex(range(n))
    left_month = left_month.reindex(range(n))
    right = right.reindex(range(n))
    return pd.concat([left, left_month, right], axis=1)


def build_sd_sheet_rows(teacher_summary: pd.DataFrame, detail: pd.DataFrame, bi: pd.DataFrame) -> list[list[object]]:
    upper = build_sd_display(teacher_summary, detail, bi).fillna("")
    valid_detail = detail[detail["样本数"] == 1].copy()

    order_rows: list[tuple[str, str | None]] = []
    seen: set[tuple[str, str | None]] = set()
    for _, row in teacher_summary[["大组", "小组"]].drop_duplicates().iterrows():
        dept = row["大组"]
        group = row["小组"]
        if (dept, None) not in seen:
            seen.add((dept, None))
            order_rows.append((dept, None))
        key = (dept, group)
        if key not in seen:
            seen.add(key)
            order_rows.append(key)

    def dept_label(dept: str) -> str:
        if dept in LABEL_MAP:
            return LABEL_MAP[dept]
        if "质量" in str(dept) or "运营" in str(dept):
            return "职能岗"
        if "教学部" in str(dept):
            return "督导"
        return str(dept)

    lower_rows: list[dict[str, object]] = []
    for dept, group in order_rows:
        if group is None:
            sub = teacher_summary[teacher_summary["大组"] == dept].copy()
            dsub = valid_detail[valid_detail["教师部门5"] == dept].copy()
            label = dept_label(dept)
        else:
            sub = teacher_summary[(teacher_summary["大组"] == dept) & (teacher_summary["小组"] == group)].copy()
            dsub = valid_detail[(valid_detail["教师部门5"] == dept) & (valid_detail["教师部门6"] == group)].copy()
            label = "兼职" if str(group).strip() == "兼职老师组" else str(group)
        if sub.empty:
            continue
        lower_rows.append({
            "团队": label,
            "样本数 ": int(len(dsub)),
            "老师数": int(sub["昵称"].nunique()),
            "背景布合格率": safe_float(pd.to_numeric(sub["背景布\n合格率"], errors="coerce").mean()),
            "工服合格率": safe_float(pd.to_numeric(sub["工服\n合格率"], errors="coerce").mean()),
            "光线合格率": safe_float(pd.to_numeric(sub["光线\n合格率"], errors="coerce").mean()),
            "妆发合格率": safe_float(pd.to_numeric(sub["妆发\n合格率"], errors="coerce").mean()),
            "坐姿合格率": safe_float(pd.to_numeric(sub["坐姿\n合格率"], errors="coerce").mean()),
            "行为合格率": safe_float(pd.to_numeric(sub["行为\n合格率"], errors="coerce").mean()),
            "感染力合格率": safe_float(pd.to_numeric(sub["感染力合格率"], errors="coerce").mean()),
            "有效互动频次": safe_float(pd.to_numeric(sub["有效互动频次"], errors="coerce").mean()),
            "课前作业提醒合格率": safe_float(pd.to_numeric(sub["课前作业提醒合格率"], errors="coerce").mean()),
            "课后作业布置合格率": safe_float(pd.to_numeric(sub["课后作业布置合格率"], errors="coerce").mean()),
        })

    lower = pd.DataFrame(lower_rows, columns=SD_LOWER_COLUMNS).fillna("")

    rows: list[list[object]] = []
    rows.append(["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "周期", PERIOD_LABEL] + [""] * 14)
    rows.append(["周次", PERIOD_LABEL] + [""] * 30)
    rows.append(["当周", "", "", "", "当月", "", "", "", "", "", "", "", "", "", "", "", "团队", "样本数 ", "老师数", "感染力分数", "声音洪亮分数 ", "知识点讲错", "知识点漏讲", "课堂规则提及", "开口时长", "课中正确率 ", "上台次数", "学员正面动作次数", "读题审题 ", "关键提问 ", "情绪策略 ", "课程预告 "])
    for i in range(len(upper)):
        rows.append(upper.iloc[i].tolist())
    rows.append([""] * 16 + ["当月"] + [""] * 15)
    rows.append([""] * 16 + ["团队", "样本数 ", "老师数", "感染力分数", "声音洪亮分数 ", "知识点讲错", "知识点漏讲", "课堂规则提及", "开口时长", "课中正确率 ", "上台次数", "学员正面动作次数", "读题审题 ", "关键提问 ", "情绪策略 ", "课程预告 "])
    for i in range(len(upper)):
        block = [""] * 16 + upper.iloc[i, 6:].tolist()
        rows.append(block)
    rows.append(["团队", "样本数 ", "老师数", "背景布合格率", "工服合格率", "光线合格率", "妆发合格率", "坐姿合格率", "行为合格率", "感染力合格率", "有效互动频次", "课前作业提醒合格率", "课后作业布置合格率"] + [""] * 19)
    for i in range(len(lower)):
        rows.append(lower.iloc[i].tolist() + [""] * 19)
    rows.append(["当月"] + [""] * 31)
    rows.append(["团队", "样本数 ", "老师数", "背景布合格率", "工服合格率", "光线合格率", "妆发合格率", "坐姿合格率", "行为合格率", "感染力合格率", "有效互动频次", "课前作业提醒合格率", "课后作业布置合格率"] + [""] * 19)
    for i in range(len(lower)):
        rows.append(lower.iloc[i].tolist() + [""] * 19)
    return rows


def build_monthly_trend_display(teacher_summary: pd.DataFrame) -> pd.DataFrame:
    return pd.read_excel(REFERENCE_W5_PATH, sheet_name="月度汇总", header=None)


def build_monthly_trend_sheet_rows(month_display: pd.DataFrame) -> list[list[object]]:
    return month_display.fillna("").values.tolist()


def export_workbook(raw: pd.DataFrame, detail: pd.DataFrame, teacher_summary: pd.DataFrame, sd_display: pd.DataFrame, month_display: pd.DataFrame, bi: pd.DataFrame, effective: pd.DataFrame) -> Path:
    output_path = OUTPUT_DIR / "海外监课指标7月W1_重建版.xlsx"
    raw_export = sanitize_excel_text(raw)
    detail_export = sanitize_excel_text(detail)
    effective_export = sanitize_excel_text(effective)
    sd_rows = build_sd_sheet_rows(teacher_summary, detail, bi)
    with pd.ExcelWriter(
        output_path,
        engine="xlsxwriter",
        engine_kwargs={"options": {"strings_to_urls": False}},
    ) as writer:
        detail_export.to_excel(writer, sheet_name="质检明细", index=False)
        teacher_summary.to_excel(writer, sheet_name="W1老师达成情况", index=False, startrow=1)
        pd.DataFrame(sd_rows).to_excel(writer, sheet_name="数透", index=False, header=False)
        pd.DataFrame(build_monthly_trend_sheet_rows(month_display)).to_excel(writer, sheet_name="月度汇总", index=False, header=False)
        bi.to_excel(writer, sheet_name="BI上台次数", index=False)
        effective_export.to_excel(writer, sheet_name="有效课节数", index=False)
        raw_export.to_excel(writer, sheet_name="原始合并底表", index=False)
        workbook = writer.book
        fmt_header = workbook.add_format({"bold": True, "align": "center", "valign": "vcenter", "text_wrap": True, "bg_color": "#D9EAF7", "border": 1})
        fmt_sub = workbook.add_format({"bold": True, "align": "center", "valign": "vcenter", "text_wrap": True, "bg_color": "#FBE5D6", "border": 1})
        fmt_yellow = workbook.add_format({"bold": True, "align": "center", "valign": "vcenter", "text_wrap": True, "bg_color": "#FFF2CC", "border": 1})
        fmt_text = workbook.add_format({"align": "center", "valign": "vcenter", "text_wrap": True, "border": 1})
        fmt_num = workbook.add_format({"align": "center", "valign": "vcenter", "num_format": "0.00_);[Red]\\(0.00\\)", "border": 1})
        fmt_pct = workbook.add_format({"align": "center", "valign": "vcenter", "num_format": "0.00%", "border": 1})
        fmt_pct0 = workbook.add_format({"align": "center", "valign": "vcenter", "num_format": "0%", "border": 1})
        fmt_pct1 = workbook.add_format({"align": "center", "valign": "vcenter", "num_format": "0.0%", "border": 1})
        targets = load_standard_targets()

        def pick_teacher_fmt(name: str):
            if str(name) in TEACHER_PERCENT_2_COLUMNS:
                return fmt_pct
            if str(name) in TEACHER_PERCENT_0_COLUMNS:
                return fmt_pct0
            return fmt_num

        def pick_sd_fmt(name: str):
            return fmt_pct if str(name) in SD_PERCENT_2_COLUMNS else fmt_num

        # 质检明细
        ws = writer.sheets["质检明细"]
        ws.freeze_panes(1, 0)
        for c, name in enumerate(detail_export.columns):
            ws.write(0, c, name, fmt_yellow)
            ws.set_column(c, c, 14 if c <= 34 else 18, fmt_text)

        # 老师达成
        ws = writer.sheets["W1老师达成情况"]
        ws.freeze_panes(2, 0)
        for c, name in enumerate(teacher_summary.columns):
            ws.write(0, c, targets.get(name, "-"), fmt_sub)
            ws.write(1, c, name, fmt_header)
            ws.set_column(c, c, 14, fmt_text if c < 6 else pick_teacher_fmt(name))

        # 数透
        ws = writer.sheets["数透"]
        ws.freeze_panes(4, 0)
        for c in range(32):
            ws.set_column(c, c, 14, fmt_text if c in [0, 3, 6, 16] else fmt_num)
        for r, row in enumerate(sd_rows):
            for c, val in enumerate(row):
                if r in [0, 1]:
                    fmt = fmt_sub if val != "" else fmt_text
                elif r in [2, 31, 32, 60, 90]:
                    fmt = fmt_sub if val != "" else fmt_text
                elif r in [3, 33, 61, 91]:
                    fmt = fmt_header if val != "" else fmt_text
                else:
                    if c in [0, 16]:
                        fmt = fmt_text
                    elif c <= 12 and r >= 62:
                        if c in [3, 4, 5, 6, 7, 8, 9, 11, 12]:
                            fmt = fmt_pct0
                        else:
                            fmt = fmt_num
                    elif c >= 19:
                        fmt = fmt_pct if c in [19, 20, 21, 22, 23, 25, 28, 29, 30, 31] else fmt_num
                    else:
                        fmt = fmt_num
                ws.write(r, c, val, fmt)

        # 月度汇总
        ws = writer.sheets["月度汇总"]
        ws.freeze_panes(2, 0)
        month_rows = build_monthly_trend_sheet_rows(month_display)
        for c, val in enumerate(month_rows[0]):
            ws.write(0, c, val, fmt_header)
            if c < 2:
                col_fmt = fmt_text
            elif 11 <= c <= 19:
                col_fmt = fmt_pct1
            else:
                col_fmt = fmt_num
            ws.set_column(c, c, 14, col_fmt)
        for c, val in enumerate(month_rows[1]):
            ws.write(1, c, val, fmt_sub)

        # 其他sheet
        for name, df in [("BI上台次数", bi), ("有效课节数", effective_export), ("原始合并底表", raw_export)]:
            ws = writer.sheets[name]
            ws.freeze_panes(1, 0)
            for c, col in enumerate(df.columns):
                ws.write(0, c, col, fmt_header)
                ws.set_column(c, c, 16, fmt_text)

        # 条件格式
        # 老师达成：仪容仪表低于90标红；其余色阶
        teacher_col_map = {name: idx for idx, name in enumerate(teacher_summary.columns)}
        for name in TEACHER_GROOMING_COLUMNS:
            if name in teacher_col_map:
                c = teacher_col_map[name]
                ws = writer.sheets["W1老师达成情况"]
                ws.conditional_format(2, c, len(teacher_summary) + 1, c, {
                    "type": "cell",
                    "criteria": "<",
                    "value": 0.9,
                    "format": workbook.add_format({"bg_color": "#F4CCCC", "font_color": "#9C0006", "border": 1}),
                })
        ws = writer.sheets["W1老师达成情况"]
        for name, c in teacher_col_map.items():
            if name in TEACHER_SCALE_COLUMNS:
                ws.conditional_format(2, c, len(teacher_summary) + 1, c, {
                    "type": "3_color_scale",
                    "min_color": "#F8696B",
                    "mid_color": "#FFEB84",
                    "max_color": "#63BE7B",
                })

        ws = writer.sheets["数透"]
        for c in range(19, 32):
            ws.conditional_format(4, c, 31, c, {
                "type": "3_color_scale",
                "min_color": "#F8696B",
                "mid_color": "#FFEB84",
                "max_color": "#63BE7B",
            })
            ws.conditional_format(33, c, 60, c, {
                "type": "3_color_scale",
                "min_color": "#F8696B",
                "mid_color": "#FFEB84",
                "max_color": "#63BE7B",
            })
        for c in range(3, 13):
            ws.conditional_format(62, c, 89, c, {
                "type": "3_color_scale",
                "min_color": "#F8696B",
                "mid_color": "#FFEB84",
                "max_color": "#63BE7B",
            })
            ws.conditional_format(92, c, len(sd_rows) - 1, c, {
                "type": "3_color_scale",
                "min_color": "#F8696B",
                "mid_color": "#FFEB84",
                "max_color": "#63BE7B",
            })

        ws = writer.sheets["月度汇总"]
        month_last_row = len(month_rows)
        for c in range(2, len(month_rows[0])):
            ws.conditional_format(2, c, month_last_row - 1, c, {
                "type": "3_color_scale",
                "min_color": "#F8696B",
                "mid_color": "#FFEB84",
                "max_color": "#63BE7B",
            })
    return output_path


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    raw = load_raw()
    detail = build_detail(raw)
    bi = load_bi()
    effective = pd.read_excel(EFFECTIVE_PATH)
    teacher_summary = build_teacher_summary(detail, bi)
    sd_summary = build_sd_display(teacher_summary, detail, bi)
    month_df = build_monthly_trend_display(teacher_summary)

    merged_raw_path = OUTPUT_DIR / "质检记录列表7.1-7.7_合并.xlsx"
    raw.to_excel(merged_raw_path, index=False)
    detail.to_csv(OUTPUT_DIR / "7w1_质检明细.csv", index=False, encoding="utf-8-sig")
    teacher_summary.to_csv(OUTPUT_DIR / "7w1_老师达成情况.csv", index=False, encoding="utf-8-sig")
    sd_summary.to_csv(OUTPUT_DIR / "7w1_数透.csv", index=False, encoding="utf-8-sig")
    month_df.to_csv(OUTPUT_DIR / "7w1_月度汇总.csv", index=False, encoding="utf-8-sig")
    workbook_path = export_workbook(raw, detail, teacher_summary, sd_summary, month_df, bi, effective)

    print(merged_raw_path)
    print(workbook_path)


if __name__ == "__main__":
    main()
