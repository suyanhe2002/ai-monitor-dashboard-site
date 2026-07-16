from __future__ import annotations

from pathlib import Path
import pandas as pd


ROOT = Path("/Users/lilblackmac/Documents/New project/tools/ai_monitor_dashboard_v1")
OUTPUT_DIR = ROOT / "normalized_outputs"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

NS_FILES = {
    "W1": Path("/Users/lilblackmac/Desktop/AI监课&北极星/北极星指标-海外团队6月W1（截至0604）.xlsx"),
    "W2": Path("/Users/lilblackmac/Desktop/AI监课&北极星/北极星指标-海外团队6月W2（截至0611）.xlsx"),
}

TARGET_METRICS = [
    "课中首答正确率",
    "课中末答正确率",
    "课后首答正确率",
    "课后末答正确率",
    "作业完成率",
    "预习率",
    "小老师视频提交率",
    "出勤率",
]


def safe_float(value):
    if pd.isna(value):
        return None
    try:
        return float(value)
    except Exception:
        return None


def build_stage_metric_actual_map(raw: pd.DataFrame) -> dict[str, dict[str, int]]:
    stage_row = raw.iloc[1]
    metric_row = raw.iloc[2]
    value_type_row = raw.iloc[3]
    result: dict[str, dict[str, int]] = {}
    current_stage = None
    for col_idx in range(len(raw.columns)):
        stage_value = stage_row.iloc[col_idx]
        if pd.notna(stage_value):
            current_stage = str(stage_value).strip()
        metric_value = metric_row.iloc[col_idx]
        value_type = value_type_row.iloc[col_idx]
        if not current_stage or pd.isna(metric_value):
            continue
        metric_name = str(metric_value).strip()
        value_type_text = "" if pd.isna(value_type) else str(value_type).strip()
        if value_type_text == "实际值":
            result.setdefault(current_stage, {})[metric_name] = col_idx
        # Some blocks place the metric label in one column and the actual value in the next column.
        elif col_idx + 1 < len(raw.columns):
            next_value_type = value_type_row.iloc[col_idx + 1]
            next_value_type_text = "" if pd.isna(next_value_type) else str(next_value_type).strip()
            if next_value_type_text == "实际值":
                result.setdefault(current_stage, {})[metric_name] = col_idx + 1
    return result


def normalize_week(week: str) -> pd.DataFrame:
    raw = pd.read_excel(NS_FILES[week], sheet_name="老师达成情况", header=None)
    stage_metric_actual_map = build_stage_metric_actual_map(raw)
    df = pd.read_excel(NS_FILES[week], sheet_name="老师达成情况", header=2)
    df = df[df["老师艺名"].notna()].copy()
    df = df[df["老师艺名"].astype(str).str.strip() != ""].copy()
    df = df[df["老师艺名"].astype(str).str.strip() != "老师艺名"].copy()

    records = []
    for _, row in df.iterrows():
        stage = None if pd.isna(row["课堂主学段"]) else str(row["课堂主学段"]).strip()
        mapping = stage_metric_actual_map.get(stage, {})
        record = {
            "周次": week,
            "课堂主学段": stage,
            "老师ID": None if pd.isna(row["老师ID"]) else str(row["老师ID"]).strip(),
            "老师艺名": None if pd.isna(row["老师艺名"]) else str(row["老师艺名"]).strip(),
            "教学老师五级部门": None if pd.isna(row["教学老师五级部门"]) else str(row["教学老师五级部门"]).strip(),
            "教学老师六级部门": None if pd.isna(row["教学老师六级部门"]) else str(row["教学老师六级部门"]).strip(),
        }
        for metric in TARGET_METRICS:
            col_idx = mapping.get(metric)
            record[metric] = None if col_idx is None else safe_float(pd.to_numeric(row.iloc[col_idx], errors="coerce"))
        records.append(record)
    return pd.DataFrame(records)


def main():
    for week in NS_FILES:
        df = normalize_week(week)
        out = OUTPUT_DIR / f"ns_teacher_normalized_{week}.csv"
        df.to_csv(out, index=False, encoding="utf-8-sig")
        print(out)


if __name__ == "__main__":
    main()
