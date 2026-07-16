from __future__ import annotations

import base64
import html
from pathlib import Path

import pandas as pd


ROOT = Path("/Users/lilblackmac/Documents/New project/tools/ai_monitor_dashboard_v1")
SOURCE = Path("/Users/lilblackmac/Desktop/AI监课&北极星/AI监课面板统计口径与判定逻辑总表_V5_详细版.xlsx")
OUTPUT = ROOT / "ai_monitor_logic_reference.html"
LOGO_PATH = ROOT / "assets" / "company_logo.jpg"


def safe_text(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return html.escape(str(value))


def read_sheet(sheet_name: str) -> pd.DataFrame:
    raw = pd.read_excel(SOURCE, sheet_name=sheet_name, header=None)
    header = [str(x).strip() if not pd.isna(x) else f"col_{i}" for i, x in enumerate(raw.iloc[0].tolist())]
    df = raw.iloc[1:].copy()
    df.columns = header
    df = df.dropna(how="all")
    return df.fillna("")


def table_html(df: pd.DataFrame, compact: bool = False) -> str:
    classes = "matrix compact" if compact else "matrix"
    head = "".join(f"<th>{safe_text(col)}</th>" for col in df.columns)
    body_rows = []
    for _, row in df.iterrows():
        cells = "".join(f"<td>{safe_text(val)}</td>" for val in row.tolist())
        body_rows.append(f"<tr>{cells}</tr>")
    body = "".join(body_rows)
    return f"<table class='{classes}'><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table>"


def build_cards(df: pd.DataFrame) -> str:
    cards = []
    for _, row in df.iterrows():
        title = safe_text(row.iloc[0])
        value = safe_text(row.iloc[1])
        cards.append(
            f"""
            <article class="info-card">
              <div class="info-title">{title}</div>
              <div class="info-value">{value}</div>
            </article>
            """
        )
    return "".join(cards)


def build_html() -> str:
    overview = read_sheet("01_总览")
    risk_rules = read_sheet("02_风险判定")
    metric_detail = read_sheet("03_指标详表")
    data_source = read_sheet("04_数据来源")
    management = read_sheet("05_管理解释")
    logo_src = f"data:image/jpeg;base64,{base64.b64encode(LOGO_PATH.read_bytes()).decode('ascii')}"

    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>AI监课面板口径说明</title>
  <style>
    :root {{
      --bg: #f4efe7;
      --panel: #fffaf3;
      --card: #fff;
      --ink: #16324f;
      --muted: #6a7788;
      --line: #dde4ee;
      --accent: #cb5f2b;
      --accent-soft: #f7d8c8;
      --shadow: 0 18px 40px rgba(22,50,79,.08);
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: "PingFang SC","Microsoft YaHei",sans-serif;
      color: var(--ink);
      background:
        radial-gradient(circle at top left, rgba(203,95,43,.16), transparent 24%),
        linear-gradient(180deg, #f8f4ed 0%, var(--bg) 100%);
    }}
    .shell {{ max-width: 1480px; margin: 0 auto; padding: 28px 24px 60px; }}
    .hero, .section, .info-card {{
      background: linear-gradient(160deg, rgba(255,250,243,.95), rgba(255,255,255,.92));
      border: 1px solid rgba(203,95,43,.14);
      border-radius: 24px;
      box-shadow: var(--shadow);
    }}
    .hero {{ padding: 24px; margin-bottom: 20px; }}
    .brand {{ display:flex; align-items:center; gap:16px; flex-wrap:wrap; }}
    .brand img {{ height: 40px; width:auto; display:block; }}
    .eyebrow {{
      display:inline-block; padding:5px 10px; border-radius:999px;
      background: rgba(22,50,79,.08); font-size:11px; font-weight:700; letter-spacing:.06em;
    }}
    h1 {{ margin: 8px 0 10px; font-size: 34px; line-height: 1.1; }}
    p {{ margin: 0; color: var(--muted); line-height: 1.7; }}
    .tag-row {{ display:flex; gap:8px; flex-wrap:wrap; margin-top:16px; }}
    .tag {{ padding:7px 12px; border-radius:999px; background: var(--accent-soft); color: var(--accent); font-size:12px; font-weight:700; }}
    .section {{ padding: 20px; margin-bottom: 18px; }}
    .section h2 {{ margin: 0 0 6px; font-size: 22px; }}
    .section-intro {{ margin-bottom: 14px; }}
    .card-grid {{
      display:grid; grid-template-columns: repeat(2, minmax(0, 1fr));
      gap:14px;
    }}
    .info-card {{ padding: 16px; }}
    .info-title {{ font-size: 12px; color: var(--muted); margin-bottom: 8px; }}
    .info-value {{ font-size: 15px; line-height: 1.7; white-space: pre-wrap; }}
    .table-wrap {{
      overflow:auto; border:1px solid var(--line); border-radius:18px; background:#fff;
    }}
    .matrix {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
    .matrix th, .matrix td {{
      padding: 12px 12px; border-bottom: 1px solid #eef2f6; text-align: left; vertical-align: top; white-space: normal;
    }}
    .matrix th {{
      position: sticky; top: 0; background: #f8fafc; z-index: 1; color: var(--muted); font-weight: 700;
    }}
    .matrix.compact th, .matrix.compact td {{ padding: 10px 10px; font-size: 12px; }}
    .note {{
      margin-top: 12px; padding: 12px 14px; border-radius: 14px;
      background: #f8f4ee; color: var(--muted); font-size: 12px; line-height: 1.7;
    }}
    @media (max-width: 1120px) {{
      .card-grid {{ grid-template-columns: 1fr; }}
    }}
  </style>
</head>
<body>
  <div class="shell">
    <section class="hero">
      <div class="brand">
        <img src="{logo_src}" alt="VIPTHINK">
        <div>
          <span class="eyebrow">VIPTHINK · AI Monitor Logic</span>
          <h1>AI监课面板口径说明</h1>
          <p>这是一份给业务部门配套使用的说明面板，用来解释主面板里的统计口径、风险判定逻辑、指标范围和管理解释，方便主管、组长和中台统一理解。</p>
        </div>
      </div>
      <div class="tag-row">
        <span class="tag">总览</span>
        <span class="tag">风险判定</span>
        <span class="tag">指标详表</span>
        <span class="tag">数据来源</span>
        <span class="tag">管理解释</span>
      </div>
    </section>

    <section class="section">
      <h2>总览</h2>
      <p class="section-intro">先快速看这版主面板当前采用的最终口径，包括风险范围、目标值来源优先级和业务口径补充。</p>
      <div class="card-grid">{build_cards(overview)}</div>
    </section>

    <section class="section">
      <h2>风险判定</h2>
      <p class="section-intro">业务和组长最关心的部分是风险到底怎么算，这里把步骤顺序和详细说明都摊开。</p>
      <div class="table-wrap">{table_html(risk_rules, compact=True)}</div>
    </section>

    <section class="section">
      <h2>指标详表</h2>
      <p class="section-intro">这一页回答每个指标是否参与风险、目标值从哪里来、现在按什么方式判定、权重是多少。</p>
      <div class="table-wrap">{table_html(metric_detail)}</div>
      <div class="note">如果后面面板再加新指标或调整目标值，这一块也应该同步更新，不然业务部门会看到“面板结果”和“解释口径”不一致。</div>
    </section>

    <section class="section">
      <h2>数据来源</h2>
      <p class="section-intro">这一页帮助大家快速对应“这个数到底来自哪张表”，方便中台和业务一起核对。</p>
      <div class="table-wrap">{table_html(data_source, compact=True)}</div>
    </section>

    <section class="section">
      <h2>管理解释</h2>
      <p class="section-intro">这里把高风险、中风险、低风险、Top抓手这些业务上常说的话，翻成更统一的管理解释。</p>
      <div class="card-grid">{build_cards(management)}</div>
    </section>
  </div>
</body>
</html>"""


def main() -> None:
    OUTPUT.write_text(build_html(), encoding="utf-8")
    print(OUTPUT)


if __name__ == "__main__":
    main()
