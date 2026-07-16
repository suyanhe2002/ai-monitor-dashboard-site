import fs from "node:fs/promises";
import path from "node:path";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const sourceHtmlPath = "/Users/lilblackmac/Documents/New project/tools/ai_monitor_dashboard_v1/ai_monitor_dashboard_v1.html";
const outputPath = "/Users/lilblackmac/Documents/New project/share/AI监课小组风险老师跟进套表.xlsx";

function extractPayload(html) {
  const match = html.match(/<script id="payload-data" type="application\/json">([\s\S]*?)<\/script>/);
  if (!match) throw new Error("payload-data script not found");
  return JSON.parse(match[1]);
}

function pct(value) {
  if (value == null || Number.isNaN(value)) return "";
  return value;
}

function riskOrder(level) {
  if (level === "高") return 0;
  if (level === "中") return 1;
  return 2;
}

function isRealGroup(name) {
  if (!name) return false;
  const text = String(name).trim();
  if (!text || text === "未分组") return false;
  if (text === "质量培训组") return false;
  if (text.endsWith("组")) return true;
  return false;
}

function naturalGroupSort(name) {
  const match = String(name).match(/^(.*?)(\d+)组$/);
  if (match) return [match[1], Number(match[2])];
  return [String(name), 0];
}

const html = await fs.readFile(sourceHtmlPath, "utf8");
const payload = extractPayload(html);
const rows = (payload.teachers || [])
  .filter((row) => (row["风险等级"] === "高" || row["风险等级"] === "中") && isRealGroup(row["小组"]))
  .sort((a, b) => {
    const groupA = naturalGroupSort(a["小组"]);
    const groupB = naturalGroupSort(b["小组"]);
    if (groupA[0] !== groupB[0]) return String(groupA[0]).localeCompare(String(groupB[0]), "zh-Hans-CN");
    if (groupA[1] !== groupB[1]) return groupA[1] - groupB[1];
    const levelCmp = riskOrder(a["风险等级"]) - riskOrder(b["风险等级"]);
    if (levelCmp !== 0) return levelCmp;
    return (b["风险分"] || 0) - (a["风险分"] || 0);
  });

const headers = [
  "序号",
  "业务部门",
  "小组",
  "小组跟进人",
  "老师",
  "老师工号",
  "风险等级",
  "风险分",
  "样本课节数",
  "Top1抓手",
  "Top1跟进策略",
  "Top2抓手",
  "Top2跟进策略",
  "Top3抓手",
  "Top3跟进策略",
  "跟进结果/复盘",
];

function buildBodyRows(sourceRows) {
  return sourceRows.map((row, index) => {
    const issues = row["重点问题明细"] || [];
    return [
      index + 1,
      row["业务部门"] || "",
      row["小组"] || "",
      "",
      row["老师"] || "",
      row["老师工号"] || "",
      row["风险等级"] || "",
      row["风险分"] ?? null,
      row["样本课节数_W2"] ?? null,
      issues[0]?.metric || "",
      "",
      issues[1]?.metric || "",
      "",
      issues[2]?.metric || "",
      "",
      "",
    ];
  });
}

function formatSheet(sheet, title, subtitle, body) {
  sheet.showGridLines = false;
  sheet.getRange("A1:P1").merge();
  sheet.getRange("A1").values = [[title]];
  sheet.getRange("A2:P2").merge();
  sheet.getRange("A2").values = [[subtitle]];
  sheet.getRange("A4:P4").values = [headers];
  if (body.length) {
    sheet.getRange(`A5:P${4 + body.length}`).values = body;
  }

  sheet.getRange("A1:P1").format = {
    fill: { color: "#16324F" },
    font: { color: "#FFFFFF", bold: true, size: 16 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };

  sheet.getRange("A2:P2").format = {
    fill: { color: "#F7D8C8" },
    font: { color: "#16324F", size: 11 },
    horizontalAlignment: "left",
    verticalAlignment: "center",
  };

  sheet.getRange("A4:P4").format = {
    fill: { color: "#CB5F2B" },
    font: { color: "#FFFFFF", bold: true },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: "#E7CFC2" },
  };

  if (body.length) {
    const dataRange = sheet.getRange(`A5:P${4 + body.length}`);
    dataRange.format = {
      fill: { color: "#FFFDF9" },
      font: { color: "#16324F", size: 10 },
      verticalAlignment: "center",
      wrapText: true,
      borders: { preset: "all", style: "thin", color: "#E9E2D8" },
    };

    sheet.getRange(`G5:G${4 + body.length}`).format = {
      fill: { color: "#FFF4E8" },
      font: { bold: true, color: "#16324F" },
      horizontalAlignment: "center",
      verticalAlignment: "center",
      borders: { preset: "all", style: "thin", color: "#E9E2D8" },
    };

    sheet.getRange(`H5:H${4 + body.length}`).format.numberFormat = "0.0";
    sheet.getRange(`I5:I${4 + body.length}`).format.numberFormat = "0";
  }

  sheet.getRange("J4:O4").format = {
    fill: { color: "#8C5A2B" },
    font: { color: "#FFFFFF", bold: true },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: "#E7CFC2" },
  };

  sheet.getRange("D4:D4").format = {
    fill: { color: "#6B7A90" },
    font: { color: "#FFFFFF", bold: true },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: "#E7CFC2" },
  };

  sheet.freezePanes.freezeRows(4);
  sheet.getRange("A:P").format.columnWidth = 14;
  sheet.getRange("A:A").format.columnWidth = 8;
  sheet.getRange("B:C").format.columnWidth = 14;
  sheet.getRange("D:D").format.columnWidth = 14;
  sheet.getRange("E:E").format.columnWidth = 14;
  sheet.getRange("F:F").format.columnWidth = 12;
  sheet.getRange("G:G").format.columnWidth = 10;
  sheet.getRange("H:I").format.columnWidth = 11;
  sheet.getRange("J:O").format.columnWidth = 18;
  sheet.getRange("P:P").format.columnWidth = 24;
  sheet.getRange("1:2").format.rowHeight = 24;
  sheet.getRange("4:4").format.rowHeight = 26;
}

const workbook = Workbook.create();
const summarySheet = workbook.worksheets.add("总表");
formatSheet(
  summarySheet,
  "AI监课小组风险老师跟进套表",
  `数据来源：当前AI监课面板 | 记录范围：在小组内的高风险 + 中风险老师 | 总人数：${rows.length}`,
  buildBodyRows(rows),
);

const groupNames = [...new Set(rows.map((row) => row["小组"]))];
for (const groupName of groupNames) {
  const groupRows = rows.filter((row) => row["小组"] === groupName);
  const sheet = workbook.worksheets.add(groupName);
  formatSheet(
    sheet,
    `${groupName}风险老师跟进表`,
    `仅展示 ${groupName} 内高风险 + 中风险老师 | 人数：${groupRows.length}`,
    buildBodyRows(groupRows),
  );
}

await fs.mkdir(path.dirname(outputPath), { recursive: true });
const out = await SpreadsheetFile.exportXlsx(workbook);
await out.save(outputPath);
console.log(outputPath);
