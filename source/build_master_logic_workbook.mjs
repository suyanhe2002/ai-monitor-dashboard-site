import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = "/Users/lilblackmac/Documents/New project";
const outputDir = path.join(root, "outputs");
const outputPath = path.join(outputDir, "AI监课面板长期逻辑说明主文件_新版样式.xlsx");

const palette = {
  navy: "#17324D",
  blue: "#DCEAF7",
  blue2: "#EEF5FB",
  sand: "#F7F1E3",
  green: "#E8F3EC",
  red: "#F8E7E7",
  text: "#1F2937",
  white: "#FFFFFF",
  border: "#C9D6E2",
  accent: "#B7791F",
};

const sheets = [
  {
    name: "01_总览",
    title: "AI监课面板长期逻辑说明主文件",
    note: "本文件沉淀长期稳定逻辑，不绑定某一具体周次。后续仅在底层逻辑变化时更新；单周数据切换建议记录到每周更新页。",
    widths: [18, 24, 76],
    data: [
      ["模块", "主题", "长期口径说明"],
      ["定位", "文件用途", "用于沉淀 AI 监课面板长期稳定逻辑，包括指标分类、目标值逻辑、风险分逻辑、风险等级逻辑、Top抓手逻辑、老师分组逻辑、北极星辅助逻辑等。"],
      ["适用", "适用场景", "解释面板长期统计口径；对外说明面板底层规则；作为后续每周数据更新时的主依据。"],
      ["不适用", "不适用场景", "不用于记录某一具体周次的文件名、样本数量、高风险人数、临时阈值微调等动态内容。"],
      ["结构", "面板三层", "首页总览层、部门/小组层、老师下钻层。"],
      ["总逻辑", "展示层", "统一展示当前周、上一周、变化、目标、Gap。"],
      ["总逻辑", "判定层", "按标准分档 / 目标差距 / 周度变化三套逻辑并行判定。"],
      ["总逻辑", "风险层", "先计算老师风险分，再按风险分进行风险等级分层。"],
    ],
  },
  {
    name: "02_数据与分组",
    title: "数据来源与老师分组逻辑",
    note: "本页只写长期稳定的数据字段映射、聚合方式和组织分组逻辑，不写某一周的具体文件名版本。",
    widths: [18, 24, 28, 64],
    data: [
      ["模块", "对象", "字段 / 来源", "长期口径说明"],
      ["AI监课", "老师唯一粒度", "教师艺名 + 教师工号 + 教师部门5 + 教师部门6", "当前老师层聚合唯一粒度固定为老师姓名、工号、业务部门、小组四项组合。"],
      ["AI监课", "业务部门", "教师部门5", "映射为面板中的业务部门。"],
      ["AI监课", "小组", "教师部门6", "映射为面板中的小组。"],
      ["AI监课", "老师层指标", "质检明细", "每个 AI 监课指标按该老师该周全部样本课节取均值。"],
      ["AI监课", "样本课节数", "id", "按周对该老师 id 计数，作为样本课节数。"],
      ["部门/小组", "部门层聚合", "业务部门", "按业务部门聚合，风险分为老师层风险分均值。"],
      ["部门/小组", "小组层聚合", "业务部门 + 小组", "按业务部门和小组聚合，风险分为老师层风险分均值。"],
      ["缺失值", "空部门 / 空小组", "未分组", "若源数据缺失组织字段，则统一补为未分组。"],
      ["排序", "小组展示", "自然排序", "如 01组、02组、10组 按数字顺序展示。"],
      ["北极星", "接入方式", "老师达成情况", "北极星指标按老师主学段匹配正确列，只用于老师下钻辅助观察。"],
      ["北极星", "是否进风险分", "否", "北极星结果当前不参与风险分计算。"],
    ],
  },
  {
    name: "03_指标分类",
    title: "指标分类与目标值逻辑",
    note: "本页同时说明面板中的长期分类结构、哪些指标长期设目标值、哪些指标当前只看周变化。",
    widths: [18, 24, 16, 58],
    data: [
      ["一级分类", "指标", "是否设目标", "说明"],
      ["课堂质量核心", "感染力分数", "是", "核心风险指标，长期设目标值。"],
      ["课堂质量核心", "开口时长", "是", "标准表分钟口径统一换算成秒后参与计算。"],
      ["课堂质量核心", "课中正确率", "是", "核心风险指标，长期设目标值。"],
      ["课堂质量核心", "上台平均数", "是", "核心风险指标，长期设目标值。"],
      ["课堂质量核心", "有效互动次数", "是", "核心风险指标，长期设目标值。"],
      ["课堂质量核心", "镜头感", "是", "核心风险指标，长期设目标值。"],
      ["课堂质量核心", "愉悦度", "是", "核心风险指标，长期设目标值。"],
      ["教学动作", "主动性", "是", "展示指标，长期设目标值。"],
      ["教学动作", "读题审题", "否", "当前不通过目标值判断，只看周变化。"],
      ["教学动作", "关键提问", "否", "当前不通过目标值判断，只看周变化。"],
      ["教学动作", "情绪策略", "否", "当前不通过目标值判断，只看周变化。"],
      ["课前课后链路", "课前作业提醒", "否", "当前不通过目标值判断，只看周变化。"],
      ["课前课后链路", "课后作业布置", "否", "当前不通过目标值判断，只看周变化。"],
      ["课前课后链路", "课程预告", "否", "当前不通过目标值判断，只看周变化。"],
      ["教学风险", "知识点讲错通过率", "否", "当前不通过目标值判断，只看周变化。"],
      ["教学风险", "知识点漏讲通过率", "否", "当前不通过目标值判断，只看周变化。"],
      ["基础规范", "课堂规则提及通过率", "否", "当前不通过目标值判断，只看周变化。"],
      ["基础规范", "背景布 / 工服 / 光线 / 妆发 / 坐姿 / 行为", "是", "基础规范类合格率长期设目标值。"],
    ],
  },
  {
    name: "04_判定逻辑",
    title: "指标判定逻辑与 Gap 逻辑",
    note: "当前面板不是单一判定逻辑，而是三套并行逻辑。Gap 只用于展示，不等于风险分。",
    widths: [18, 24, 66],
    data: [
      ["逻辑类型", "适用对象", "长期口径说明"],
      ["标准分档", "有 A/B/C/D 的指标", "优先按标准表 A/B/C/D 分档。A 不计风险，B=轻度1分，C=中度2分，D=重度3分。卡片展示继续显示 A/B/C/D。"],
      ["目标差距", "无分档但有目标值", "达到目标不计风险；低于目标按差距大小分为轻度 / 中度 / 重度。百分比类看绝对百分点差距，数值类看相对目标差距比例。"],
      ["周度退步", "无分档且无目标值", "当前周未低于上一周则不计风险；若退步则按退步幅度分为轻度 / 中度 / 重度。百分比类看绝对百分点退步，数值类看相对退步比例。"],
      ["Gap", "所有有目标值指标", "Gap = 当前周值 - 目标值。Gap >= 0 代表达到或超过目标；Gap < 0 代表低于目标；无目标值指标 Gap 显示为 -。"],
      ["关键说明", "风险识别", "Gap 只是展示字段，不直接等于风险分。风险分必须结合三套判定逻辑综合计算。"],
    ],
  },
  {
    name: "05_风险分层",
    title: "风险分、风险变化与风险等级逻辑",
    note: "本页只写长期稳定的分值结构与分层方式；具体某一周的阈值如有微调，可另外写入每周更新记录。",
    widths: [18, 26, 64],
    data: [
      ["项目", "长期口径", "说明"],
      ["核心指标池", "固定核心指标", "感染力分数、开口时长、课中正确率、上台平均数、有效互动次数、镜头感、愉悦度、课前作业提醒、课后作业布置、课程预告、读题审题、关键提问、情绪策略、知识点讲错通过率、知识点漏讲通过率。"],
      ["风险分公式", "重度×3 + 中度×2 + 轻度×1", "每个核心指标先判成无风险 / 轻度 / 中度 / 重度，再累计成老师风险分。"],
      ["风险分含义", "累计分", "分数越高，说明问题越多、问题越重。"],
      ["风险变化", "当前周风险分 - 上一周基线风险分", "大于0代表风险变差，小于0代表风险改善。"],
      ["风险等级", "按风险分分层", "风险等级长期逻辑固定为“按风险分分层”，而不是单独再跑一套风险标签规则。"],
      ["阈值管理", "允许阶段性校准", "高 / 中 / 低的具体分界值可随真实样本分布校准，但“按风险分分层”的长期逻辑不变。"],
    ],
  },
  {
    name: "06_抓手与北极星",
    title: "Top抓手、老师下钻与北极星辅助逻辑",
    note: "本页沉淀的是抓手排序与老师下钻的长期逻辑，不记录某一具体周的结果。",
    widths: [18, 22, 66],
    data: [
      ["模块", "对象", "长期口径说明"],
      ["Top抓手", "首页 / 部门 / 小组", "不再展示单一混合 Top3，而是按指标分类分别展示 Top 抓手。"],
      ["Top抓手", "老师层", "展示当前老师优先级最高的 3 个重点问题。"],
      ["抓手排序", "排序依据", "先看问题分值 3/2/1，再看同分值下的严重度。"],
      ["老师下钻", "重点问题明细", "展示当前周值、上一周值、目标值、Gap、判定依据。"],
      ["北极星", "定位", "只用于老师下钻时的结果侧辅助观察。"],
      ["北极星", "展示内容", "课中首答正确率、课中末答正确率、课后首答正确率、课后末答正确率、作业完成率、预习率、小老师视频提交率。"],
      ["北极星", "是否参与风险分", "否，不参与当前风险分。"],
    ],
  },
  {
    name: "07_更新建议",
    title: "主文件与每周更新记录的配套建议",
    note: "推荐把长期主文件和每周更新记录拆开维护，这样以后更新数据时不用每周重写整份说明。",
    widths: [20, 24, 62],
    data: [
      ["文件类型", "建议用途", "说明"],
      ["长期逻辑主文件", "沉淀稳定规则", "用于记录长期稳定的面板结构、指标分类、判定逻辑、风险分逻辑、分组逻辑等。只有底层逻辑变化时才更新。"],
      ["每周更新记录", "记录周次变化", "用于记录当前周 / 上一周分别是哪一周、使用了哪些数据文件、当前风险等级阈值、是否有临时调整。"],
      ["什么时候不用改主文件", "只换周数据", "若只是把当前周和上一周的数据文件切换掉，且底层逻辑不变，则无需改长期逻辑主文件。"],
      ["什么时候要改主文件", "逻辑变化", "若风险分公式、风险等级分层方式、哪些指标设目标值、Top抓手逻辑、北极星是否进风险分等发生变化，则需要更新主文件。"],
      ["推荐总口径", "统一对外表述", "当前 AI监课面板采用“展示层看当前周/上一周/变化/目标/Gap，判定层按标准分档/目标差距/周度变化三套逻辑并行，老师风险分按轻中重问题累计计分，风险等级按风险分分层”的统一长期逻辑。"],
    ],
  },
];

function setCellFill(range, color) {
  const current = range.format || {};
  range.format = { ...current, fill: color };
}

function setBorders(range) {
  range.format.borders = { preset: "all", style: "thin", color: palette.border };
}

function styleWorkbookSheet(sheetDef, sheet) {
  const cols = sheetDef.data[0].length;
  const lastCol = String.fromCharCode(64 + cols);
  const titleRange = sheet.getRange(`A1:${lastCol}1`);
  const noteRange = sheet.getRange(`A2:${lastCol}2`);
  const headerRange = sheet.getRange(`A4:${lastCol}4`);
  const bodyStart = 5;
  const bodyEnd = bodyStart + sheetDef.data.length - 2;
  const bodyRange = sheet.getRange(`A${bodyStart}:${lastCol}${bodyEnd}`);

  titleRange.merge();
  titleRange.values = [[sheetDef.title]];
  titleRange.format = {
    fill: palette.navy,
    font: { name: "Aptos Display", size: 16, bold: true, color: palette.white },
    horizontalAlignment: "left",
    verticalAlignment: "center",
  };
  titleRange.format.rowHeight = 28;

  noteRange.merge();
  noteRange.values = [[sheetDef.note]];
  noteRange.format = {
    fill: palette.sand,
    font: { name: "Aptos", size: 10, color: palette.text, bold: false },
    wrapText: true,
    verticalAlignment: "center",
    borders: { preset: "all", style: "thin", color: palette.border },
  };
  noteRange.format.rowHeight = 40;

  sheet.getRange(`A4:${lastCol}${bodyEnd}`).values = sheetDef.data;
  headerRange.format = {
    fill: palette.navy,
    font: { name: "Aptos", size: 10, bold: true, color: palette.white },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    borders: { preset: "all", style: "thin", color: palette.border },
  };
  headerRange.format.rowHeight = 24;

  bodyRange.format = {
    font: { name: "Aptos", size: 10, bold: false, color: palette.text },
    wrapText: true,
    verticalAlignment: "top",
    borders: { preset: "all", style: "thin", color: palette.border },
  };

  for (let r = bodyStart; r <= bodyEnd; r++) {
    const rowRange = sheet.getRange(`A${r}:${lastCol}${r}`);
    rowRange.format = {
      ...rowRange.format,
      fill: r % 2 === 0 ? palette.blue2 : palette.white,
      font: { name: "Aptos", size: 10, bold: false, color: palette.text },
      wrapText: true,
      verticalAlignment: "top",
      borders: { preset: "all", style: "thin", color: palette.border },
    };
    rowRange.format.rowHeight = 32;
  }

  for (let i = 0; i < sheetDef.widths.length; i++) {
    sheet.getRangeByIndexes(0, i, bodyEnd, 1).format.columnWidth = sheetDef.widths[i];
  }

  sheet.freezePanes.freezeRows(4);
  sheet.showGridLines = false;

  if (sheetDef.name === "03_指标分类") {
    const categoryCol = sheet.getRange(`A${bodyStart}:A${bodyEnd}`);
    categoryCol.format = {
      ...categoryCol.format,
      font: { name: "Aptos", size: 10, bold: true, color: palette.navy },
    };
  }
}

const workbook = Workbook.create();
for (const def of sheets) {
  const sheet = workbook.worksheets.add(def.name);
  styleWorkbookSheet(def, sheet);
}

await fs.mkdir(outputDir, { recursive: true });
const out = await SpreadsheetFile.exportXlsx(workbook);
await out.save(outputPath);
console.log(outputPath);
