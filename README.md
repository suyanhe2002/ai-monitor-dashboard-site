# AI Class Monitor CLI

AI监课面板项目的可发布 CLI 版本，包含：

- 周度 AI 监课面板构建
- 单文件 / 站点版面板导出
- 跟进套表生成
- 当前前端面板静态资源

## 项目结构

- `source/`
  当前 AI 监课面板核心脚本与前端资源
- `config/example.local.json`
  本地数据路径配置示例
- `cli.py`
  统一命令入口

## 安装依赖

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

如果需要生成 `build_followup_workbook.mjs` 对应的 Excel 套表，还需要 Node.js 环境。

## 常用命令

1. 构建主面板 HTML

```bash
python3 cli.py build-dashboard --config config/example.local.json
```

2. 导出可分享站点目录

```bash
python3 cli.py export-share --config config/example.local.json --share-dir dist/site
```

3. 打包站点 zip

```bash
python3 cli.py zip-site --site-dir dist/site --output dist/ai_monitor_dashboard_site.zip
```

## 配置说明

`config/example.local.json` 里主要维护：

- 标准表路径
- 各周 AI 监课指标表路径
- 北极星表路径
- 当前周 / 对比周口径
- 输出 HTML 路径

CLI 会先把这些配置映射成环境变量，再调用现有构建脚本，避免把本地绝对路径写死在发布命令里。

## 当前默认口径

- 当前周：`7W1`
- 对比周：`W5`
- 当前月标签：`202607`
- 对比月标签：`202606`

## 部署说明

当前线上路径为：

- `http://teaching-quality.vipthink.cn/ai-class-monitor/`

如果要更新线上内容，可先重新导出站点目录，再把 `dist/site/` 覆盖到服务器站点目录。
