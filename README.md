# 🧋 xiaohongshu-beverage-scraper

小红书饮料赛道数据抓取工具 —— 从搜索到增强 Excel，一条龙输出。

[![Python](https://img.shields.io/badge/Python-3.10+-blue)](https://python.org)
[![License](https://img.shields.io/badge/license-MIT-green)](./LICENSE)

## 这是什么

一套标准化的小红书数据采集 + 分析工具链，专门针对**泛饮品赛道**（奶茶/果茶/咖啡/低卡饮料）。12 个关键词分 4 组，正负向均衡覆盖，自动抓取笔记 + 评论含楼中楼，输出增强版 Excel 含数据透视表。

**核心卖点：用 OpenCLI 复用 Chrome 登录态，无需额外登录，无需 Playwright 注入。**

## 快速开始

### 前提

| 依赖 | 怎么装 |
|---|---|
| Python 3.10+ | [python.org](https://python.org) |
| Node.js | [nodejs.org](https://nodejs.org) |
| Agent Reach + OpenCLI | `npm i -g @jackwener/opencli` + [安装指南](https://github.com/Panniantong/agent-reach) |
| Chrome 扩展 | 应用商店搜 `OpenCLI Companion` 安装 |
| 小红书账号 | Chrome 登录（建议小号） |

### 安装

```bash
git clone https://github.com/02Ruichen/xiaohongshu-beverage-scraper.git
cd xiaohongshu-beverage-scraper
pip install -r requirements.txt
```

### 一键运行（Windows）

双击 `一键抓取.bat`

### 手动运行

```bash
# 1. 验证环境
opencli doctor                 # 全绿才行
opencli xiaohongshu whoami     # 确认已登录

# 2. 抓取数据
python beverage_scanner_v2.py

# 3. 生成增强版 Excel
python enhance_excel.py

# 4. 追加数据透视表
python add_pivot_sheets.py
```

## 工作流

```
环境检查 → 搜索48条笔记 → 抓TOP10笔记评论 → 增强Excel → 数据透视 → ✅
```

| 步骤 | 脚本 | 产出 |
|---|---|---|
| ① | `beverage_scanner_v2.py` | `完整版.xlsx`（笔记+评论底表） |
| ② | `enhance_excel.py` | `增强版.xlsx`（组别归类，8 Sheet） |
| ③ | `add_pivot_sheets.py` | 含 4 个透视 Sheet 的最终版 |

## 抓取配置

### 关键词分组（4组 × 3词）

| 组别 | 关键词 | 倾向 |
|---|---|---|
| 🍹 口味上新 | 奶茶新品、果茶推荐、咖啡新品 | 正向 |
| 💚 健康诉求 | 低卡奶茶、无糖饮料推荐、减脂期喝什么 | 正向 |
| 📸 视觉圈层 | 奶茶探店、夏日饮品合集、网红奶茶打卡 | 正向 |
| ⚠️ 踩雷避坑 | 奶茶避雷、饮料踩雷、最难喝饮料 | 负向 |

### 规模参数

| 参数 | 默认值 | 在哪改 |
|---|---|---|
| 每关键词笔记 | 4 条 | `beverage_scanner_v2.py` → `MAX_NOTES_PER_KEYWORD` |
| 每笔记评论 | 30 条 | `beverage_scanner_v2.py` → `MAX_COMMENTS_PER_NOTE` |
| 评论笔记 TOP N | 10 条 | `beverage_scanner_v2.py` → `TOP_COMMENT_NOTES` |
| 笔记抓取间隔 | 8-16 秒 | `beverage_scanner_v2.py` → `COMMENT_NOTE_DELAY_MIN/MAX` |

### 换赛道

编辑 `beverage_scanner_v2.py` 的 `KEYWORD_GROUPS` 字典，替换成你自己的关键词。同时同步更新 `enhance_excel.py` 和 `add_pivot_sheets.py` 中的 `KW_TO_GROUP` 映射。

## 输出 Excel 结构

最终增强版共 **8 个 Sheet**：

| Sheet | 内容 |
|---|---|
| 📋 笔记列表 | 47 条笔记 + 组别列，按组排序 |
| 💬 评论明细 | 420 条评论 + 组别，主评论浅紫标注 🟣 |
| 📊 数据概览 | 四组横向汇总对比 |
| 🗂️ 分组评论 | 按组分区展示，蓝底标题分隔 |
| 🔀 透视-组别×评论类型 | 四组矩阵对比 |
| 🔀 透视-关键词明细 | 12 词逐词拆解 |
| 🔀 透视-TOP笔记排行 | 热度前 15 名 |
| 🔀 透视-评论焦点 | TOP 20 热评 + 每组 TOP 3 |

## 作为 Claude Code Skill 使用

### 当前目录激活

在项目目录下打开 Claude Code，Skill 自动生效。

调用方式：
```
/xiaohongshu-beverage-scraper
```
或自然语言：「帮我抓小红书饮料数据」

### 全局安装（队友用）

```bash
cp -r .claude/skills/xiaohongshu-beverage-scraper ~/.claude/skills/
```

这样在任何目录都能调用这个 Skill。

## 文件结构

```
.
├── beverage_scanner_v2.py   # 主爬虫（OpenCLI 版）
├── enhance_excel.py         # Excel 增强脚本
├── add_pivot_sheets.py      # 数据透视表追加
├── 一键抓取.bat             # Windows 一键运行
├── requirements.txt         # Python 依赖
├── .gitignore
└── .claude/skills/
    └── xiaohongshu-beverage-scraper/
        └── SKILL.md         # SOP 标准化 Skill
```

## 防封策略

- 搜索阶段：每关键词 3-6 秒间隔，快速扫过
- 评论阶段：每笔记 8-16 秒间隔，模拟人工浏览
- 被拦自动 fallback 到次热笔记（最多 3 层）
- **建议新号跑，跑完一次等 24 小时**
- 自动过滤「蹲」「码」「打卡」等灌水评论

## 常见问题

| 问题 | 解决 |
|---|---|
| `opencli doctor` 不全绿 | Chrome 扩展未连接 → 点扩展图标重连 |
| `SECURITY_BLOCK` | 风控拦截 → 换新号 + 等 24h |
| 搜得到但评论全是空 | 账号被标记 → 换新号 |
| GBK 报错 | 脚本已内置 UTF-8 修复，不需额外操作 |
| 输出数量不够 | 小红书上该关键词内容少 → 换词 |

## License

MIT
