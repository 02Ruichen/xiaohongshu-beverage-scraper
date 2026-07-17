# 🧋 xiaohongshu-beverage-scraper

小红书多赛道数据抓取工具 —— 从搜索到增强 Excel，一条龙输出。

[![Python](https://img.shields.io/badge/Python-3.10+-blue)](https://python.org)
[![License](https://img.shields.io/badge/license-MIT-green)](./LICENSE)

## 这是什么

一套标准化的小红书数据采集 + 分析工具链。支持**自定义关键词分组**（正向+负向均衡），自动抓取笔记 + 评论含楼中楼，输出增强版 Excel 含数据透视表。

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

# 配置关键词（重要！）
cp config.example.py config.py
# 然后编辑 config.py，填入你自己的关键词
```

### 手动运行

```bash
# 0. 先配置关键词
cp config.example.py config.py    # 仅首次需要

# 1. 验证环境
opencli doctor                     # 全绿才行
opencli xiaohongshu whoami        # 确认已登录

# 2. 抓取数据
python tools/beverage_scanner_v2.py

# 3. 生成增强版 Excel
python tools/enhance_excel.py

# 4. 追加数据透视表
python tools/add_pivot_sheets.py
```

### 一键运行（Windows）

```bash
# 或直接双击
tools/一键抓取.bat
```

## 自定义关键词

**关键词不再硬编码在脚本里**，而是通过 `config.py` 统一管理。

### 1. 创建配置文件

```bash
cp config.example.py config.py
```

### 2. 编辑你的关键词

```python
# config.py（已加入 .gitignore，不会被上传）

KEYWORD_GROUPS = {
    "🐱 猫咪养护": ["猫咪掉毛怎么办", "猫粮推荐", "猫咪呕吐原因"],
    "🐶 狗狗日常": ["狗狗零食推荐", "小型犬好养吗", "狗狗训练技巧"],
    "🏠 宠物家居": ["猫爬架推荐", "宠物友好租房", "阳台养猫改造"],
    "💊 宠物健康": ["猫咪绝育注意事项", "狗狗皮肤病", "宠物保险值得买吗"],
}

KW_TO_GROUP = {
    "猫咪掉毛怎么办": "🐱 猫咪养护",
    "猫粮推荐": "🐱 猫咪养护",
    # ... 跟上面一一对应
}

GROUP_ORDER = ["🐱 猫咪养护", "🐶 狗狗日常", "🏠 宠物家居", "💊 宠物健康"]
```

> 💡 示例用了**宠物赛道**，你可以换成任何品类：美妆、数码、食品、穿搭……

### 3. 组数不限、词数不限

你需要几组就写几组，每组几个词完全自由。所有分析透视表会自动适配。

## 工作流

```
配置文件 → 搜索笔记 → 抓取评论 → 增强Excel → 数据透视 → ✅
```

| 步骤 | 脚本 | 产出 |
|---|---|---|
| ① | `tools/beverage_scanner_v2.py` | `完整版.xlsx`（笔记+评论底表） |
| ② | `tools/enhance_excel.py` | `增强版.xlsx`（组别归类，8 Sheet） |
| ③ | `tools/add_pivot_sheets.py` | 含 4 个透视 Sheet 的最终版 |

## 规模参数

| 参数 | 默认值 | 在哪改 |
|---|---|---|
| 每关键词笔记 | 4 条 | `tools/beverage_scanner_v2.py` → `MAX_NOTES_PER_KEYWORD` |
| 每笔记评论 | 30 条 | `tools/beverage_scanner_v2.py` → `MAX_COMMENTS_PER_NOTE` |
| 评论笔记 TOP N | 10 条 | `tools/beverage_scanner_v2.py` → `TOP_COMMENT_NOTES` |
| 笔记抓取间隔 | 8-16 秒 | `tools/beverage_scanner_v2.py` → `COMMENT_NOTE_DELAY_MIN/MAX` |

## 输出 Excel 结构

最终增强版共 **8 个 Sheet**：

| Sheet | 内容 |
|---|---|
| 📋 笔记列表 | 笔记 + 组别列，按组排序 |
| 💬 评论明细 | 评论 + 组别，主评论浅紫标注 🟣 |
| 📊 数据概览 | 各组横向汇总对比 |
| 🗂️ 分组评论 | 按组分区展示，蓝底标题分隔 |
| 🔀 透视-组别×评论类型 | 各组矩阵对比 |
| 🔀 透视-关键词明细 | 逐词拆解 |
| 🔀 透视-TOP笔记排行 | 热度前 15 名 |
| 🔀 透视-评论焦点 | TOP 20 热评 + 每组 TOP 3 |

## 作为 Claude Code Skill 使用

### 当前目录激活

在项目目录下打开 Claude Code，Skill 自动生效。

调用方式：
```
/xiaohongshu-beverage-scraper
```
或自然语言：「帮我抓小红书数据」

### 全局安装（队友用）

```bash
cp -r .claude/skills/xiaohongshu-beverage-scraper ~/.claude/skills/
```

## 文件结构

```
.
├── config.example.py            # 📋 配置模板（公开，可提交）
├── config.py                    # 🔒 你的真实配置（本地，不提交）
├── README.md
├── requirements.txt
├── .gitignore
├── tools/
│   ├── beverage_scanner_v2.py   # 主爬虫（OpenCLI 版）
│   ├── enhance_excel.py         # Excel 增强脚本
│   ├── add_pivot_sheets.py      # 数据透视表追加
│   └── 一键抓取.bat             # Windows 一键运行
├── output/                      # 🔒 输出目录（不提交）
└── .claude/skills/
    └── xiaohongshu-beverage-scraper/
        └── SKILL.md             # SOP 标准化 Skill
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
| 找不到 config.py | 先执行 `cp config.example.py config.py` |

## License

MIT
