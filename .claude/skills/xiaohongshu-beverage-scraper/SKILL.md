---
name: xiaohongshu-beverage-scraper
description: >
  小红书多赛道数据抓取全流程 SOP —— 从安装 OpenCLI 到生成增强版 Excel +
  数据透视表 + 使用说明文档，一键标准化执行。

  适用场景：任意品类的小红书舆情调研、消费者洞察数据收集、品牌口碑监测。
  支持自定义关键词分组（正向+负向均衡）、评论楼中楼抓取、
  自动去噪、Excel 增强输出。关键词通过 config.py 配置，完全脱离代码。
triggers:
  - xiaohongshu: 小红书 调研/爬虫/数据抓取/舆情
  - keyword: 小红书爬虫/小红书数据/小红书调研
  - action: 抓取小红书数据/跑小红书爬虫/更新小红书数据
metadata:
  author: 诡秘专供
  version: "2.0"
  last_updated: "2026-07-14"
  requires:
    - agent-reach (OpenCLI + xiaohongshu channel)
    - Chrome 浏览器（已登录小红书）
    - Python 3.10+ (pandas, openpyxl)
---

# 🧋 小红书数据抓取 SOP

> 完整标准化工作流：从零安装 → 配置关键词 → 数据抓取 → 增强 Excel → 数据透视

## 工作流总览

```
安装环境 → 配置关键词 → 验证登录 → 运行爬虫 → 增强 Excel → 追加透视 → ✅ 完成
```

| 阶段 | 做什么 | 产出 |
|---|---|---|
| ① 环境准备 | 安装 Agent Reach + OpenCLI + 小红书频道 | 可用的抓取环境 |
| ② 配置关键词 | `cp config.example.py config.py` → 编辑 | 自定义关键词配置 |
| ③ 登录验证 | Chrome 登录小红书 + opencli doctor 检查 | 登录态就绪 |
| ④ 数据抓取 | 运行 `tools/beverage_scanner_v2.py` | 原始 Excel（完整版） |
| ⑤ 数据增强 | 运行 `tools/enhance_excel.py` | 增强版 Excel（8 Sheet） |
| ⑥ 追加透视 | 运行 `tools/add_pivot_sheets.py` | 含透视表的最终 Excel |

---

## ① 环境准备（首次才需要）

### 1.1 安装 Agent Reach

```bash
# 官方安装脚本（如果还没装）
curl -s https://raw.githubusercontent.com/Panniantong/agent-reach/main/docs/install.md
```

Windows 用户使用 venv 方式：
```bash
python -m venv ~/.agent-reach-venv
~/.agent-reach-venv/Scripts/pip install agent-reach
```

### 1.2 安装 OpenCLI + 小红书频道

```bash
npm install -g @jackwener/opencli
agent-reach install opencli
agent-reach install xiaohongshu
```

### 1.3 安装 Chrome 扩展

1. Chrome 应用商店搜索 **OpenCLI Companion** 并安装
2. 点击扩展图标 → 连接本地 daemon
3. 运行 `opencli doctor` 确认全绿

### 1.4 Python 依赖

```bash
pip install pandas openpyxl
```

---

## ② 配置关键词（首次需要）

### 创建配置文件

```bash
cp config.example.py config.py
```

### 编辑你的关键词

`config.py` 结构如下（示例使用宠物赛道，你可换成任何品类）：

```python
KEYWORD_GROUPS = {
    "🐱 猫咪养护": ["猫咪掉毛怎么办", "猫粮推荐", "猫咪呕吐原因"],
    "🐶 狗狗日常": ["狗狗零食推荐", "小型犬好养吗", "狗狗训练技巧"],
    "🏠 宠物家居": ["猫爬架推荐", "宠物友好租房", "阳台养猫改造"],
    "💊 宠物健康": ["猫咪绝育注意事项", "狗狗皮肤病", "宠物保险值得买吗"],
}

KW_TO_GROUP = { ... }   # 跟上面一一对应
GROUP_ORDER = ["🐱 猫咪养护", "🐶 狗狗日常", ...]
```

- 组数不限、每组词数不限
- 所有脚本自动从 `config.py` 读取，无需改代码
- `config.py` 已加入 `.gitignore`，不会上传 GitHub

---

## ③ 登录验证（每次抓取前必做）

### 2.1 登录小红书

在已连接 OpenCLI 扩展的 Chrome 中：
1. 打开 `xiaohongshu.com` → 扫码登录
2. **建议用小号**（降低主号被封风险）
3. **新号第一次跑成功率最高**

### 2.2 验证状态

```bash
# 两条命令都必须通过
opencli doctor          # 所有项 [OK]
opencli xiaohongshu whoami  # logged_in: true
```

输出示例：
```
opencli v1.8.6 doctor (node v24.16.0)
[OK] Daemon: running on port 19825
[OK] Extension: connected (v1.0.22)
[OK] Connectivity: connected

logged_in: true
site: xiaohongshu
username: Art3mis
```

### ⚠️ 风控须知

- **只跑一次！** 同一个号 24 小时内不要重复跑
- 搜索阶段安全（100% 成功率），评论抓取是风控重点
- 如果大面积 SECURITY_BLOCK → 换新号 + 等 24 小时
- 看到封控迹象立即 Ctrl+C 停止，不要硬跑

---

## ③ 数据抓取

### 一键运行（推荐）

```bash
# 或双击
tools/一键抓取.bat
```

### 手动运行

```bash
cd "项目目录"
python tools/beverage_scanner_v2.py
```

### 抓取配置（v2 默认参数）

| 参数 | 值 | 说明 |
|---|---|---|
| 关键词 | 通过 `config.py` 自定义 | 组数、词数完全自由 |
| 每关键词笔记 | 4 条 | `MAX_NOTES_PER_KEYWORD` |
| 评论策略 | 全局热度 TOP 10 笔记 | 按点赞数排序取前 10 |
| 每笔记评论 | 30 条（含子回复） | --with-replies |
| 关键词间隔 | 3-6 秒 | 搜索阶段，快扫快过 |
| 笔记间隔 | 8-16 秒 | 评论阶段，拉长间隔防封 |
| Fallback | 3 层 | 被拦自动切换次热笔记 |
| 灌水过滤 | 自动 | 过滤"蹲""码""打卡"等无意义评论 |

### 关键词分组

关键词全部通过 `config.py` 配置，详见第②步。示例（宠物赛道）：

| 组别 | 关键词 |
|---|---|
| 🐱 猫咪养护 | 猫咪掉毛怎么办、猫粮推荐、猫咪呕吐原因 |
| 🐶 狗狗日常 | 狗狗零食推荐、小型犬好养吗、狗狗训练技巧 |
| 🏠 宠物家居 | 猫爬架推荐、宠物友好租房、阳台养猫改造 |
| 💊 宠物健康 | 猫咪绝育注意事项、狗狗皮肤病、宠物保险值得买吗 |

### 预期产出

- 笔记：取决于关键词数 × 4
- 评论：~300-500 条（主评论 + 子回复）
- 输出目录：`output/[关键词前缀]_小红书_完整版_YYYYMMDD_HHMMSS.xlsx`

---

## ④ 数据增强

```bash
python tools/enhance_excel.py
```

**做了什么：**
- 笔记列表新增「组别」列，按组排序
- 评论明细新增「组别」列
- 新增「数据概览」Sheet（4 组汇总对比）
- 新增「分组评论」Sheet（4 组分区展示）
- 主评论行浅紫色填充 🟣
- 所有 Sheet 加自动筛选 + 列宽适配 + 冻结首行

**输出：** `output/饮料赛道_小红书_增强版_YYYYMMDD.xlsx`

---

## ⑤ 追加数据透视

```bash
python tools/add_pivot_sheets.py
```

**追加 4 个透视 Sheet：**

| Sheet | 内容 |
|---|---|
| 透视-组别×评论类型 | 四组横向对比矩阵（笔记数/评论/均赞/效率） |
| 透视-关键词明细 | 12 个搜索词逐词拆解 |
| 透视-TOP笔记排行 | 热度前 15 名榜单 |
| 透视-评论焦点 | TOP 20 热评 + 每组 TOP 3 |

---

## ⑥ 生成使用说明

### Markdown 详细版

脚本会自动生成 `output/使用说明_详细版.md`，包含：
- 文件信息 & 数据概览
- 8 个 Sheet 逐张拆解（每列含义 + 操作技巧 + 分析思路）
- Excel 操作速成指南（筛选/排序/搜索/条件格式）
- 5 分钟快速分析路线
- 5 个常见分析场景 Recipe（舆情/选题/洞察/预警/报告）
- 颜色标记速查表
- 数据质量 & 局限性说明
- 数据更新指南

### 转 Word / PDF

```bash
cd "D:\claude\AI race forest\output"
pandoc "使用说明_详细版.md" -o "饮料赛道小红书数据_使用说明书.docx"
```

Word 打开后 → 文件 → 另存为 → PDF（推荐，字体/表格/颜色全保留）

---

## 📁 最终产出物清单

```
output/
├── 饮料赛道_小红书_完整版_YYYYMMDD_HHMMSS.xlsx    ← 原始抓取（留底）
├── 饮料赛道_小红书_增强版_YYYYMMDD.xlsx            ← ⭐ 最终数据（8 Sheet）
├── 使用说明_详细版.md                              ← 说明书源文件
└── 饮料赛道小红书数据_使用说明书.docx               ← Word 版说明书
```

---

## 🔧 自定义抓取参数

如需调整抓取规模，编辑 `beverage_scanner_v2.py` 顶部配置区：

```python
# 每关键词收录笔记数
MAX_NOTES_PER_KEYWORD = 4       # 改这个调笔记量

# 每笔记抓取评论数
MAX_COMMENTS_PER_NOTE = 30      # 改这个调评论量

# 评论笔记间休息（秒）
COMMENT_NOTE_DELAY_MIN = 8      # 太快容易封
COMMENT_NOTE_DELAY_MAX = 16

# 全局热度 TOP N 条笔记
TOP_COMMENT_NOTES = 10          # 改这个调抓评论的笔记数
```

### 换关键词

**编辑 `config.py`** 即可，无需改任何代码：

```python
KEYWORD_GROUPS = {
    "你的组名": ["你的关键词1", "你的关键词2", "你的关键词3"],
    # ... 组数不限、词数不限
}
KW_TO_GROUP = { ... }   # 同步更新
GROUP_ORDER = [...]     # 控制 Excel 中的排序
```

所有脚本（爬虫、增强、透视）都从 `config.py` 读取，一步到位。

---

## 🚨 常见问题

| 问题 | 原因 | 解决 |
|---|---|---|
| `opencli: command not found` | npm 全局路径不在 PATH | 用完整路径 `%APPDATA%/Roaming/npm/opencli.cmd` |
| `SECURITY_BLOCK` | 小红书风控拦截 | 换新号 + 等 24h |
| GBK 编码报错 | Windows 终端 emoji 乱码 | 脚本已内置 UTF-8 修复 |
| Excel 文件打不开 | 文件被 Excel 占用 | 关闭 Excel 后重跑 |
| 搜索有结果但评论全空 | 账号被标记 | 换新号 |
| `opencli doctor` 扩展未连接 | Chrome 扩展未启动 | 打开 Chrome → 点扩展图标 → 重连 |

---

## 📌 核心原则

1. **换号第一次跑效果最好** → 每次抓取前切新号
2. **一次只跑一轮，不要反复跑** → 连续跑会累积风控分
3. **被封立刻停，不要硬扛** → 硬跑可能封号
4. **搜索安全，评论危险** → 搜索阶段的延迟可以短，评论阶段必须拉长
5. **小号是耗材** → 不要用主号/大号跑爬虫

---

## 🔗 相关文件

| 文件 | 说明 |
|---|---|
| `config.py` | 🔒 你的关键词配置（本地，不提交） |
| `config.example.py` | 📋 配置模板（公开，可提交） |
| `tools/beverage_scanner_v2.py` | 🐍 主爬虫脚本（OpenCLI 版） |
| `tools/beverage_scanner.py` | 🕸️ 旧版 Playwright 方案（备用） |
| `tools/enhance_excel.py` | 📊 Excel 增强脚本 |
| `tools/add_pivot_sheets.py` | 🔀 数据透视表追加脚本 |
| `tools/一键抓取.bat` | 🖱️ Windows 一键运行脚本 |
