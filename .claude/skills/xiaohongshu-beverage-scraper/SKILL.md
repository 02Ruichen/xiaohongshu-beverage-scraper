---
name: xiaohongshu-beverage-scraper
description: >
  小红书饮料赛道数据抓取全流程 SOP —— 从安装 OpenCLI 到生成增强版 Excel +
  数据透视表 + 使用说明文档，一键标准化执行。

  适用场景：泛饮品/饮料品类的小红书舆情调研、消费者洞察数据收集、
  品牌口碑监测。支持多关键词分组（正向+负向均衡）、评论楼中楼抓取、
  自动去噪、Excel 增强输出。
triggers:
  - xiaohongshu: 小红书饮料/饮品/奶茶/咖啡 调研/爬虫/数据抓取/舆情
  - keyword: 饮料赛道/饮品调研/奶茶数据/咖啡舆情/小红书爬虫
  - action: 抓取小红书饮料数据/跑饮料爬虫/更新饮料数据
metadata:
  author: 诡秘专供
  version: "2.0"
  last_updated: "2026-07-14"
  requires:
    - agent-reach (OpenCLI + xiaohongshu channel)
    - Chrome 浏览器（已登录小红书）
    - Python 3.10+ (pandas, openpyxl)
---

# 🧋 小红书饮料赛道数据抓取 SOP

> 完整标准化工作流：从零安装 → 数据抓取 → 增强 Excel → 数据透视 → 使用说明

## 工作流总览

```
安装环境 → 验证登录 → 运行爬虫 → 增强 Excel → 追加透视 → 生成说明 → ✅ 完成
```

| 阶段 | 做什么 | 产出 |
|---|---|---|
| ① 环境准备 | 安装 Agent Reach + OpenCLI + 小红书频道 | 可用的抓取环境 |
| ② 登录验证 | Chrome 登录小红书 + opencli doctor 检查 | 登录态就绪 |
| ③ 数据抓取 | 运行 beverage_scanner_v2.py | 原始 Excel（完整版） |
| ④ 数据增强 | 运行 enhance_excel.py | 增强版 Excel（8 Sheet） |
| ⑤ 追加透视 | 运行 add_pivot_sheets.py | 含透视表的最终 Excel |
| ⑥ 生成文档 | 编写使用说明书 | Markdown + Word 文档 |

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

## ② 登录验证（每次抓取前必做）

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

双击 `一键抓取.bat`，脚本会自动：
1. 检查 OpenCLI 可用性
2. 验证小红书登录态
3. 等你按 Enter 确认
4. 执行完整抓取流程

### 手动运行

```bash
cd "D:\claude\AI race forest"
python beverage_scanner_v2.py
```

### 抓取配置（v2 默认参数）

| 参数 | 值 | 说明 |
|---|---|---|
| 关键词 | 12 个，4 组 | 🍹口味上新 / 💚健康诉求 / 📸视觉圈层 / ⚠️踩雷避坑 |
| 每关键词笔记 | 4 条 | 12 × 4 = 最多 48 条 |
| 评论策略 | 全局热度 TOP 10 笔记 | 按点赞数排序取前 10 |
| 每笔记评论 | 30 条（含子回复） | --with-replies |
| 关键词间隔 | 3-6 秒 | 搜索阶段，快扫快过 |
| 笔记间隔 | 8-16 秒 | 评论阶段，拉长间隔防封 |
| Fallback | 3 层 | 被拦自动切换次热笔记 |
| 灌水过滤 | 自动 | 过滤"蹲""码""打卡"等无意义评论 |

### 关键词分组

```
🍹 口味上新: 奶茶新品、果茶推荐、咖啡新品
💚 健康诉求: 低卡奶茶、无糖饮料推荐、减脂期喝什么
📸 视觉圈层: 奶茶探店、夏日饮品合集、网红奶茶打卡
⚠️ 踩雷避坑: 奶茶避雷、饮料踩雷、最难喝饮料
```

### 预期产出

- 笔记：~47-48 条
- 评论：~300-500 条（主评论 + 子回复）
- 输出目录：`output/饮料赛道_小红书_完整版_YYYYMMDD_HHMMSS.xlsx`

---

## ④ 数据增强

```bash
cd "D:\claude\AI race forest"
python enhance_excel.py
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
cd "D:\claude\AI race forest"
python add_pivot_sheets.py
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

修改 `KEYWORD_GROUPS` 字典即可：
```python
KEYWORD_GROUPS = {
    "🍹 口味上新": ["你的关键词1", "你的关键词2", "你的关键词3"],
    # ... 可以增减组别和关键词
}
```

同步更新 `enhance_excel.py` 和 `add_pivot_sheets.py` 中的 `KW_TO_GROUP` 映射。

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
| `beverage_scanner_v2.py` | 🐍 主爬虫脚本（OpenCLI 版） |
| `beverage_scanner.py` | 🕸️ 旧版 Playwright 方案（备用） |
| `enhance_excel.py` | 📊 Excel 增强脚本 |
| `add_pivot_sheets.py` | 🔀 数据透视表追加脚本 |
| `一键抓取.bat` | 🖱️ Windows 一键运行脚本 |
| `output/使用说明_详细版.md` | 📝 数据使用说明书 |
