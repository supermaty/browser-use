# Yuntu-Data（巨量云图数据助手）项目文档

## 1. 项目简介

**Yuntu-Data** 是一个基于 Browser-Use 的自动化营销数据提取工具。它通过 Gradio Web 界面接收用户的自然语言指令，自动驱动浏览器登录[巨量云图](https://yuntu.oceanengine.com)平台，按预定的取数路径提取营销复盘数据，并生成标准化的 Excel 报表。

**核心价值**: 将原本需要人工花费数小时在云图平台上反复点击、复制、整理的工作，缩短为一次自然语言对话。

**技术栈**:
- **Browser-Use**: AI 浏览器自动化引擎
- **Google Gemini**: LLM 驱动的意图理解和浏览器操控
- **Gradio**: Web UI 框架
- **Pydantic**: 结构化数据模型
- **openpyxl / pandas**: Excel 读写

---

## 2. 系统架构

整个系统是一个 **9 步流水线**，从用户输入到 Excel 输出：

```
┌─────────────────────────────────────────────────────────────┐
│                 Gradio Web UI (app.py)                       │
│                 用户输入自然语言任务描述                       │
└──────────────────────────┬──────────────────────────────────┘
                           │
                           ▼
┌──────────────────────────────────────────────────────────────┐
│  Step 1: 意图提取 (intent.py)                                │
│  Gemini LLM 解析用户输入 → YuntuTask                         │
└──────────────────────────┬───────────────────────────────────┘
                           │
                           ▼
┌──────────────────────────────────────────────────────────────┐
│  Step 2: 登录检查 (login.py)                                 │
│  检查浏览器 profile → 复用 cookies / 引导手动登录             │
└──────────────────────────┬───────────────────────────────────┘
                           │
                           ▼
┌──────────────────────────────────────────────────────────────┐
│  Step 3: 提示词组装 (config.py)                              │
│  品牌模板选择 + 变量填充 → 最终 Agent 提示词                  │
└──────────────────────────┬───────────────────────────────────┘
                           │
                           ▼
┌──────────────────────────────────────────────────────────────┐
│  Step 4: Agent 执行 (app.py + Browser-Use)                   │
│  自动导航云图 → 切换品牌 → 提取 6 大数据模块                  │
│  同时下载 KOL/投流 Excel 文件                                │
└──────────────────────────┬───────────────────────────────────┘
                           │
                           ▼
┌──────────────────────────────────────────────────────────────┐
│  Step 5: 数据提取 (data_extractor.py)                        │
│  正则解析 Agent 输出 → YuntuDataReport (Pydantic)            │
└──────────────────────────┬───────────────────────────────────┘
                           │
                           ▼
┌──────────────────────────────────────────────────────────────┐
│  Step 6: Excel 文件处理 (excel_processor.py)                 │
│  读取下载的 Excel → 提取 KOL/投流细分数据                    │
└──────────────────────────┬───────────────────────────────────┘
                           │
                           ▼
┌──────────────────────────────────────────────────────────────┐
│  Step 7: 指标计算 (excel_processor.py)                       │
│  CPM / CPE / CPS / CPA3 / ROI 等成本指标                    │
└──────────────────────────┬───────────────────────────────────┘
                           │
                           ▼
┌──────────────────────────────────────────────────────────────┐
│  Step 8: Excel 导出 (excel_exporter.py)                      │
│  模板模式（优先） / 标准模式 → 生成最终 Excel 报表            │
└──────────────────────────┬───────────────────────────────────┘
                           │
                           ▼
┌──────────────────────────────────────────────────────────────┐
│  Step 9: UI 反馈 (app.py)                                    │
│  展示执行日志 + 结果摘要 + Excel 下载路径                     │
└──────────────────────────────────────────────────────────────┘
```

---

## 3. 核心模块详解

### 3.1 app.py — Gradio UI + 主编排

**职责**: Web 界面 + 9 步流水线的总调度

- Gradio 聊天界面，支持消息历史和实时反馈
- `process_task_generator()`: 异步生成器，依次执行 9 个步骤
- 实时 Agent 日志：通过 `on_step_start` / `on_step_end` 回调展示 Agent 的思考和操作
- 自定义 Tool: `sum_search_count_in_date_range()` — 从下载的 Excel 中按日期范围汇总回搜次数
- 全局状态管理：`_session_intent`（跨轮次保持意图）、`_current_agent`（支持停止按钮）
- 任务隔离：每次运行创建独立目录 `tasks/task_YYYYMMDD_HHMMSS/`

### 3.2 intent.py — 自然语言意图提取

**职责**: 用 Gemini LLM 从自然语言中提取结构化任务参数

**输入**: 用户的自然语言描述（如「获取黑玩洗护星图报告数据，消耗 50 万」）

**输出**: `YuntuTask` Pydantic 模型

**关键模型**:
- `YuntuTask`: 任务根模型
  - `report_name`: 单报告名（通用品牌模式）
  - `brand_categories`: 品牌分类列表（黑玩多分类模式），每个含星图/竞价报告名和消耗金额
  - `kol_content_date_type` / `kol_content_date_range`: 达人内容复盘的时间范围
  - `insight_time_ranges`: 行业搜索洞察的一到多个时间范围
  - `brand_name`: 品牌名
  - `consumption_amount_star` / `consumption_amount_bid`: 星图/竞价消耗金额
- `BrandCategory`: 品牌分类（含 category_name、report_name_star/bid、consumption_amount_star/bid）
- `InsightTimeRange`: 搜索洞察时间范围（date_range_type + insight_date_range）

**多轮上下文**: 支持跨轮对话——新一轮提取的值会与之前的 `previous_intent` 合并，未提及的字段保持不变。

### 3.3 config.py — 配置管理与提示词模板

**职责**: 环境变量、路径常量、Agent 提示词模板

**提示词模板**（两套）:

| 模板 | 适用场景 | 内容 |
|------|---------|------|
| `HEYONE_PROMPT_TEMPLATE` | 黑玩品牌（多分类） | ~270 行，6 步 SOP 详细取数路径 |
| `GENERIC_PROMPT_TEMPLATE` | 通用品牌（单报告） | 简化版 4 步取数路径 |

模板中使用占位变量（如 `{report_name}`、`{brand_name}`、`{date_ranges}`），在 Step 3 中动态填充。

**品牌名映射**: 中文品牌名 ↔ 英文品牌名的映射表，适配云图平台的品牌切换。

### 3.4 login.py — 浏览器登录与 Profile 管理

**职责**: 确保浏览器已登录云图平台

- `check_profile_exists()`: 检查 `browser_profiles/yuntu_profile/storage_state.json` 是否存在
- `perform_manual_login()`: 若 profile 不存在，打开可视化浏览器引导用户手动登录
  - 导航到云图登录页
  - 轮询 URL 变化检测登录完成（5 分钟超时）
  - 保存 `storage_state.json`（cookies + localStorage）

### 3.5 data_models.py — Pydantic 数据模型

**职责**: 定义结构化数据模型，映射云图平台的 6 大数据维度

```
YuntuDataReport (根模型)
├── 报告名称: str
├── 品牌名称: str
├── 日期范围: str
├── 项目整体: ProjectOverview          ← 21 个字段
├── 五A人群资产流转: A5AssetFlow       ← 12 个字段
├── 达人及内容复盘: KOLContentReview   ← Excel 关联数据
├── 搜索与溢出价值: SearchInsight      ← 3 个字段
├── 投流数据精细化复盘: AdFlowReview   ← Excel 关联数据
└── TA人群画像精准度复盘: TAPortraitReview ← 4 个字段
```

### 3.6 data_extractor.py — Agent 输出解析

**职责**: 用正则表达式从 Agent 的文本输出中提取结构化数据

- `extract_from_agent_history()`: 主入口，处理 `AgentHistoryList`
- 支持多报告分割：用正则按报告名拆分 Agent 输出
- 每个数据模块有独立的提取函数：
  - `_extract_project_overview()`: 20+ 正则模式
  - `_extract_a5_asset_flow()`: 12 个正则模式
  - `_extract_search_insight()`: 3 个正则模式
  - `_extract_ta_portrait()`: 自由文本提取
- `parse_number()` 工具函数：处理百分比、千位分隔符、中文单位（万、亿）

### 3.7 excel_processor.py — Excel 读取与指标计算

**职责**: 读取 Agent 下载的 Excel 文件，提取细分数据并计算衍生指标

**Excel 文件处理**:
- `find_downloaded_excel_files()`: 扫描下载目录，过滤 24 小时内的文件
- `identify_excel_file_type()`: 按文件名关键词分类（KOL 内容 vs 投流数据）
- `read_kol_content_excel()`: 用 pandas 读取所有 sheet，提取曝光/互动列，计算爆文率
- `read_ad_flow_excel()`: 提取人群包表现、A3 流转率

**指标计算**:
- `calculate_metrics()`: 从已有数据计算缺失的成本指标

```
CPM  = 消耗金额 / 曝光次数 × 1000
CPE  = 消耗金额 / 互动量
CPS  = 消耗金额 / 回搜次数
CPA3 = 消耗金额 / A3流转人数
ROI  = 本次活动转化金额 / 消耗金额
```

### 3.8 excel_exporter.py — Excel 报表生成

**职责**: 将结构化数据导出为最终 Excel 报表

**两种导出模式**:

| 模式 | 触发条件 | 特点 |
|------|---------|------|
| **模板模式**（优先） | `docs/黑玩复盘-template.xlsx` 存在 | 加载模板文件，保留原有格式和样式，写入固定行列位置 |
| **标准模式**（兜底） | 模板不存在 | 从零创建 Workbook，7 个 sheet，蓝色表头 |

**模板模式**的固定布局:
- 表头行: 第 5 行
- 数据行: 第 6 行起（多报告时每行一份）
- 列范围: 第 3-20 列，每列对应特定字段

**标准模式**的 7 个 Sheet:
1. 汇总 — 报告概览 + 各模块提取状态
2. 项目整体 — 指标名 / 数值 / 单位
3. 5A人群资产流转
4. 达人及内容复盘
5. 搜索与溢出价值
6. 投流数据精细化复盘
7. TA人群画像

### 3.9 utils.py — 辅助函数

- `parse_number()`: 健壮的数字解析（百分比、千位分隔符、中文单位）
- `parse_chinese_number()`: 中文数字转换（1.5万 → 15000）
- `format_percentage()`: 百分比格式化输出
- `find_downloaded_excel_files()`: 文件发现与年龄过滤
- `identify_excel_file_type()`: 文件名分类

---

## 4. 数据模型详细字段

### 4.1 ProjectOverview（项目整体）— 21 字段

**基础指标**:

| 字段 | 类型 | 说明 |
|------|------|------|
| 消耗金额 | float | 实际消耗金额（元） |
| 曝光次数 | int | 曝光次数 |
| 曝光人数 | int | 曝光人数 |
| 互动率 | float | 互动率（%） |
| 互动量 | int | 互动次数 |
| 七日回搜人数 | int | 7 日回搜人数 |
| 七日回搜率 | float | 七日回搜率（%） |
| 回搜次数 | int | 回搜次数 |
| 完播率 | float | 完播率（%） |
| 完播数 | int | 完播数 |
| A3流转人数 | int | A3 流转人数 |
| A3流转率 | float | A3 流转率（%） |
| 本次活动转化金额 | float | 转化金额（元） |
| 拉新人群规模 | int | 拉新人群规模 |
| 新客占比 | float | 新客占比（%） |

**成本指标**:

| 字段 | 类型 | 说明 | 计算公式 |
|------|------|------|---------|
| CPM | float | 每千次曝光成本 | 消耗 / 曝光 × 1000 |
| CPE | float | 每次互动成本 | 消耗 / 互动量 |
| CPS | float | 每次回搜成本 | 消耗 / 回搜次数 |
| CPA3 | float | 每个 A3 流转成本 | 消耗 / A3流转人数 |
| CP5A | float | 每个 5A 人群成本 | — |
| ROI | float | 投资回报率 | 转化金额 / 消耗 |

### 4.2 A5AssetFlow（5A 人群资产流转）— 12 字段

| 字段 | 类型 | 说明 |
|------|------|------|
| 投后人群规模 | int | 投放后人群规模 |
| 投后增长率 | float | 投放后增长率（%） |
| A1新增量级 | int | A1（了解）新增量 |
| A1流转率 | float | A1 流转率（%） |
| A2新增量级 | int | A2（吸引）新增量 |
| A2流转率 | float | A2 流转率（%） |
| A3新增量级 | int | A3（种草）新增量 |
| A3流转率 | float | A3 流转率（%） |
| A4新增量级 | int | A4（购买）新增量 |
| A4流转率 | float | A4 流转率（%） |
| A5新增量级 | int | A5（复购）新增量 |
| A5流转率 | float | A5 流转率（%） |

> **5A 模型**: 巨量引擎的人群资产分层——A1(了解) → A2(吸引) → A3(种草) → A4(购买) → A5(复购)

### 4.3 KOLContentReview（达人及内容复盘）

Excel 关联模型，数据主要来源于 Agent 下载的 Excel 文件：

| 字段 | 类型 | 说明 |
|------|------|------|
| excel_file_path | str | Excel 文件路径 |
| excel_downloaded | bool | 是否已下载 |
| 爆文数据 | dict | 爆文详情 |
| 达人效率对比 | dict | 不同层级达人 CPM/CPE/CPA3 |
| 内容数量 | int | 内容总数 |
| 爆文率 | float | 爆文率（%） |
| 看后搜索率 | float | 看后搜索率（%） |

### 4.4 SearchInsight（搜索与溢出价值）

| 字段 | 类型 | 说明 |
|------|------|------|
| SOV | float | Share of Voice（声量份额指数） |
| 搜索人数变化趋势 | float | 搜索人数环比变化（%） |
| 搜索次数变化趋势 | float | 搜索次数环比变化（%） |

### 4.5 AdFlowReview（投流数据精细化复盘）

Excel 关联模型：

| 字段 | 类型 | 说明 |
|------|------|------|
| excel_file_path | str | Excel 文件路径 |
| excel_downloaded | bool | 是否已下载 |
| 人群包表现 | dict | 各人群包表现详情 |
| 投放A3流转率 | float | 投放 A3 流转率（%） |
| 各人群包CPA3 | dict | 各人群包 CPA3 数据 |
| 人群精准度 | float | 人群精准度 |

### 4.6 TAPortraitReview（TA 人群画像精准度复盘）

| 字段 | 类型 | 说明 |
|------|------|------|
| 全触达人群基础画像 | dict | 八大人群、年龄、性别、城市线级 |
| 全触达人群内容偏好 | dict | 全触达人群的内容偏好 |
| 投后A3人群基础画像 | dict | A3 人群基础画像 |
| 投后A3人群内容偏好 | dict | A3 人群内容偏好 |

---

## 5. 工作流程

### 5.1 单报告模式（通用品牌）

```
用户: "获取施华蔻的复盘数据，报告名：赛马复盘0105，消耗金额星图50万竞价30万"
  │
  ├─→ 意图提取: report_name="赛马复盘0105", brand_name="施华蔻",
  │              consumption_amount_star=500000, consumption_amount_bid=300000
  │
  ├─→ 登录检查 → 复用 cookies
  │
  ├─→ 提示词: GENERIC_PROMPT_TEMPLATE + 变量填充
  │
  ├─→ Agent 执行: 导航云图 → 找到报告 → 提取 6 大模块数据 → 下载 Excel
  │
  ├─→ 数据解析 → 单个 YuntuDataReport
  │
  └─→ Excel 导出 → 施华蔻_赛马复盘0105_20260206_100000.xlsx
```

### 5.2 多分类模式（黑玩品牌）

```
用户: "获取黑玩洗护和美发的数据
       洗护星图报告：黑玩洗护星图1月，消耗50万；竞价报告：黑玩洗护竞价1月，消耗30万
       美发星图报告：黑玩美发星图1月，消耗40万；竞价报告：黑玩美发竞价1月，消耗25万"
  │
  ├─→ 意图提取: brand_categories=[
  │     {category_name="洗护", star="黑玩洗护星图1月", star_amount=500000, bid="黑玩洗护竞价1月", bid_amount=300000},
  │     {category_name="美发", star="黑玩美发星图1月", star_amount=400000, bid="黑玩美发竞价1月", bid_amount=250000}
  │   ]
  │
  ├─→ 提示词: HEYONE_PROMPT_TEMPLATE（遍历每个分类 + 每分类的星图/竞价报告）
  │
  ├─→ Agent 执行: 对每个分类的每个报告依次取数
  │
  ├─→ 数据解析 → list[YuntuDataReport]（4 份报告）
  │
  └─→ Excel 导出: 模板模式，每份报告一行（数据行 6、7、8、9）
```

### 5.3 多轮对话

系统支持跨轮次的上下文保持：

```
第 1 轮: "我要获取黑玩的数据"
  → 提取: brand_name="黑玩"

第 2 轮: "洗护星图报告是：黑玩洗护星图1月"
  → 合并: brand_name="黑玩" + brand_categories=[{star="黑玩洗护星图1月"}]

第 3 轮: "消耗金额是50万"
  → 合并: + consumption_amount_star=500000

第 4 轮: "开始执行"
  → 使用完整的合并意图执行任务
```

---

## 6. 配置与环境变量

### 必需

| 变量 | 说明 |
|------|------|
| `GOOGLE_API_KEY` | Google Gemini API Key |

### 可选

| 变量 | 默认值 | 说明 |
|------|--------|------|
| `BROWSER_PROFILE_DIR` | `browser_profiles/yuntu_profile` | 浏览器 profile 保存路径 |
| `BROWSER_DOWNLOADS_PATH` | `downloads` | 文件下载目录 |
| `GEMINI_MODEL` | `gemini-2.5-pro-preview-05-06` | Gemini 模型名称 |

### 目录结构

```
yuntu-data/
├── browser_profiles/
│   └── yuntu_profile/
│       └── storage_state.json      # 登录态（cookies）
├── data-output/
│   ├── excel_exports/              # 最终 Excel 报表
│   │   └── <品牌>_<报告>_<时间戳>.xlsx
│   └── tasks/                      # 每次任务的工作目录
│       └── task_YYYYMMDD_HHMMSS/
│           ├── conversations/      # Agent 对话历史
│           └── *.xlsx              # 下载的原始 Excel
└── docs/
    ├── 黑玩复盘-template.xlsx      # 导出模板
    ├── 施华蔻赛马复盘-template.xlsx
    ├── template_structure.json     # 模板结构元数据
    └── 复盘路径拆解.docx           # 详细取数路径文档
```

---

## 7. 如何使用

### 启动

```bash
cd examples/apps/yuntu-data

# 设置 API Key
export GOOGLE_API_KEY="your-gemini-api-key"

# 启动 Gradio Web UI
python app.py
```

浏览器打开后，在聊天框输入自然语言任务描述即可。

### 首次登录

首次使用时，系统会自动打开浏览器窗口，引导你手动登录巨量云图。登录成功后自动保存 cookies，后续无需重复登录。

### 任务描述示例

**简单模式**:
```
获取施华蔻赛马复盘0105的数据，消耗金额星图50万竞价30万
```

**黑玩多分类**:
```
获取黑玩品牌数据
洗护分类：星图报告「黑玩洗护星图1月」消耗50万，竞价报告「黑玩洗护竞价1月」消耗30万
美发分类：星图报告「黑玩美发星图1月」消耗40万，竞价报告「黑玩美发竞价1月」消耗25万
爆文加热时间：按周
搜索洞察时间：按月 + 近30天
```

---

## 8. 已知限制与优化方向

### 当前限制

1. **数据提取依赖正则**: Agent 输出格式不固定，正则可能遗漏或误提取
2. **Excel 读取**: 不支持 `.xls` 格式和加密文件
3. **模板模式**: 目前仅实现了「总体数据表现」sheet 的固定布局填充，其他 sheet（人群包投放效果、by达人表现、投放人群画像）尚未完成
4. **TA 画像**: 提取的是自由文本，缺少结构化字段
5. **错误恢复**: Agent 执行中断后没有断点续传机制

### 优化方向

1. **LLM 辅助解析**: 用 LLM 替代正则做输出解析，提高鲁棒性
2. **模板完善**: 补充更多 sheet 的固定布局填充
3. **Excel 增强**: 使用 `template_structure.json` 元数据驱动模板填充
4. **并行处理**: 多个分类/报告的数据提取并行化
5. **数据缓存**: 缓存已提取的数据，避免重复执行
6. **提示词优化**: 缩短提示词长度以降低 Token 消耗
7. **部分成功**: 支持部分数据提取成功时的导出

---

## 附录：文件清单

| 文件 | 大小 | 职责 |
|------|------|------|
| `app.py` | ~33KB | Gradio UI + 主编排 |
| `config.py` | ~22KB | 配置 + 提示词模板 |
| `excel_exporter.py` | ~17KB | Excel 报表生成 |
| `excel_processor.py` | ~12KB | Excel 读取 + 指标计算 |
| `data_extractor.py` | ~11KB | Agent 输出正则解析 |
| `intent.py` | ~9KB | 意图提取（Gemini） |
| `data_models.py` | ~7KB | Pydantic 数据模型 |
| `utils.py` | ~5KB | 辅助函数 |
| `login.py` | ~3KB | 浏览器登录管理 |
| `__init__.py` | ~0.1KB | 包初始化 |
