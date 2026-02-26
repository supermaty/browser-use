# Yuntu Data Scraper (巨量云图数据助手)

基于 Browser Use + Gradio 的云图数据抓取应用。  
当前代码主链路是：自然语言任务 -> 意图提取 -> Agent 执行 -> JSON 导出。

## 当前状态（按最新代码）

### Done

- `Done` 聊天式任务输入与执行（Gradio UI）
- `Done` 意图识别（报告名、品牌、时间范围、消耗金额等）
- `Done` 登录态复用与首次手动登录引导
- `Done` 黑玩品牌专用提示词模板与多分类任务输入
- `Done` Agent 实时日志回显（步骤、动作、提取内容）
- `Done` 任务隔离目录输出（`data-output/tasks/task_时间戳/`）
- `Done` Agent 原始输出 JSON 导出
- `Done` 结构化数据 JSON 导出（Pydantic 模型）
- `Done` 回搜次数 Tool（按日期区间从下载的 Excel 累加搜索次数）

### Pending

- `Pending` 下载 Excel 的自动解析结果接入主流程（`excel_processor.py` 尚未在 pipeline 主链路中调用）
- `Pending` 指标计算（CPM/CPE/CPS/CPA3/ROI）接入主流程
- `Pending` 模板化 Excel 结果导出接入主流程（`excel_exporter.py` 已实现但未接入 pipeline）
- `Pending` 文档中部分“已实现 Excel 全流程”描述与实际代码完全对齐

## 环境要求

- Python 3.11+
- `.env` 中配置 `GOOGLE_API_KEY`

## 安装

```bash
# 在仓库根目录执行
uv sync --all-extras

# 或仅安装代码相关依赖
uv sync --extra code
```

## 运行

```bash
python examples/apps/yuntu-data/app.py
```

启动后访问：`http://127.0.0.1:7860`

## 示例任务

```text
从巨量云图获取黑玩复盘数据，报告名：ZY-Heyone12.1-12.31；
星图 10万、竞价 20万；
爆文加热：近30天；
行业搜索洞察：自定义（2025/12/01-2025/12/31）
```

## 首次运行说明

如果未检测到可用登录信息，应用会：

1. 打开浏览器并进入云图登录页
2. 等待你手动登录
3. 保存浏览器 Profile 到 `browser_profiles/yuntu_profile`
4. 后续任务复用登录态

## 输出产物（当前版本）

每次任务默认写入独立目录：`data-output/tasks/task_YYYYMMDD_HHMMSS/`

- `agent_output.json`：Agent 的原始 `final_result` 与 `extracted_content`
- `<品牌>_<报告>_<时间戳>.json`：从 Agent 输出中解析的 JSON（或 raw_text 兜底）
- `structured_data_<时间戳>.json`：结构化后的 `YuntuDataReport` 列表
- `conversations/`：对话与执行过程记录
- 任务内下载文件：Agent 下载的 Excel 等文件

## 配置项（`config.py`）

- `GEMINI_MODEL`（当前默认：`gemini-3-pro-preview`）
- `YUNTU_LOGIN_URL`
- `BROWSER_PROFILE_DIR`（可通过 `.env` 覆盖）
- `BROWSER_DOWNLOADS_PATH`（可通过 `.env` 覆盖）
- 黑玩/通用品牌提示词模板

## 参考链接

1. 黑玩复盘最终项目结果 sample Excel：  
https://k74u68pb5x.feishu.cn/wiki/KC10wNnjtiNndHk5SiQc34FunYc?sheet=Y7UCNB
2. 黑玩复盘数据获取 SOP 路径：  
https://k74u68pb5x.feishu.cn/wiki/I26owf0fdiRujlkFhMVcJsD8nHb
3. 巨量云图网站地址：  
https://yuntu.oceanengine.com/yuntu_brand/ecom/home/overview?aadvid=1804155605152852
