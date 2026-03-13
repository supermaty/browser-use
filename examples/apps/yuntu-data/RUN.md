# Yuntu Data 使用文档

本文档面向使用 `examples/apps/yuntu-data` 的业务/运营同学，说明如何完成本地环境准备、配置品牌映射、启动页面并输入任务。

## 1. 适用范围

当前应用用于从巨量云图中执行以下类型的数据采集：

- 结案报告取数
- 人群分析 / 画像分析文件下载
- 达人及内容复盘文件下载
- 行业搜索洞察按时间范围取数
- 结果汇总输出为结构化 JSON

应用入口文件：

- [app.py](browser-use/examples/apps/yuntu-data/app.py)

核心配置文件：

- [config.py](browser-use/examples/apps/yuntu-data/config.py)

## 2. 运行环境配置

### 2.1 从 GitHub 拉取代码并切换分支

如果本地还没有项目代码，先执行：

```bash
git clone https://github.com/supermaty/browser-use.git
cd browser-use
git fetch origin
git checkout feat/mvp-opt
```

如果已经有仓库，可以直接在项目根目录执行：

```bash
git fetch origin
git checkout feat/mvp-opt
git pull origin feat/mvp-opt
```

### 2.2 安装 uv

`uv` 是本项目统一使用的 Python 包和虚拟环境管理工具。

#### Windows

PowerShell 执行：

```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

安装完成后，重新打开终端，验证：

```powershell
uv --version
```

#### macOS

终端执行：

```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

安装完成后，重新打开终端，验证：

```bash
uv --version
```

### 2.3 配置 Node.js / npm 环境

当前 `yuntu-data` 主流程以 Python 为主，但建议本地一并准备好 `Node.js` 和 `npm` 环境，避免后续安装、调试或运行相关前端依赖时缺环境。

参考链接：

- https://zhuanlan.zhihu.com/p/7314838716

安装完成后，可在终端验证：

```bash
node -v
npm -v
```

### 2.4 创建虚拟环境并安装依赖

在项目根目录 `browser-use` 执行：

```bash
uv venv --python 3.12
```

激活虚拟环境：

#### Windows

```powershell
.\.venv\Scripts\activate
```

#### macOS

```bash
source .venv/bin/activate
```

安装当前应用需要的依赖：
```bash
uv sync --all-extras
```

或者如果只需要code依赖（包含pandas和openpyxl）
```bash
uv sync --extra code
```

说明：

- 项目要求 Python `>=3.11,<4.0`
- `yuntu-data` 用到的 `gradio`、`pandas`、`openpyxl` 都包含在 `examples` 依赖组里

## 3. 环境变量配置

在项目根目录创建 `.env` 文件。

最少需要：

```bash
GOOGLE_API_KEY=你的Google_API_Key
```

可选配置：

```bash
BROWSER_PROFILE_DIR=browser_profiles/yuntu_profile
BROWSER_DOWNLOADS_PATH=downloads
```

说明：

- `GOOGLE_API_KEY`：当前意图识别和执行链路依赖 Google 模型
- `BROWSER_PROFILE_DIR`：浏览器登录态保存目录；不配则默认使用项目下的 `browser_profiles/yuntu_profile`
- `BROWSER_DOWNLOADS_PATH`：浏览器下载目录；不配则默认使用项目下的 `downloads`

配置读取位置：

- [config.py](browser-use/examples/apps/yuntu-data/config.py)

## 4. 品牌映射配置

品牌英文名映射不要再写到别的文件里，统一配置在：

- [config.py](browser-use/examples/apps/yuntu-data/config.py)

当前配置项名称：

```
BRAND_NAME_MAPPING
```

用途：

- 用于页面右上角切换品牌时，中文品牌名找不到时，补充英文品牌名作为兜底
- `get_brand_english_name()` 会优先从这个映射里取值

示例：

```python
BRAND_NAME_MAPPING: dict[str, str] = {
    "黑玩": "Heyone",
    "施华蔻": "Schwarzkopf",
    "冷酸灵": "Leng Suan Ling",
}
```

建议：

- 如果某个品牌在页面里显示的是英文名或别名，先把它补到这里
- 品牌切换失败时，优先检查这里

## 5. 启动方式

在项目根目录执行：

```bash
python examples/apps/yuntu-data/app.py
```

或者使用 `uv run`：

```bash
uv run python examples/apps/yuntu-data/app.py
```

启动后，终端会输出本地访问地址，通常是：

```bash
http://127.0.0.1:7860
```

在浏览器中打开即可使用。

## 6. 首次运行说明

首次运行时，如果没有已保存的登录态，应用会要求你完成一次手动登录。

默认行为：

- 打开浏览器
- 进入巨量云图登录页
- 等待你手动登录
- 登录成功后保存会话到本地 profile 目录

默认登录态目录：

- `browser_profiles/yuntu_profile`

注意：

- 建议先关闭可能占用同一个 Chrome Profile 的浏览器实例
- 如果你更换了 profile 目录，需要同步更新 `.env` 里的 `BROWSER_PROFILE_DIR`

## 7. 输入样例

建议直接输入完整自然语言，包含品牌、报告名、消耗金额、时间范围等关键信息。

### 7.1 单报告示例

```text
品牌：冷酸灵；报告名：ZY_竞价整体_0224；星图消耗：15000；竞价消耗：10000；爆文加热：近30天；行业搜索洞察：自定义（2026.01.23 ~ 2026.02.14）
```

### 7.2 单品牌多报告示例

```text
品牌：冷酸灵；星图报告：ZY_星图整体_0224；星图消耗：15000；竞价报告：ZY_竞价整体_0224；竞价消耗：10000；爆文加热：近30天；行业搜索洞察：自定义（2026.01.23 ~ 2026.02.14）
```

### 7.4 时间格式建议

推荐统一使用以下格式：

- `2026.01.23 ~ 2026.02.14`
- `2026/01/23-2026/02/14`
- `近30天`
- `近7天`
- `按周`
- `按月`
- `自定义（2026.01.23 ~ 2026.02.14）`

## 8. 输出结果说明

每次任务会在 `data-output/tasks/任务目录/` 下生成一组文件。

常见输出包括：

- `agent_structured_output.json`
  - 结构化采集结果
- `business_summary.json`
  - 面向业务查看的汇总结果
- `execution_logs.txt`
  - 逐步执行日志
- `conversations/`
  - 每一步模型输出与浏览器状态快照
- 下载的 CSV / Excel 文件
  - 例如：人群包投放效果分析_20260312_152535.csv

如果任务中间中断，也会尽量保留：

- `agent_structured_output.partial.json`
- `business_summary.partial.json`
- `checkpoints/`

## 9. 你通常需要改哪些地方

普通用户一般只需要改 2 个地方：

1. `.env`
   - 配 `GOOGLE_API_KEY`
   - 可选配 `BROWSER_PROFILE_DIR` / `BROWSER_DOWNLOADS_PATH`
2. [config.py](browser-use/examples/apps/yuntu-data/config.py)
   - 维护 `BRAND_NAME_MAPPING`

通常不需要改：

- `pipeline.py`
- `prompt_builder.py`
- `dsl.py`
- `specs.py`

如果你只是新增品牌，优先改 `BRAND_NAME_MAPPING`，不要直接改 DSL。

## 10. 常见问题

### 10.1 启动时报 `GOOGLE_API_KEY is not set`

原因：

- `.env` 未配置
- 或终端当前目录不在项目根目录

处理：

- 在项目根目录创建/检查 `.env`
- 确认存在：

```bash
GOOGLE_API_KEY=你的Google_API_Key
```

### 10.2 页面切换品牌失败

优先检查：

- [config.py](browser-use/examples/apps/yuntu-data/config.py) 里的 `BRAND_NAME_MAPPING`

如果页面显示的是英文品牌名或别名，而输入的是中文品牌名，需要补上映射。

### 10.3 首次运行无法复用登录态

检查：

- `.env` 中的 `BROWSER_PROFILE_DIR`
- 默认目录下是否已生成 `storage_state.json`

### 10.4 下载文件后找不到

检查：

- `.env` 中的 `BROWSER_DOWNLOADS_PATH`
- `data-output/tasks/当前任务目录/`
- 项目根目录下的 `downloads/`

### 10.5 取数结果和页面不一致

优先查看：

- `execution_logs.txt`
- `conversations/`
- `agent_structured_output.json`

先确认是：

- 页面没点到正确位置
- 字段没提到
- 还是汇总阶段覆盖了结果

## 11. 其他说明

### 11.1 当前模型

当前配置在 [config.py](browser-use/examples/apps/yuntu-data/config.py)：

```python
GEMINI_MODEL = "gemini-3-flash-preview"
```

如果后续切模型，先确认对应 API Key 和依赖链路。

### 11.2 输出目录

当前基础输出目录：

```python
OUTPUT_BASE_DIR = "./data-output"
```

任务目录：

```python
TASKS_BASE_DIR = Path(OUTPUT_BASE_DIR) / "tasks"
```

### 11.3 默认下载目录

当前默认下载目录来自：

```python
BROWSER_DOWNLOADS_PATH
```

未配置时默认是项目根目录下的：

```text
downloads
```

## 12. 推荐使用方式

推荐顺序：

1. 配好 `.env`
2. 在 `config.py` 里补齐 `BRAND_NAME_MAPPING`
3. 启动应用
4. 先用单报告样例验证
5. 再跑多报告

这样排查成本最低。



