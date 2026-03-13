# Yuntu Data

`yuntu-data` 是一个面向巨量云图的数据采集应用。

它的目标是：

- 根据用户自然语言输入识别品牌、报告名、消耗金额、时间范围
- 自动登录并进入巨量云图页面
- 按 DSL 配置执行页面导航、取数、下载文件
- 输出结构化结果和业务汇总结果

## 文档入口

先看这 2 个文档：

- 使用文档：[RUN.md](browser-use/examples/apps/yuntu-data/RUN.md)
- 架构文档：[ARCHITECTURE.md](browser-use/examples/apps/yuntu-data/ARCHITECTURE.md)

分工如下：

- `README.md`
  - 项目简介
  - 快速开始
  - 目录说明
- `RUN.md`
  - 用户安装与运行说明
  - 环境变量
  - 品牌映射
  - 输入样例
  - 输出与排障
- `ARCHITECTURE.md`
  - 整体架构
  - 执行链路
  - 关键文件职责

## 快速开始

### 1. 拉取代码并切换分支

```bash
git clone https://github.com/supermaty/browser-use.git
cd browser-use
git fetch origin
git checkout feat/mvp-opt
```

### 2. 安装依赖

项目统一使用 `uv`。

```bash
uv venv --python 3.12
```

激活虚拟环境后安装依赖：

```bash
uv sync --extra examples
```

更详细的 Windows / macOS 安装说明见：

- [RUN.md](browser-use/examples/apps/yuntu-data/RUN.md)

### 3. 配置环境变量

在项目根目录创建 `.env`：

```bash
GOOGLE_API_KEY=你的Google_API_Key
```

可选：

```bash
BROWSER_PROFILE_DIR=browser_profiles/yuntu_profile
BROWSER_DOWNLOADS_PATH=downloads
```

### 4. 配置品牌映射

品牌英文别名统一配置在：

- [config.py](browser-use/examples/apps/yuntu-data/config.py)

配置项：

```python
BRAND_NAME_MAPPING
```

### 5. 启动应用

在项目根目录执行：

```bash
python examples/apps/yuntu-data/app.py
```

或：

```bash
uv run python examples/apps/yuntu-data/app.py
```

启动后在浏览器打开：

```text
http://127.0.0.1:7860
```

## 典型输入

### 单报告

```text
品牌：冷酸灵；报告名：ZY_竞价整体_0224；星图消耗：10000；竞价消耗：10000；爆文加热：近30天；行业搜索洞察：自定义（2026.01.23 ~ 2026.02.14）
```

### 多品牌分类

```text
品牌：冷酸灵；品牌分类：洗护 星图报告ZY_星图整体_0228 10000 竞价报告ZY_竞价整体_0228 10000；美发 星图报告ZY_星图整体_0301 15000 竞价报告ZY_竞价整体_0301 20000；爆文加热：近30天；行业搜索洞察：自定义（2026.01.23 ~ 2026.02.14）
```

更多输入样例见：

- [RUN.md](browser-use/examples/apps/yuntu-data/RUN.md)

## 输出结果

每次任务会在 `data-output/tasks/<任务目录>/` 下生成结果文件，常见包括：

- `agent_structured_output.json`
- `business_summary.json`
- `execution_logs.txt`
- `conversations/`
- 下载的 `csv/xlsx` 文件

## 关键文件

用户通常关注：

- [app.py](browser-use/examples/apps/yuntu-data/app.py)
- [config.py](browser-use/examples/apps/yuntu-data/config.py)
- [RUN.md](browser-use/examples/apps/yuntu-data/RUN.md)

开发通常关注：

- [pipeline.py](browser-use/examples/apps/yuntu-data/pipeline.py)
- [prompt_builder.py](browser-use/examples/apps/yuntu-data/prompt_builder.py)
- [dsl.py](browser-use/examples/apps/yuntu-data/dsl.py)
- [specs.py](browser-use/examples/apps/yuntu-data/specs.py)
- [intent.py](browser-use/examples/apps/yuntu-data/intent.py)

## 什么时候改哪里

如果你只是使用应用：

- 改 `.env`
- 改 [config.py](browser-use/examples/apps/yuntu-data/config.py) 里的 `BRAND_NAME_MAPPING`

如果你要改取数路径或字段：

- 改 [specs.py](browser-use/examples/apps/yuntu-data/specs.py)

如果你要改执行提示：

- 改 [prompt_builder.py](browser-use/examples/apps/yuntu-data/prompt_builder.py)

如果你要改执行流程或结果汇总：

- 改 [pipeline.py](browser-use/examples/apps/yuntu-data/pipeline.py)

## 说明

这份 README 只保留入口信息，不展开所有细节。

详细安装、配置、输入、排障请看：

- [RUN.md](browser-use/examples/apps/yuntu-data/RUN.md)
