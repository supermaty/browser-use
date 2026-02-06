# Browser-Use 项目全景文档

## 1. 项目简介

**Browser-Use** 是一个异步 Python（>=3.11）库，让 AI 大语言模型（LLM）能够像人一样操作浏览器：导航网页、点击元素、输入文本、提取数据、下载文件——完成复杂的端到端任务。

核心思路：将浏览器的 DOM 状态（可访问性树 + 截图）序列化后交给 LLM 决策，LLM 返回动作指令，再通过 Chrome DevTools Protocol（CDP）执行。这个「感知 → 决策 → 执行」的循环由 Agent 自动驱动，直到任务完成或超过最大步数。

**版本**: 0.11.3 | **协议**: MIT | **作者**: Gregor Zunic

---

## 2. 架构总览

```
┌──────────────────────────────────────────────────────────────────┐
│                           用户代码                               │
│   agent = Agent(task="...", llm=ChatOpenAI(...))                 │
│   result = await agent.run()                                     │
└──────────────────┬───────────────────────────────────────────────┘
                   │
                   ▼
┌──────────────────────────────────────────────────────────────────┐
│                    Agent (编排层)                                 │
│   browser_use/agent/service.py                                   │
│                                                                  │
│   ┌─────────┐  ┌──────────────┐  ┌───────────────┐              │
│   │ Message  │  │ System       │  │ Judge         │              │
│   │ Manager  │  │ Prompt (.md) │  │ (任务校验)    │              │
│   └─────────┘  └──────────────┘  └───────────────┘              │
└──────────────────┬───────────────────────────────────────────────┘
                   │  LLM 决策 → AgentOutput (动作列表)
                   ▼
┌──────────────────────────────────────────────────────────────────┐
│                  Tools / Registry (动作层)                        │
│   browser_use/tools/service.py                                   │
│   browser_use/tools/registry/service.py                          │
│                                                                  │
│   内置动作: search, navigate, click, input_text, scroll,         │
│            send_keys, upload_file, switch_tab, extract, done...  │
│   自定义动作: @tools.action() 装饰器                              │
└──────────────────┬───────────────────────────────────────────────┘
                   │  执行动作 → 发出事件
                   ▼
┌──────────────────────────────────────────────────────────────────┐
│              BrowserSession (浏览器管理层)                        │
│   browser_use/browser/session.py                                 │
│                                                                  │
│   ┌──────────────┐  ┌─────────────┐  ┌────────────────────┐     │
│   │ CDP Client   │  │ Event Bus   │  │ Session Manager    │     │
│   │ (cdp-use)    │  │ (bubus)     │  │ (Target 管理)      │     │
│   └──────────────┘  └──────┬──────┘  └────────────────────┘     │
│                            │                                     │
│              ┌─────────────┼──────────────┐                      │
│              ▼             ▼              ▼                       │
│   ┌──────────────┐ ┌────────────┐ ┌──────────────┐              │
│   │ DOM Watchdog │ │ Downloads  │ │ Security     │  ... (12个)  │
│   │              │ │ Watchdog   │ │ Watchdog     │              │
│   └──────────────┘ └────────────┘ └──────────────┘              │
└──────────────────────────────────────────────────────────────────┘
                   │
                   ▼
┌──────────────────────────────────────────────────────────────────┐
│                 DomService (DOM 提取层)                           │
│   browser_use/dom/service.py                                     │
│                                                                  │
│   CDP DOMSnapshot + Accessibility API → 可访问性树 → LLM 消费    │
└──────────────────────────────────────────────────────────────────┘
```

整个系统是**事件驱动**的：BrowserSession 内部维护一个 `bubus` 事件总线，12 个 Watchdog 通过事件订阅/发布协调工作。Agent 层和 Tools 层也通过事件与 BrowserSession 通信。

---

## 3. 核心组件详解

### 3.1 Agent — 主编排器

**文件**: `browser_use/agent/service.py`

Agent 是整个系统的入口。用户给它一个自然语言任务描述，它驱动一个「感知 → 决策 → 执行」的循环：

1. **感知**: 通过 DomService 获取当前页面的 DOM 状态（可访问性树 + 可选截图）
2. **决策**: 将 DOM 状态 + 任务描述 + 历史消息发送给 LLM，获取 `AgentOutput`（包含一个或多个动作）
3. **执行**: 通过 Tools 层将动作映射到浏览器操作，依次执行
4. **循环**: 回到步骤 1，直到 LLM 返回 `DoneAction` 或达到 `max_steps`

**关键配置**（`AgentSettings`）:
- `max_steps`: 最大步数（默认 100）
- `max_failures`: 最大连续失败次数（默认 10）
- `use_vision`: 是否启用截图分析（`True` / `False` / `'auto'`）
- `flash_mode`: 禁用 LLM thinking 以提高速度
- `multi_act`: 是否允许单步执行多个动作

**辅助组件**:
- `MessageManager`（`browser_use/agent/message_manager/service.py`）: 管理对话历史、截断长消息、注入系统提示词
- `SystemPrompt`（`browser_use/agent/prompts.py`）: 加载 Markdown 格式的系统提示词模板，支持多种 LLM 变体
- `Judge`（`browser_use/agent/judge.py`）: 可选的 LLM 判官，用于验证任务是否正确完成

**关键数据模型**（`browser_use/agent/views.py`）:
- `AgentOutput`: LLM 返回的决策结果（当前状态评估 + 动作列表）
- `AgentHistory`: 单步执行记录（LLM 输出 + 执行结果 + 浏览器状态快照）
- `AgentHistoryList`: 完整的执行历史，提供 `.final_result()`, `.errors()`, `.model_actions()`, `.model_thoughts()` 等便捷方法
- `ActionResult`: 单个动作的执行结果（提取的内容、是否完成、错误信息等）

### 3.2 BrowserSession — 浏览器生命周期管理

**文件**: `browser_use/browser/session.py`

BrowserSession 封装了浏览器实例的完整生命周期：启动/连接、标签页管理、CDP 通信、事件协调。

**两种连接模式**:
- **本地模式**: 启动本地 Chrome/Chromium，支持 `executable_path`、`user_data_dir`、`headless` 等参数
- **远程模式**: 通过 `cdp_url` 连接已运行的浏览器（适合容器化/远程场景）

**CDP 集成**: 使用 `cdp-use` 库提供类型化的 CDP 接口：
```python
cdp_client.send.DOMSnapshot.enable(session_id=session_id)
cdp_client.send.Target.attachToTarget(params={'targetId': target_id, 'flatten': True})
cdp_client.register.Browser.downloadWillBegin(callback)
```

**SessionManager**（`browser_use/browser/session_manager.py`）: 事件驱动的会话池管理器，通过 Target attach/detach 事件自动同步浏览器状态，是所有 target 和 CDP session 的单一事实来源。

### 3.3 BrowserProfile — 浏览器配置

**文件**: `browser_use/browser/profile.py`

Pydantic 模型，封装所有浏览器启动参数：

- **显示配置**: 自动检测屏幕尺寸（macOS 用 `AppKit.NSScreen`，Linux/Windows 用 `screeninfo`）
- **扩展管理**: uBlock Origin、Cookie 处理器等默认扩展，支持白名单/黑名单
- **安全设置**: `disable_security`、`allowed_domains`、`prohibited_domains`
- **代理**: `ProxySettings` 支持 HTTP/SOCKS 代理
- **会话持久化**: `storage_state`（cookies + localStorage）、`user_data_dir`
- **录制**: `record_video_dir`、`record_har_path`
- **云浏览器**: `CloudBrowserParams`（profile ID、代理国家、超时）

### 3.4 Watchdog 系统 — 事件驱动的浏览器行为管理

**目录**: `browser_use/browser/watchdogs/`

每个 Watchdog 是一个独立的服务，通过 `bubus` 事件总线监听特定事件并作出响应。基类 `BaseWatchdog` 提供自动 handler 注册（按方法名 `on_EventName` 约定）和错误恢复。

| Watchdog | 文件 | 职责 |
|----------|------|------|
| DOMWatchdog | `dom_watchdog.py` | DOM 快照处理、截图、元素高亮 |
| DownloadsWatchdog | `downloads_watchdog.py` | PDF 自动下载、文件管理 |
| PopupsWatchdog | `popups_watchdog.py` | JS 对话框和弹窗处理 |
| SecurityWatchdog | `security_watchdog.py` | 域名限制和安全策略执行 |
| AboutBlankWatchdog | `aboutblank_watchdog.py` | 空白页重定向处理 |
| PermissionsWatchdog | `permissions_watchdog.py` | 浏览器权限请求管理 |
| CrashWatchdog | `crash_watchdog.py` | 浏览器崩溃检测与恢复 |
| StorageStateWatchdog | `storage_state_watchdog.py` | Cookies 和本地存储管理 |
| RecordingWatchdog | `recording_watchdog.py` | 视频/HAR 录制 |
| ScreenshotWatchdog | `screenshot_watchdog.py` | 截图捕获 |
| DefaultActionWatchdog | `default_action_watchdog.py` | 默认浏览器行为处理 |
| LocalBrowserWatchdog | `local_browser_watchdog.py` | 本地浏览器专属操作 |

### 3.5 Tools / Registry — 动作注册与执行

**文件**: `browser_use/tools/service.py`, `browser_use/tools/registry/service.py`

Tools 是一个泛型动作注册表（`Tools[Context]`），负责：
- 注册和管理所有可用动作（内置 + 自定义）
- 将 LLM 的输出（动作名 + 参数）映射到实际函数调用
- 注入特殊参数（`browser_session`, `file_system`, `llm` 等）

**内置动作**（`browser_use/tools/views.py`）:

| 动作 | 模型类 | 说明 |
|------|--------|------|
| `search` | `SearchAction` | 搜索网页（DuckDuckGo/Google/Bing） |
| `navigate` | `NavigateAction` | 导航到 URL，可选新标签页 |
| `click_element` | `ClickElementAction` | 通过元素索引或坐标点击 |
| `input_text` | `InputTextAction` | 向输入框填入文本 |
| `scroll` | `ScrollAction` | 滚动页面或特定元素 |
| `send_keys` | `SendKeysAction` | 发送键盘快捷键 |
| `upload_file` | `UploadFileAction` | 上传文件 |
| `switch_tab` | `SwitchTabAction` | 切换标签页 |
| `close_tab` | `CloseTabAction` | 关闭标签页 |
| `extract` | `ExtractAction` | 提取页面内容（支持链接提取） |
| `done` | `DoneAction` | 标记任务完成 |
| `get_dropdown_options` | `GetDropdownOptionsAction` | 获取下拉选项 |
| `select_dropdown_option` | `SelectDropdownOptionAction` | 选择下拉选项 |
| `screenshot` | `NoParamsAction` | 截取当前页面截图 |

### 3.6 DomService — DOM 提取与处理

**文件**: `browser_use/dom/service.py`

DomService 通过 CDP 的 `DOMSnapshot` 和 `Accessibility` API 提取页面结构，生成 LLM 可消费的文本表示：

- **DOM 快照**: 获取完整的 DOM 树，包括元素位置、可见性、计算样式
- **可访问性树**: 从 AX 节点构建可访问性树，包含角色（role）、名称（name）、属性
- **元素高亮**: 在页面上为可交互元素添加高亮标记和索引号
- **Iframe 处理**: 可配置的跨域 iframe 提取，支持深度和数量限制
- **坐标映射**: 处理高 DPI 显示和 iframe 嵌套时的坐标转换

**关键数据模型**（`browser_use/dom/views.py`）:
- `EnhancedDOMTreeNode`: 带位置和可见性的 DOM 节点
- `EnhancedAXNode`: 增强的可访问性树节点
- `SerializedDOMState`: 序列化后的 DOM 状态，直接传给 LLM

### 3.7 LLM 集成层 — 多提供者统一接口

**目录**: `browser_use/llm/`

通过 `BaseChatModel` 协议（`browser_use/llm/base.py`）统一所有 LLM 提供者的接口：

```python
class BaseChatModel(Protocol):
    async def ainvoke(self, messages, ...) -> BaseMessage: ...
    @property
    def model_name(self) -> str: ...
    @property
    def provider_name(self) -> str: ...
```

**支持的提供者**:

| 提供者 | 类名 | 目录 |
|--------|------|------|
| OpenAI | `ChatOpenAI` | `llm/openai/` |
| Anthropic | `ChatAnthropic` | `llm/anthropic/` |
| Google (Gemini) | `ChatGoogle` | `llm/google/` |
| Azure OpenAI | `ChatAzureOpenAI` | `llm/azure/` |
| Groq | `ChatGroq` | `llm/groq/` |
| Ollama | `ChatOllama` | `llm/ollama/` |
| Mistral | `ChatMistral` | `llm/mistral/` |
| Browser-Use | `ChatBrowserUse` | `llm/browser_use/` |
| Vercel | `ChatVercel` | `llm/vercel/` |
| OCI (Oracle) | `ChatOCIRaw` | `llm/oci_raw/` |
| Cerebras | `ChatCerebras` | `llm/cerebras/` |
| DeepSeek | `ChatDeepSeek` | `llm/deepseek/` |
| AWS Bedrock | — | `llm/aws/` |
| OpenRouter | — | `llm/openrouter/` |

**消息系统**（`browser_use/llm/messages.py`）: 提供统一的 `BaseMessage`、`SystemMessage`、`UserMessage`、`AssistantMessage`，支持文本和图片内容块。

### 3.8 MCP 集成 — 双向协议支持

**目录**: `browser_use/mcp/`

Browser-Use 同时支持 MCP 的两个方向：

**作为 MCP 服务器**（`mcp/server.py`）:
- 将浏览器自动化能力暴露为 MCP 工具
- 可集成到 Claude Desktop 等 MCP 客户端
- 启动方式: `uvx browser-use[cli] --mcp`

**作为 MCP 客户端**（`mcp/client.py`）:
- 连接外部 MCP 服务器（文件系统、GitHub 等）
- 自动发现和注册外部 MCP 工具为 browser-use 动作
- 通过 stdio 通信

---

## 4. 功能特性一览

### 浏览器操作
- 网页导航、搜索（DuckDuckGo/Google/Bing）
- 元素点击（索引或坐标）、文本输入、表单填写
- 页面滚动、键盘快捷键
- 文件上传/下载、PDF 自动处理
- 截图捕获

### 多标签页管理
- 创建/切换/关闭标签页
- 跨标签页任务协调

### 安全控制
- 域名白名单（`allowed_domains`）和黑名单（`prohibited_domains`），支持通配符
- 敏感数据处理：LLM 看到的是占位符，实际值在执行时替换
- 安全策略由 SecurityWatchdog 统一执行

### 视觉模式
- 截图分析：将页面截图作为图片传给多模态 LLM
- 支持 `vision_detail_level` 控制图片精度
- `use_vision='auto'` 自动检测 LLM 是否支持视觉

### 结构化输出
- 通过 Pydantic 模型定义期望的输出结构
- Agent 自动引导 LLM 按 schema 返回数据
- `StructuredOutputAction[T]` 泛型支持任意输出类型

### 视频录制
- 通过 `BrowserProfile.record_video_dir` 启用
- 录制完整的浏览器操作过程为 MP4

### 云浏览器
- 支持 Browser-Use Cloud 服务
- 代理、profile 持久化、地理限制突破

### 并行 Agent
- 多个 Agent 可共享同一 BrowserSession
- 多个独立 BrowserSession 可并行运行

### 其他
- 任务判定（Judge）：用独立 LLM 验证任务完成质量
- 历史回放（`rerun_history`）
- 回调钩子：`on_step_start`, `on_step_end` 等
- Fallback 模型：主模型失败时自动切换备选模型
- 文件系统访问：Agent 可读写本地文件

---

## 5. 使用方式

### 5.1 安装与配置

```bash
# 创建虚拟环境
uv venv --python 3.11
source .venv/bin/activate
uv sync

# 设置 API Key（以 OpenAI 为例）
export OPENAI_API_KEY="sk-..."
```

### 5.2 基础用法

```python
import asyncio
from browser_use import Agent, ChatOpenAI

async def main():
    agent = Agent(
        task="搜索 Python 3.12 的新特性，并提取前 5 个要点",
        llm=ChatOpenAI(model="gpt-4o"),
    )
    result = await agent.run()
    print(result.final_result())

asyncio.run(main())
```

### 5.3 自定义浏览器配置

```python
from browser_use import Agent, BrowserSession, BrowserProfile, ChatGoogle

session = BrowserSession(
    browser_profile=BrowserProfile(
        headless=False,                    # 显示浏览器窗口
        prohibited_domains=["ads.com"],    # 屏蔽域名
        downloads_path="./downloads",      # 下载目录
        keep_alive=True,                   # 保持浏览器开启
    )
)

agent = Agent(
    task="你的任务描述",
    llm=ChatGoogle(model="gemini-2.5-flash"),
    browser_session=session,
    max_steps=20,
)
```

### 5.4 自定义工具

```python
from browser_use import Agent, Tools, ActionResult, BrowserSession, ChatOpenAI

tools = Tools()

@tools.action("计算数学表达式")
async def calculate(expression: str) -> ActionResult:
    result = eval(expression)  # 示意，实际应使用安全求值
    return ActionResult(extracted_content=str(result))

@tools.action("获取当前时间")
async def get_time() -> ActionResult:
    import datetime
    return ActionResult(extracted_content=str(datetime.datetime.now()))

agent = Agent(
    task="计算 123 * 456 并告诉我结果",
    llm=ChatOpenAI(model="gpt-4o"),
    tools=tools,
)
```

### 5.5 结构化输出

```python
from pydantic import BaseModel
from browser_use import Agent, ChatOpenAI

class ProductInfo(BaseModel):
    name: str
    price: float
    rating: float

agent = Agent(
    task="从页面提取产品信息",
    llm=ChatOpenAI(model="gpt-4o"),
    output_model=ProductInfo,
)
result = await agent.run()
# result.final_result() 返回 ProductInfo 实例
```

### 5.6 MCP 服务器模式

```bash
# 作为 MCP 服务器启动，可与 Claude Desktop 集成
uvx browser-use[cli] --mcp
```

---

## 6. 扩展点

### 6.1 自定义 Tools

通过 `@tools.action()` 装饰器注册自定义动作：

```python
tools = Tools()

@tools.action(
    "在指定域名上执行特定操作",
    domains=["*.example.com"]  # 可选：域名过滤，仅在匹配域名时可用
)
async def my_action(browser_session: BrowserSession, param: str) -> ActionResult:
    # browser_session 自动注入
    page = await browser_session.get_current_page()
    # ... 自定义逻辑
    return ActionResult(extracted_content="结果")
```

支持注入的特殊参数：`browser_session`, `file_system`, `llm`, `page_extraction_llm`。

### 6.2 自定义 LLM

实现 `BaseChatModel` 协议即可接入任意 LLM：

```python
from browser_use.llm.base import BaseChatModel
from browser_use.llm.messages import BaseMessage

class MyLLM:
    async def ainvoke(self, messages: list[BaseMessage], **kwargs) -> BaseMessage:
        # 调用你的 LLM API
        ...

    @property
    def model_name(self) -> str:
        return "my-model"

    @property
    def provider_name(self) -> str:
        return "my-provider"
```

### 6.3 自定义 Watchdog

继承 `BaseWatchdog`，方法名以 `on_` + 事件名命名即可自动注册：

```python
from browser_use.browser.watchdogs.watchdog_base import BaseWatchdog

class MyWatchdog(BaseWatchdog):
    async def on_NavigationCompleteEvent(self, event):
        # 每次页面导航完成时触发
        ...
```

### 6.4 MCP 工具集成

连接外部 MCP 服务器，将其工具自动注册为 Agent 可用的动作：

```python
from browser_use.mcp.client import MCPClient

mcp_client = MCPClient(server_command=["npx", "some-mcp-server"])
# 外部工具自动变成 Agent 的动作
```

### 6.5 系统提示词自定义

```python
agent = Agent(
    task="...",
    llm=...,
    extend_system_message="你是一个专门处理电商任务的助手，优先使用搜索功能。",
)
```

### 6.6 BrowserProfile 自定义

支持完整的 Chrome 启动参数、扩展管理、代理、录制等配置。详见 `browser_use/browser/profile.py`。

---

## 7. 示例应用

### examples/apps/ — 完整应用

| 应用 | 说明 |
|------|------|
| **ad-use** | AI 广告生成器：分析着陆页 → 用 Gemini 生成 Instagram 图片广告 / 用 Veo3 生成 TikTok 视频广告 |
| **news-use** | 新闻监控：持续抓取新闻网站最新文章，带情感分析和去重 |
| **msg-use** | 消息/邮件自动化：浏览器登录流程 + 定时任务调度 |
| **yuntu-data** | 巨量云图数据助手：自动提取营销复盘数据并生成 Excel 报表 |

### examples/ — 按类别

| 类别 | 路径 | 内容 |
|------|------|------|
| **入门** | `getting_started/` | 5 步渐进式教程（搜索 → 表单 → 提取 → 多步任务 → 快速模式） |
| **功能演示** | `features/` | 26 个示例覆盖安全控制、多标签页、视频录制、结构化输出等 |
| **自定义函数** | `custom-functions/` | 9 个高级示例（2FA、文件上传、域名过滤、外部 API 集成） |
| **浏览器管理** | `browser/` | 6 个示例（真实浏览器、Cookie 保存、CDP 连接、云浏览器） |
| **LLM 模型** | `models/` | 20+ 模型适配示例（OpenAI、Anthropic、Google、Groq、Ollama 等） |
| **云 API** | `cloud/` | 5 个 Browser-Use Cloud 使用示例 |
| **集成** | `integrations/` | Gmail 2FA、Slack、Discord、AgentMail 集成 |
| **用例** | `use-cases/` | 10 个场景（求职、购物、CAPTCHA、PDF 提取、产品对比等） |
| **UI 框架** | `ui/` | 命令行、Gradio、Streamlit 三种界面示例 |
| **文件系统** | `file_system/` | Agent 读写文件、Excel 操作 |
| **可观测性** | `observability/` | OpenTelemetry 集成 |

---

## 8. 开发指南

### 环境搭建

```bash
uv venv --python 3.11
source .venv/bin/activate
uv sync
```

### 测试

```bash
# 运行 CI 测试套件
uv run pytest -vxs tests/ci

# 运行所有测试
uv run pytest -vxs tests/

# 运行单个测试
uv run pytest -vxs tests/ci/test_specific_test.py
```

### 质量检查

```bash
# 类型检查
uv run pyright

# Lint + 格式化
uv run ruff check --fix
uv run ruff format

# Pre-commit hooks
uv run pre-commit run --all-files
```

### 代码风格约定

- 使用 **Tab** 缩进（非空格）
- 使用 Python >= 3.12 类型语法：`str | None` 而非 `Optional[str]`，`list[str]` 而非 `List[str]`
- 使用 Pydantic v2 模型，`ConfigDict(extra='forbid')` 用于严格校验
- 核心逻辑放 `service.py`，数据模型放 `views.py`，事件定义放 `events.py`
- 日志方法以 `_log_` 前缀命名，与业务逻辑分离
- ID 字段使用 `uuid7str()`
- 函数首尾使用运行时断言（`assert`）验证约束

### 测试约定

- 永不 mock 真实对象（LLM 除外），使用 `pytest-httpserver` 搭建本地测试服务器
- 不使用真实远程 URL
- 异步测试直接使用 `async def`，无需 `@pytest.mark.asyncio`
- 通过的测试移入 `tests/ci/` 作为 CI 默认测试集

### 目录结构约定

```
browser_use/
├── agent/              # Agent 编排层
│   ├── service.py      # Agent 主类
│   ├── views.py        # 数据模型
│   ├── prompts.py      # 提示词管理
│   ├── judge.py        # 任务校验
│   ├── message_manager/# 消息管理
│   └── system_prompts/ # Markdown 提示词模板
├── browser/            # 浏览器管理层
│   ├── session.py      # BrowserSession
│   ├── profile.py      # BrowserProfile
│   ├── session_manager.py
│   ├── events.py       # 浏览器事件
│   └── watchdogs/      # 12 个 Watchdog
├── tools/              # 动作层
│   ├── service.py      # Tools 主类
│   ├── views.py        # 内置动作模型
│   └── registry/       # 动作注册表
├── dom/                # DOM 提取层
│   ├── service.py      # DomService
│   └── views.py        # DOM 数据模型
├── llm/                # LLM 集成层
│   ├── base.py         # BaseChatModel 协议
│   ├── messages.py     # 统一消息类型
│   ├── openai/         # OpenAI 适配
│   ├── anthropic/      # Anthropic 适配
│   ├── google/         # Google 适配
│   └── ...             # 其他提供者
├── mcp/                # MCP 协议集成
│   ├── server.py       # MCP 服务器
│   └── client.py       # MCP 客户端
├── filesystem/         # 文件系统访问
├── skills/             # 技能系统
├── screenshots/        # 截图服务
└── config.py           # 全局配置
```
