## 0. 概览

> `YuntuTask` 负责定义“用户想要什么”，`DSL_MODULE_OVERRIDES` 负责定义“应该走哪些页面和步骤”，`prompt_builder` 负责定义“Agent 如何理解这些步骤”，`tools.py` 负责定义“具体执行的工具”，
`pipeline.py` 负责定义“整个执行过程如何被约束、记录、恢复和汇总”。

所以它的执行并不是单靠 Prompt，也不是单靠 Tool，而是：

> **Intent -> DSL/Profile -> Prompt -> Agent/Tools -> Runtime Guards -> Structured Output ->  Summary**

## 1. 总体定位

`yuntu-data` 当前不是一个“纯工具流”系统，也不是一个“纯 Prompt 流”系统，而是一个混合架构：

1. `intent.py` 负责把用户输入解析成结构化任务。
2. `specs.py` + `dsl.py` 负责把 DSL 配置编译成运行时模块配置。
3. `prompt_builder.py` 负责把任务和模块配置拼成 Agent 任务提示词。
4. `tools.py` 负责浏览器动作、选择、下载、提取、Excel 求和等具体执行能力。
5. `pipeline.py` 负责调度 Agent、维护运行状态、做 guard、落日志、做中断恢复和结果汇总。
6. `completeness.py`、`excel_processor.py` 负责后处理校验和业务计算。

系统的真实执行模式是：

> DSL 提供流程骨架，Prompt 告诉 Agent 如何按骨架执行，Tools 提供确定性动作，Pipeline 在运行时做约束、补救和结果汇总。


## 2. 核心代码入口

| 层级 | 主要文件 | 关键入口 | 作用 |
| --- | --- | --- | --- |
| 任务入口 | `pipeline.py` | `OptTaskPipeline.process()` | 端到端主流程 |
| 意图解析 | `intent.py` | `extract_intent()` | 多轮对话 -> `YuntuTask` |
| DSL 配置 | `specs.py` | `DSL_MODULE_OVERRIDES` | 定义模块、路径、提取和下载 |
| DSL 编译 | `dsl.py` | `parse_route_dsl()` | DSL -> `route_path + interaction_steps` |
| 运行时配置构建 | `specs.py` | `build_opt_profile()` | DSL -> `BrandProfileSpec` |
| Prompt 构建 | `prompt_builder.py` | `build_agent_prompt()` | 任务 + 配置 -> Agent 提示词 |
| 工具注册 | `tools.py` | `create_opt_tools()` | 注册动作工具 |
| 阶段执行 | `pipeline.py` | `_run_phase()` | 单次 Agent 运行和 guard |
| 后处理 | `pipeline.py` | `_handle_post_processing()` | 输出 JSON、恢复、汇总 |
| 完整性校验 | `completeness.py` | `check_completeness()` | 检查模块/字段/文件是否齐全 |
| Excel 计算 | `excel_processor.py` | `sum_search_count_from_excel()` | 从下载文件中计算回搜次数等 |


## 3. 配置层架构

### 3.1 DSL 配置源

所有模块的主配置入口在 `DSL_MODULE_OVERRIDES`。

对应代码：

- `examples/apps/yuntu-data/specs.py:163`

当前 DSL 配置表达的信息包括：

1. 模块名
2. 是否依赖报告页面 `requires_report`
3. 路径骨架 `route_dsl`
4. 路径中的点击节点
5. 步骤中的 `[select]`
6. 步骤中的 `[set_date_range]`
7. 步骤中的 `[extract]`
8. 步骤中的 `[hover]`
9. 步骤中的 `[download]`

例如：

- `项目整体Overview`
- `5A人群资产流转`
- `人群包投放效果分析`
- `投放人群画像-5A人群总资产`
- `投放人群画像-A3问询`
- `达人及内容复盘`
- `行业搜索洞察`


### 3.2 DSL 编译成运行时模块

系统不会直接执行 DSL 文本，而是先把 DSL 编译成 `ModuleSpec`。

相关代码：

- `examples/apps/yuntu-data/dsl.py:15`
- `examples/apps/yuntu-data/specs.py:319`

编译过程包括：

1. `parse_route_dsl()` 把字符串 DSL 拆成：
   - `route_path`
   - `interaction_steps`
2. `_build_module()` 自动推导：
   - `field_keys`
   - `file_keys`
   - `acquisition_mode`
3. `build_opt_profile()` 把全部模块合并成 `BrandProfileSpec`

因此，运行时真正使用的不是原始 DSL 字符串，而是：

1. `BrandProfileSpec`
2. `ModuleSpec`
3. `InteractionStep`
4. `FieldSpec`
5. `FileSpec`


## 4. 执行层架构

### 4.1 运行入口

端到端入口是 `OptTaskPipeline.process()`。

对应代码：

- `examples/apps/yuntu-data/pipeline.py:3425`

这个函数负责：

1. 停掉旧 agent
2. 提取当前用户消息
3. 调 `extract_intent()`
4. 校验 intent
5. 检查登录状态
6. 构建 profile
7. 构建 prompt
8. 创建 Browser
9. 创建 Agent
10. 调 `_run_phase()`
11. 最后做 `_handle_post_processing()`


### 4.2 Agent 运行模式

Agent 初始化在：

- `examples/apps/yuntu-data/pipeline.py:3527`

当前关键参数包括：

1. `task=build_agent_prompt(...)`
2. `tools=self._build_tools("agent_first")`
3. `max_actions_per_step=1`
4. `flash_mode=True`
5. `use_thinking=False`
6. `use_vision=False`
7. `max_history_items=6`
8. `llm_timeout=180`
9. `step_timeout=210`

这说明当前系统的设计目标是：

1. 尽量单步执行
2. 限制上下文长度
3. 把浏览器执行交给工具和轻量 guard
4. 避免长思考链


### 4.3 Tools 层

工具注册入口：

- `examples/apps/yuntu-data/tools.py:9098`

`create_opt_tools()` 会把浏览器动作能力注册成可调用工具。结合 `pipeline.py` 和 `prompt_builder.py`，当前关键工具包括：

1. `click_in_content`
2. `select_in_content`
3. `download_files_by_spec`
4. `sum_search_count_in_date_range`
5. 一些辅助导航/选择/页面判断工具

另外，Prompt 明确要求 Agent 直接用原生动作做两类事：

1. `extract`
2. `evaluate`

也就是说当前系统是“工具 + 原生 Agent 动作”混用的。


## 5. 端到端执行过程

下面按真实代码顺序描述一次任务运行。

### Step 1. 接收用户输入

入口：

- `examples/apps/yuntu-data/pipeline.py:3425`

用到的内容：

1. `history`
2. 当前最后一条 user message
3. `TaskSession`

输出：

1. 当前轮任务准备开始


### Step 2. 多轮意图解析

入口：

- `examples/apps/yuntu-data/intent.py:271`

用到的内容：

1. 当前用户消息 `task_description`
2. 最近若干轮对话 `conversation_history`
3. 上一轮 intent `previous_intent`
4. `ChatGoogle`
5. `INTENT_EXTRACTION_SYSTEM_PROMPT`

做的事情：

1. 让 LLM 提取结构化任务参数
2. 解析出 `YuntuTask`
3. 做多报告 fallback
4. 保留上轮已经识别过但本轮没覆盖的字段

输出：

1. `YuntuTask`


### Step 3. Intent 校验

入口：

- `examples/apps/yuntu-data/validator.py`

在 `process()` 中调用：

- `examples/apps/yuntu-data/pipeline.py:3456`

用到的内容：

1. `YuntuTask`
2. 必填字段规则

输出：

1. 校验通过，继续执行
2. 校验失败，直接返回给用户


### Step 4. 登录态检查

入口调用：

- `examples/apps/yuntu-data/pipeline.py:3462`

用到的内容：

1. `check_profile_exists()`
2. `perform_manual_login()`

输出：

1. 有登录态则继续
2. 无登录态则要求人工登录


### Step 5. 构建运行时 Profile

入口：

- `examples/apps/yuntu-data/specs.py:409`

用到的内容：

1. `resolve_profile_name(task_intent)`
2. `build_opt_profile(profile_name)`
3. `DSL_MODULE_OVERRIDES`
4. `parse_route_dsl()`

做的事情：

1. 读取 DSL 模块配置
2. 解析每个模块的 `route_path`
3. 解析每个模块的 `interaction_steps`
4. 自动推导字段清单
5. 自动推导文件清单
6. 组装成 `BrandProfileSpec`

输出：

1. `profile`


### Step 6. 构建 Prompt

入口：

- `examples/apps/yuntu-data/prompt_builder.py:316`

用到的内容：

1. `task_intent`
2. `profile.module_specs`
3. 目标报告列表
4. 推断触点
5. insight cycles

做的事情：

1. 把 DSL 转成人类可读的执行计划
2. 给 Agent 约束动作边界
3. 规定 `[click]/[select]/[extract]/[hover]/[download]` 的执行方式
4. 规定重试规则、日期规则、hover 规则、route-tail 规则

输出：

1. 一整段 `task prompt`


### Step 7. 创建 Browser 和 Agent

入口：

- `examples/apps/yuntu-data/pipeline.py:3509`

用到的内容：

1. `Browser`
2. 浏览器 profile / storage state
3. 下载目录 `task_dir`
4. `create_opt_tools()`
5. `build_agent_prompt()`

输出：

1. 准备好的 `browser`
2. 准备好的 `agent`


### Step 8. 进入运行阶段 `_run_phase()`

入口：

- `examples/apps/yuntu-data/pipeline.py:2246`

这是执行期最核心的函数。

用到的运行时状态包括：

1. `log_messages`
2. `runtime_extracted_contents`
3. `current_report_name`
4. `report_phase_confirmed`
5. `module_stage`
6. `empty_action_streak`
7. `hover_eval_seen_by_stage`
8. `extract_signature_seen_by_stage`
9. `pending_click_target`
10. `pending_click_retry_count`
11. `recent_success_click_targets`

它还会从 `profile.module_specs` 中预计算：

1. `report_modules`
2. `field_key_to_stage`
3. `file_key_to_stage`
4. `module_route_tokens_by_key`
5. `select_route_prereq_rules`
6. `click_route_prereq_rules`

作用：

1. 把“DSL 配置”映射成“运行时阶段状态”
2. 让系统知道当前执行到哪个模块
3. 防止跳步、重复提取、回退


### Step 9. 每一步开始 `on_step_start`

入口：

- `examples/apps/yuntu-data/pipeline.py` 中 `_run_phase()` 内部定义的 `on_step_start`

用到的内容：

1. 当前 step 编号
2. `log_messages`

作用：

1. 写 step 开始日志
2. 控制日志缓存


### Step 10. 每一步前的 guard

`_run_phase()` 会在执行前注册 pre-step guard。

相关代码：

- `examples/apps/yuntu-data/pipeline.py:2246`
- `examples/apps/yuntu-data/pipeline.py` 中 `_pre_action_guard`
- `examples/apps/yuntu-data/pipeline.py` 中 `_pre_action_pending_click_retry_guard`

这些 guard 用到的内容：

1. 当前模块 stage
2. 最近成功点击的 route token
3. 待重试点击目标
4. extract signature
5. 是否是日期输入
6. 是否是回退动作

主要作用：

1. 阻止不允许的原生动作
2. route tail 未到位时，先要求重新定位
3. 防止连续重复 extract
4. 防止同一个点击目标无限重试
5. 防止回退到已完成的更早模块


### Step 11. Agent 输出动作

Agent 每一步会根据 prompt 和当前页面，生成一组动作。

当前允许的主动作分两类。

原生动作：

1. `extract`
2. `evaluate`
3. `wait`
4. `done`

工具动作：

1. `click_in_content`
2. `select_in_content`
3. `download_files_by_spec`
4. `sum_search_count_in_date_range`

限制来源：

- `examples/apps/yuntu-data/pipeline.py:2270`
- `examples/apps/yuntu-data/prompt_builder.py:400`


### Step 12. 工具或原生动作真正执行

执行主体：

1. Agent 原生 `extract/evaluate`
2. Tools 中注册的动作

Tools 用到的内容通常包括：

1. DOM 选择器
2. 目标文案
3. 模块配置
4. `TaskSession`
5. 下载目录
6. 浏览器页面对象

这一步会产出两类结果：

1. 原生文本结果，比如 `<query>...</query><result>...</result>`
2. 工具结构化结果，比如 `{"type":"field_extraction_result", ...}`


### Step 13. 每一步结束 `on_step_end`

入口：

- `_run_phase()` 内部 `on_step_end`

用到的内容：

1. `agent_instance.state.last_model_output`
2. `agent_instance.state.last_result`
3. `module_stage`
4. `current_report_name`
5. `runtime_extracted_contents`
6. `pending_click_target`

做的事情：

1. 记录动作日志
2. 识别当前 report
3. 更新 `module_stage`
4. 记录成功点击目标
5. 记录失败点击，准备重试
6. 对 select 失败做 route-tail 回退
7. 保存 `runtime_extracted_contents`
8. 遇到关键结果时落 checkpoint


### Step 14. 执行过程中间 checkpoint

入口调用：

- `examples/apps/yuntu-data/pipeline.py:3340`

用到的内容：

1. `runtime_extracted_contents`
2. `task_dir`
3. 当前 report
4. 当前 step 编号

作用：

1. 中途中断时尽量保留已拿到的数据和文件


### Step 15. 运行完成后进入后处理

入口：

- `examples/apps/yuntu-data/pipeline.py:3627`

用到的内容：

1. `result.final_result()`
2. `result.extracted_content()`
3. `task_dir`
4. `log_text`
5. `profile`
6. `task_intent`

输出文件：

1. `agent_output.json`
2. `raw_agent_json`
3. `agent_structured_output.json`
4. `business_summary.json`
5. `execution_logs.txt`
6. `conversations/`


## 6. 后处理与结果构建

### 6.1 计算周期提取

入口：

- `examples/apps/yuntu-data/pipeline.py:286`
- `examples/apps/yuntu-data/pipeline.py:350`

用到的内容：

1. `open_report_result`
2. `calculation_period`
3. 报告名归一化

作用：

1. 把报告页提取到的 `计算周期/计算时间区间` 同步到 `task_intent`
2. 供后面回搜次数和 summary 计算使用


### 6.2 部分结果恢复

入口：

- `examples/apps/yuntu-data/pipeline.py:1798`

用到的内容：

1. 工具结果
2. 原生 extract 结果
3. 扫描到的下载文件
4. `profile.module_specs`

作用：

1. 构建 `recovery_output`
2. 即使 structured output 失败，也尽量输出已取到的数据


### 6.3 Structured Output 对齐与合并

入口：

- `examples/apps/yuntu-data/pipeline.py:2006`
- `examples/apps/yuntu-data/pipeline.py:1486`

用到的内容：

1. `structured_output`
2. `recovery_output`
3. tool payload
4. native extract hint
5. `report_targets`

作用：

1. 合并主结果和恢复结果


### 6.4 业务汇总 JSON

入口：

- `examples/apps/yuntu-data/pipeline.py:1168`

用到的内容：

1. `structured_output`
2. `profile.module_specs`
3. `task_intent`
4. `report_period_map`
5. `output_dir`

作用：

1. 把 report 模块和 global 模块汇总成 `business_summary.json`
2. 把 label 作为业务输出键
3. 针对每个 report 补充计算指标

输出格式核心结构：

1. `结案报告`
2. 每个报告一行
3. 行内包含：
   - `品牌`
   - `品牌分类`
   - `报告名`
   - 各模块业务对象


### 6.5 完整性校验

入口：

- `examples/apps/yuntu-data/completeness.py:48`

用到的内容：

1. `structured_output`
2. `profile`
3. 所有模块应有字段
4. 所有模块应有文件

作用：

1. 输出缺失模块
2. 输出缺失字段
3. 输出缺失文件


## 7. Excel 与业务指标计算

### 7.1 搜索次数求和

入口：

- `examples/apps/yuntu-data/excel_processor.py:30`

用到的内容：

1. 下载得到的 Excel 文件
2. `date_start`
3. `date_end`

作用：

1. 在指定计算周期内，把 `搜索次数` 列累计起来


### 7.2 指标计算

入口：

- `examples/apps/yuntu-data/excel_processor.py:309`
- 业务汇总调用在 `pipeline.py` 的 `_apply_computed_metrics_for_report_row()`

当前业务指标来源包括：

1. 实际消耗金额
2. 曝光次数
3. 互动量
4. 回搜次数
5. A3流转人数
6. 转化金额

典型计算：

1. `互动次数 = 互动率 * 曝光次数`
2. `完播数 = 曝光次数 * 完播率`
3. `CPM = 实际消耗金额 / 曝光次数 * 1000`
4. `CPE = 实际消耗金额 / 互动量`
5. `CPS = 实际消耗金额 / 回搜次数`
6. `CPA3 = 实际消耗金额 / A3流转人数`
7. `ROI = 转化金额 / 实际消耗金额`


## 8. 当前系统里“每一步到底用了什么”

可以简化成下面这条链：

1. 用户消息
   - 用到 `history`
   - 输出 `task_description`

2. `extract_intent()`
   - 用到 LLM、对话历史、上一轮 intent
   - 输出 `YuntuTask`

3. `validate_intent()`
   - 用到 `YuntuTask`
   - 输出校验结果

4. `build_opt_profile()`
   - 用到 `DSL_MODULE_OVERRIDES`
   - 输出 `BrandProfileSpec`

5. `build_agent_prompt()`
   - 用到 `YuntuTask + BrandProfileSpec`
   - 输出 `task prompt`

6. `create_opt_tools()`
   - 用到 `TaskSession + profile + downloads_path`
   - 输出工具注册表

7. `Agent.run()`
   - 用到 prompt、browser、tools、LLM
   - 输出 history / result

8. `_run_phase()`
   - 用到模块配置、stage、guard、日志缓存
   - 输出运行日志和中间 payload

9. `_handle_post_processing()`
   - 用到 `result.final_result()`、`result.extracted_content()`
   - 输出结构化 JSON 和业务汇总

10. `check_completeness()`
    - 用到 `structured_output + profile`
    - 输出完整性结果

11. `sum_search_count_from_excel()` / 业务计算
    - 用到下载文件和周期
    - 输出回搜次数和衍生指标


## 9. 运行时产物

一次任务结束后，通常会在任务目录下看到：

1. `conversations/`
2. `execution_logs.txt`
3. `agent_output.json`
4. `agent_structured_output.json`
5. `business_summary.json`
6. 下载的 CSV / Excel 文件

其中：

1. `agent_output.json` 更接近原始执行结果
2. `agent_structured_output.json` 是结构化中间层
3. `business_summary.json` 是业务消费层
