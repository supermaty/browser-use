# pyright: reportMissingImports=false
"""
Yuntu Data Scraper - Gradio Application

A chat-based UI for extracting data from Yuntu (巨量云图) using Browser Use.
"""
import asyncio
import os
import sys
from pathlib import Path

# 避免代理导致 Gradio launch 时对 127.0.0.1 的 startup-events 请求返回 502（须在 import gradio 前设置）
_for_no_proxy = "127.0.0.1,localhost"
for _k in ("NO_PROXY", "no_proxy"):
	_existing = os.environ.get(_k, "")
	os.environ[_k] = _for_no_proxy if not _existing else f"{_existing},{_for_no_proxy}"

# Add project root to path to ensure browser_use imports work
PROJECT_ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(PROJECT_ROOT))

from dotenv import load_dotenv
load_dotenv()

import gradio as gr  # type: ignore
from browser_use import Agent, Browser, ActionResult, Tools
from browser_use.llm.google import ChatGoogle

# Add app directory to path for local imports
APP_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(APP_DIR))

from config import (
    PROFILE_DIR,
    HEYONE_PROMPT_TEMPLATE,
    GENERIC_PROMPT_TEMPLATE,
    GEMINI_MODEL,
    OUTPUT_BASE_DIR,
    API_KEY,
    STORAGE_STATE_FILE,
    DOWNLOADS_PATH,
    EXCEL_OUTPUT_DIR,
    TASKS_BASE_DIR,
    get_brand_english_name,
)
from intent import extract_intent, YuntuTask, InsightTimeRange, BrandCategory
from login import check_profile_exists, perform_manual_login
from data_extractor import extract_from_agent_history
from excel_processor import (
    process_downloaded_excel_files,
    calculate_metrics,
    sum_search_count_from_excel,
    sum_search_count_from_downloads,
)
from excel_exporter import export_to_excel
from utils import find_downloaded_excel_files
from data_models import YuntuDataReport
from datetime import datetime

# Global state to store extracted intent across conversation turns
# In production, you might want to use Gradio's State component
_session_intent: YuntuTask | None = None

# Global state to store current running agent for stop functionality
_current_agent: Agent | None = None
_agent_task: asyncio.Task | None = None

# 当前任务的工作目录（下载+导出统一放此目录，供 Tool 使用）
_current_task_dir: Path | None = None

# Custom tools for Agent (e.g. 回搜次数 from Excel)
_yuntu_tools = Tools()


@_yuntu_tools.action(
    "用于回搜次数计算，从下载的 Excel 文件中按计算时间区间累加「搜索次数」并返回合计值。"
    "Excel 需包含列：日期(YYYYMMDD)、搜索人数、搜索次数。"
)
async def sum_search_count_in_date_range(
    date_start: str,
    date_end: str,
    file_path: str | None = None
) -> ActionResult:
    """
    按 date_start 到 date_end 的区间，从 Excel 累加「搜索次数」。
    date_start/date_end 支持 "2025-12-01" 或 "20251201"。
    若不传 file_path，则从配置的下载目录中查找最近下载的符合条件的 Excel。
    """
    try:
        downloads_path = _current_task_dir if _current_task_dir is not None else DOWNLOADS_PATH
        if file_path:
            total = sum_search_count_from_excel(file_path, date_start, date_end)
        else:
            total = sum_search_count_from_downloads(
                downloads_path, date_start, date_end
            )
        return ActionResult(
            extracted_content=f"区间内搜索次数合计: {total}",
            success=True,
        )
    except FileNotFoundError as e:
        return ActionResult(error=str(e))
    except ValueError as e:
        return ActionResult(error=str(e))


async def process_task_generator(history: list):  # type: ignore
    """
    Async generator that processes user tasks and yields chat history updates.
    
    Args:
        history: List of message dictionaries with 'role' and 'content' keys
        
    Yields:
        Updated history list after each processing step
    """
    global _current_agent, _agent_task, _session_intent, _current_task_dir
    
    # Stop any currently running agent before starting a new task
    if _current_agent and not _current_agent.state.stopped:
        _current_agent.stop()
        if _agent_task and not _agent_task.done():
            _agent_task.cancel()
            try:
                await _agent_task
            except (asyncio.CancelledError, Exception):
                pass
    
    if not history:
        return
    
    # Extract user message from history
    last_msg = history[-1]
    
    if isinstance(last_msg, dict) and last_msg.get("role") == "user":
        user_message = last_msg.get("content", "")
    else:
        return
    
    if not user_message:
        return
    
    # Add assistant response placeholder
    history.append({"role": "assistant", "content": "..."})
    yield history.copy()
    
    # Step 1: Intent Recognition with history context
    history[-1]["content"] = "🔍 正在分析您的需求（包含历史上下文）..."
    yield history.copy()
    
    # Extract conversation history (all previous messages)
    conversation_history = [
        {"role": msg.get("role", "unknown"), "content": msg.get("content", "")}
        for msg in history[:-1]  # Exclude the current assistant placeholder
        if isinstance(msg, dict) and msg.get("role") in ("user", "assistant")
    ]
    
    # Use global session intent to preserve previously extracted values
    previous_intent = _session_intent
    
    # Extract intent with history context and previous values
    task_intent: YuntuTask = await extract_intent(
        user_message,
        conversation_history=conversation_history,
        previous_intent=previous_intent
    )
    
    # Update session intent with merged values
    _session_intent = task_intent
    
    # Step 2: Validation - 报告名或品牌分类（二选一）+ 消耗金额必填
    categories = task_intent.brand_categories or []
    has_single = bool(task_intent.report_name)
    has_categories = bool(categories) and all(
        getattr(c, "report_name_star", None)
        and getattr(c, "report_name_bid", None)
        and getattr(c, "consumption_amount_star", None) is not None
        and getattr(c, "consumption_amount_bid", None) is not None
        for c in categories
    )
    if not has_single and not has_categories:
        history[-1]["content"] = (
            "❌ **无法识别报告信息或消耗金额**\n\n"
            "请任选一种方式提供：\n"
            "1. 单个报告名 + 消耗金额：报告名：******；星图 10 万、竞价 20 万\n"
            "2. 品牌分类（每分类含星图报告名+消耗金额、竞价报告名+消耗金额）：\n"
            "   品牌分类：分类A 星图报告xxx 10万 竞价报告yyy 20万；分类B 星图报告xxx 消耗金额 竞价报告yyy 消耗金额"
        )
        yield history.copy()
        return

    if has_single and (
        task_intent.consumption_amount_star is None or task_intent.consumption_amount_bid is None
    ):
        missing = []
        if task_intent.consumption_amount_star is None:
            missing.append("星图（达人营销）消耗金额")
        if task_intent.consumption_amount_bid is None:
            missing.append("竞价（竞价投放）消耗金额")
        history[-1]["content"] = (
            "❌ **缺少消耗金额（必填）**\n\n"
            f"请提供：{'、'.join(missing)}。\n\n"
            "示例：星图 10 万、竞价 20 万；或「星图消耗 100000 元，竞价消耗 200000 元」"
        )
        yield history.copy()
        return

    # Step 2b: Validation - 爆文加热时间（第五步）与 行业搜索洞察时间（第六步）
    time_ranges = task_intent.insight_time_ranges or []
    if task_intent.brand_name and "黑玩" in (task_intent.brand_name or ""):
        if not task_intent.kol_content_date_type:
            history[-1]["content"] = (
                "❌ **缺少爆文加热时间（第五步 达人及内容复盘）**\n\n"
                "请提供时间类型，例如：\n"
                "- 爆文加热：近30天\n"
                "- 爆文加热：按月（2025/12/01-2025/12/31）\n\n"
                "支持的类型：实时、按周、按月、近7天、近30天(实时/近7天/近30天可不提供具体时间范围)"
            )
            yield history.copy()
            return
    if not time_ranges:
        history[-1]["content"] = (
            "❌ **缺少行业搜索洞察的时间范围信息（第六步）**\n\n"
            "请提供至少一个时间范围，例如：\n"
            "- 行业搜索洞察：自定义（2025/12/01-2025/12/31）\n"
            "- 行业搜索洞察：近7天、近30天\n\n"
            "支持的类型：按周、按月、近7天、近30天、自定义"
        )
        yield history.copy()
        return

    # Step 3: Profile Check & Login
    history[-1]["content"] = "🔐 检查登录状态..."
    yield history.copy()
    
    if not await check_profile_exists():
        history[-1]["content"] = (
            "⚠️ **未检测到登录信息**\n\n"
            "正在打开浏览器，请您在浏览器中手动登录巨量云图...\n"
            "登录成功后，系统将自动检测并继续执行任务。"
        )
        yield history.copy()
        
        await perform_manual_login()
        
        history[-1]["content"] = "✅ 登录流程已完成，准备开始执行任务..."
        yield history.copy()
    
    # Step 4: Prompt Assembly
    if task_intent.brand_name and "黑玩" in task_intent.brand_name:
        template = HEYONE_PROMPT_TEMPLATE
        brand_info = "（使用黑玩专属模板）"
    else:
        template = GENERIC_PROMPT_TEMPLATE
        brand_info = ""
        
    # 爆文加热（第五步）：单值；行业搜索洞察（第六步）：多时间范围
    insight_time_ranges_formatted = "\n".join(
        f"{i}. {tr.date_range_type}" + (f"（{tr.insight_date_range}）" if tr.insight_date_range else "")
        for i, tr in enumerate(time_ranges, 1)
    )
    kol_type = task_intent.kol_content_date_type or "N/A"
    kol_range = task_intent.kol_content_date_range or "（无具体日期）"
    # 品牌分类（黑玩）：多分类时每分类含星图报告+消耗金额、竞价报告+消耗金额
    brand_categories_formatted = ""
    report_name_for_prompt = task_intent.report_name or ""
    if categories:
        lines = []
        for i, c in enumerate(categories, 1):
            name = getattr(c, "category_name", None) or f"分类{i}"
            star = getattr(c, "report_name_star", None) or ""
            star_amt = getattr(c, "consumption_amount_star", None)
            bid = getattr(c, "report_name_bid", None) or ""
            bid_amt = getattr(c, "consumption_amount_bid", None)
            star_amt_str = f" {star_amt:,.0f}元" if star_amt is not None else ""
            bid_amt_str = f" {bid_amt:,.0f}元" if bid_amt is not None else ""
            lines.append(f"{i}. {name}：星图报告={star}{star_amt_str}，竞价报告={bid}{bid_amt_str}")
        brand_categories_formatted = "\n".join(lines)
        report_name_for_prompt = "见下方品牌分类列表"
    # Use replace() so template can contain literal { } in JSON examples without escaping
    first_insight = time_ranges[0] if time_ranges else None
    brand_display = task_intent.brand_name or "未指定"
    brand_name_en = get_brand_english_name(task_intent.brand_name)
    final_prompt = (
        template.replace("{report_name}", report_name_for_prompt)
        .replace("{brand_name}", brand_display)
        .replace("{brand_name_en}", brand_name_en)
        .replace("{brand_categories_formatted}", brand_categories_formatted)
        .replace("{kol_content_date_type}", kol_type)
        .replace("{kol_content_date_range}", kol_range)
        .replace("{insight_time_ranges_formatted}", insight_time_ranges_formatted)
        .replace("{insight_date_range}", first_insight.insight_date_range or "N/A" if first_insight else "N/A")
        .replace("{date_range_type}", first_insight.date_range_type or "N/A" if first_insight else "N/A")
    )
    
    # Display task summary（消耗金额：品牌分类时按分类展示，单报告时按星图/竞价展示）
    def _fmt_amt(val: float | None) -> str:
        return f"{val:,.0f} 元" if val is not None else "未提供"
    if categories:
        consumption_lines = (
            "- **消耗金额（按分类）**:\n"
            + "\n".join(
                f"  - {getattr(c, 'category_name', None) or f'分类{i}'}: 星图 {_fmt_amt(getattr(c, 'consumption_amount_star', None))}，竞价 {_fmt_amt(getattr(c, 'consumption_amount_bid', None))}"
                for i, c in enumerate(categories, 1)
            )
            + "\n"
        )
    else:
        consumption_lines = (
            f"- **星图（达人营销）消耗金额**: {_fmt_amt(task_intent.consumption_amount_star)}\n"
            f"- **竞价（竞价投放）消耗金额**: {_fmt_amt(task_intent.consumption_amount_bid)}\n"
        )
    kol_line = f"- **爆文加热时间（第五步）**: {task_intent.kol_content_date_type or '未指定'} {task_intent.kol_content_date_range or ''}\n"
    tr_lines = "\n".join(
        f"  - {i}. {tr.date_range_type}" + (f"（{tr.insight_date_range}）" if tr.insight_date_range else "")
        for i, tr in enumerate(time_ranges, 1)
    )
    time_ranges_block = f"- **行业搜索洞察-时间范围（第六步）**:\n{tr_lines}\n"
    report_block = (
        f"- **品牌分类（星图+竞价报告+消耗金额）**:\n"
        + "\n".join(
            f"  - {getattr(c, 'category_name', None) or f'分类{i}'}: 星图={getattr(c, 'report_name_star', '')} {_fmt_amt(getattr(c, 'consumption_amount_star', None))}，竞价={getattr(c, 'report_name_bid', '')} {_fmt_amt(getattr(c, 'consumption_amount_bid', None))}"
            for i, c in enumerate(categories, 1)
        )
        + "\n"
        if categories
        else f"- **报告名**: {task_intent.report_name or '未指定'}\n"
    )
    task_summary = (
        f"📋 **任务摘要** {brand_info}\n\n"
        f"{report_block}"
        f"- **品牌**: {task_intent.brand_name or '未指定'}\n"
        f"{kol_line}"
        f"{time_ranges_block}"
        f"{consumption_lines}\n"
        f"🚀 正在启动 Agent 执行任务..."
    )
    history[-1]["content"] = task_summary
    yield history.copy()
    
    # Step 5: Execution - 本任务使用独立目录（下载+导出+对话均写入，任务间隔离）
    task_timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    task_dir = TASKS_BASE_DIR / f"task_{task_timestamp}"
    task_dir.mkdir(parents=True, exist_ok=True)
    _current_task_dir = task_dir

    browser = None
    try:
        browser = Browser(
            user_data_dir=str(PROFILE_DIR),
            storage_state=str(STORAGE_STATE_FILE) if STORAGE_STATE_FILE.exists() else None,
            downloads_path=str(task_dir),
            headless=False  # Set to False for debugging, True for production
        )
        (task_dir / "conversations").mkdir(parents=True, exist_ok=True)
        agent = Agent(
            task=final_prompt,
            llm=ChatGoogle(model=GEMINI_MODEL, api_key=API_KEY),
            browser=browser,
            tools=_yuntu_tools,
            flash_mode=False,
            file_system_path=str(task_dir),
            save_conversation_path=str(task_dir / "conversations")
        )
        
        # Store agent globally for stop functionality
        _current_agent = agent
        
        history[-1]["content"] = task_summary + "\n\n⏳ Agent 正在执行中...\n\n---\n**实时日志：**\n等待中..."
        yield history.copy()
        
        # Collect logs for real-time display
        log_messages = []
        last_log_count = 0
        
        async def on_step_start(agent_instance):
            """Callback when a step starts"""
            step_num = agent_instance.state.n_steps
            log_messages.append(f"📌 **步骤 {step_num} 开始**")
        
        async def on_step_end(agent_instance):
            """Callback when a step ends"""
            step_num = agent_instance.state.n_steps
            model_output = agent_instance.state.last_model_output
            
            if model_output:
                # Extract action information
                actions = model_output.action
                for action in actions:
                    action_name = action.action_name if hasattr(action, 'action_name') else str(type(action).__name__)
                    log_messages.append(f"  ✓ 执行动作: {action_name}")
                
                # Extract thinking/evaluation if available
                if model_output.thinking:
                    thinking_short = model_output.thinking[:100] + "..." if len(model_output.thinking) > 100 else model_output.thinking
                    log_messages.append(f"  💡 思考: {thinking_short}")
                
                if model_output.next_goal:
                    goal_short = model_output.next_goal[:80] + "..." if len(model_output.next_goal) > 80 else model_output.next_goal
                    log_messages.append(f"  🎯 下一步: {goal_short}")
            
            # Get action results
            results = agent_instance.state.last_result
            if results:
                for result in results:
                    if result.extracted_content:
                        content_short = result.extracted_content[:150] + "..." if len(result.extracted_content) > 150 else result.extracted_content
                        log_messages.append(f"  📄 提取: {content_short}")
            
            log_messages.append(f"✅ **步骤 {step_num - 1} 完成**\n")
        
        # Run agent with callbacks in background
        async def run_agent():
            try:
                return await agent.run(
                    on_step_start=on_step_start,
                    on_step_end=on_step_end
                )
            except asyncio.CancelledError:
                log_messages.append("⚠️ **Agent 执行被用户取消**")
                raise
            except Exception as e:
                log_messages.append(f"❌ **Agent 执行出错**: {str(e)}")
                raise
        
        agent_task = asyncio.create_task(run_agent())
        _agent_task = agent_task
        
        # Periodically update UI with new logs
        agent_stopped = False
        while not agent_task.done():
            # Check if agent was stopped
            if agent.state.stopped:
                agent_stopped = True
                log_messages.append("⏹️ **Agent 已停止**")
                if not agent_task.done():
                    agent_task.cancel()
                    try:
                        await agent_task
                    except (asyncio.CancelledError, Exception):
                        pass
                break
            
            if len(log_messages) > last_log_count:
                # New logs available, update UI
                log_text = "\n".join(log_messages[-40:])  # Keep last 40 log entries
                history[-1]["content"] = task_summary + f"\n\n⏳ Agent 正在执行中...\n\n---\n**实时日志：**\n{log_text}"
                yield history.copy()
                last_log_count = len(log_messages)
            
            # Check if agent is done
            await asyncio.sleep(0.5)  # Check every 0.5 seconds
        
        # Get result if not stopped
        if not agent_stopped:
            try:
                result = await agent_task
            except asyncio.CancelledError:
                result = None
                agent_stopped = True
        else:
            result = None
        
        # Final update with all logs
        log_text = "\n".join(log_messages[-50:])  # Show last 50 entries
        
        if agent_stopped:
            _current_task_dir = None
            history[-1]["content"] = task_summary + f"\n\n⏹️ **Agent 执行已停止**\n\n---\n**执行日志：**\n{log_text}\n\n---\n您可以输入新的任务重新开始。"
        else:
            history[-1]["content"] = task_summary + f"\n\n✅ Agent 执行完成\n\n---\n**执行日志：**\n{log_text}\n\n---\n"
            yield history.copy()
            
            if result:
                # 将 Agent 的 JSON 类输出保存到任务目录
                output_dir = _current_task_dir
                if output_dir:
                    import json as _json
                    agent_output_path = output_dir / "agent_output.json"
                    try:
                        payload = {
                            "final_result": result.final_result(),
                            "extracted_content": result.extracted_content(),
                        }
                        agent_output_path.write_text(
                            _json.dumps(payload, ensure_ascii=False, indent=2),
                            encoding="utf-8",
                        )
                    except Exception as _e:
                        pass  # 不阻塞主流程
                    # 将 browseruse_agent_data 下的 json 复制到任务根目录便于查看
                    agent_data_dir = output_dir / "browseruse_agent_data"
                    if agent_data_dir.exists():
                        for _f in agent_data_dir.glob("*.json"):
                            try:
                                (output_dir / _f.name).write_text(_f.read_text(encoding="utf-8"), encoding="utf-8")
                            except Exception:
                                pass

                # 数据提取和Excel导出
                history[-1]["content"] = task_summary + f"\n\n✅ Agent 执行完成\n\n---\n**执行日志：**\n{log_text}\n\n---\n📊 正在提取数据并生成Excel报告..."
                yield history.copy()
                
                try:
                    _trs = _session_intent.insight_time_ranges if _session_intent else None
                    _dr = _trs[0].insight_date_range if _trs and len(_trs) > 0 else None
                    _cats = _session_intent.brand_categories if _session_intent else None
                    _report_name = (
                        _session_intent.report_name if _session_intent else None
                    ) or (
                        _cats[0].report_name_star if _cats and len(_cats) > 0 else None
                    )
                    report_names_ordered: list[str] | None = None
                    if _cats:
                        report_names_ordered = []
                        for c in _cats:
                            if getattr(c, "report_name_star", None):
                                report_names_ordered.append(c.report_name_star)
                            if getattr(c, "report_name_bid", None):
                                report_names_ordered.append(c.report_name_bid)
                    extract_result = extract_from_agent_history(
                        result,
                        report_name=_report_name,
                        brand_name=_session_intent.brand_name if _session_intent else None,
                        date_range=_dr,
                        report_names=report_names_ordered if report_names_ordered else None,
                    )
                    reports_to_export: list[YuntuDataReport] = (
                        [extract_result] if isinstance(extract_result, YuntuDataReport) else extract_result
                    )
                    if reports_to_export:
                        first_report = reports_to_export[0]
                        kol_content, ad_flow = process_downloaded_excel_files(str(_current_task_dir), first_report)
                        if kol_content:
                            first_report.达人及内容复盘 = kol_content
                        if ad_flow:
                            first_report.投流数据精细化复盘 = ad_flow
                        for report in reports_to_export:
                            if report.项目整体:
                                calculated_metrics = calculate_metrics(report.项目整体)
                                if report.项目整体.CPM is None and calculated_metrics.get('CPM'):
                                    report.项目整体.CPM = calculated_metrics['CPM']
                                if report.项目整体.CPE is None and calculated_metrics.get('CPE'):
                                    report.项目整体.CPE = calculated_metrics['CPE']
                                if report.项目整体.CPS is None and calculated_metrics.get('CPS'):
                                    report.项目整体.CPS = calculated_metrics['CPS']
                                if report.项目整体.CPA3 is None and calculated_metrics.get('CPA3'):
                                    report.项目整体.CPA3 = calculated_metrics['CPA3']
                                if report.项目整体.ROI is None and calculated_metrics.get('ROI'):
                                    report.项目整体.ROI = calculated_metrics['ROI']
                        excel_path = export_to_excel(
                            reports_to_export if len(reports_to_export) > 1 else reports_to_export[0],
                            _current_task_dir,
                            filename=None,
                        )
                        _current_task_dir = None
                        # 5. 更新UI显示结果
                        final_output = result.final_result()
                        excel_info = f"\n\n📊 **Excel报告已生成**\n\n文件路径：`{excel_path}`\n\n"
                        
                        if final_output:
                            history[-1]["content"] = task_summary + f"\n\n✅ **任务完成**\n\n{final_output}{excel_info}"
                        else:
                            extracted = result.extracted_content()
                            if extracted:
                                history[-1]["content"] = task_summary + f"\n\n✅ **任务完成**\n\n提取到的内容：\n{chr(10).join(str(e) for e in extracted if e)}{excel_info}"
                            else:
                                history[-1]["content"] = task_summary + f"\n\n✅ **任务执行完成**\n\nExcel报告已生成，但未提取到明确的文本结果。请检查Excel文件。{excel_info}"
                    else:
                        history[-1]["content"] = task_summary + f"\n\n⚠️ **数据提取失败**\n\nAgent执行完成，但无法提取结构化数据。请检查Agent输出格式。"
                    
                except Exception as e:
                    # Excel导出失败不影响主流程
                    import traceback
                    error_msg = f"\n\n⚠️ **Excel导出失败**: {str(e)}\n{traceback.format_exc()}"
                    final_output = result.final_result()
                    if final_output:
                        history[-1]["content"] = f"✅ **任务完成**\n\n{final_output}{error_msg}"
                    else:
                        history[-1]["content"] = f"✅ 任务执行完成，但Excel导出失败。{error_msg}"
        
        yield history.copy()
        
    except Exception as e:
        history[-1]["content"] = f"❌ **任务执行出错**\n\n```\n{str(e)}\n```"
        yield history.copy()
    finally:
        _current_task_dir = None
        _current_agent = None
        _agent_task = None
        if browser:
            try:
                await browser.stop()
            except Exception:
                pass


def main():
    """Create and launch the Gradio interface."""
    with gr.Blocks(
        title="巨量云图数据助手"
    ) as demo:
        gr.Markdown(
            """
            # 🌐 巨量云图数据获取助手
            
            通过自然语言描述您的数据获取需求，系统将自动执行并返回结果。
            
            **示例输入**

            *单报告*
            > 从巨量云图获取黑玩复盘数据，报告名：ZY-Heyone12.1-12.31；星图 10万、竞价 20万；爆文加热：近30天；行业搜索洞察：自定义（2025/12/01-2025/12/31）

            *品牌分类（多报告）*
            > 从巨量云图获取黑玩复盘数据，品牌分类：洗护 星图报告xxx 10万 竞价报告yyy 20万；美发 星图报告xxx 15万 竞价报告yyy 25万；爆文加热：近30天；行业搜索洞察：自定义（2025/12/01-2025/12/31）
            """
        )
        
        chatbot = gr.Chatbot(
            height=500,
            show_label=False,
            avatar_images=(None, "🤖"),
            type="messages",
            allow_tags=False,
        )
        
        with gr.Row():
            msg = gr.Textbox(
                label="输入任务",
                placeholder="请描述您的数据获取需求...",
                scale=9,
                show_label=False
            )
            submit_btn = gr.Button("发送", scale=1, variant="primary", visible=True)
            stop_btn = gr.Button("停止执行", variant="stop", visible=False)
        
        with gr.Row():
            clear_btn = gr.Button("清空对话")
            
        # Event handlers
        def stop_agent(history: list) -> tuple:
            """Stop the currently running agent"""
            global _current_agent, _agent_task
            if _current_agent:
                try:
                    _current_agent.stop()
                    # Add stop message to history
                    if history:
                        history.append({
                            "role": "assistant",
                            "content": "⏹️ **Agent 执行已停止**\n\n您可以输入新的任务重新开始数据抓取。"
                        })
                    return history, gr.Button("发送", variant="primary", visible=True), gr.Button("停止执行", variant="stop", visible=False)
                except Exception as e:
                    if history:
                        history.append({
                            "role": "assistant",
                            "content": f"❌ **停止失败**: {str(e)}"
                        })
                    return history, gr.Button("发送", variant="primary", visible=True), gr.Button("停止执行", variant="stop", visible=False)
            return history, gr.Button("发送", variant="primary", visible=True), gr.Button("停止执行", variant="stop", visible=False)
        
        stop_btn.click(
            stop_agent,
            [chatbot],
            [chatbot, submit_btn, stop_btn],
            queue=False
        )
        
        def user_msg(user_message: str, history: list) -> tuple:
            """Add user message to history and clear input."""
            if not user_message.strip():
                return "", history or [], gr.Button("发送", variant="primary", visible=True), gr.Button("停止执行", variant="stop", visible=False)
            # Gradio 6.x format: list of dicts with "role" and "content"
            new_history = list(history) if history else []
            new_history.append({"role": "user", "content": user_message})
            return "", new_history, gr.Button("发送", variant="primary", visible=False), gr.Button("停止执行", variant="stop", visible=True)

        def reset_buttons():
            """Reset buttons to initial state"""
            return gr.Button("发送", variant="primary", visible=True), gr.Button("停止执行", variant="stop", visible=False)
        
        # Submit on enter or button click
        msg.submit(
            user_msg, 
            [msg, chatbot], 
            [msg, chatbot, submit_btn, stop_btn], 
            queue=False
        ).then(
            process_task_generator, 
            [chatbot], 
            [chatbot]
        ).then(
            reset_buttons,
            None,
            [submit_btn, stop_btn],
            queue=False
        )
        
        submit_btn.click(
            user_msg, 
            [msg, chatbot], 
            [msg, chatbot, submit_btn, stop_btn], 
            queue=False
        ).then(
            process_task_generator, 
            [chatbot], 
            [chatbot]
        ).then(
            reset_buttons,
            None,
            [submit_btn, stop_btn],
            queue=False
        )
        
        def clear_chat():
            """Clear chat and reset session intent and agent."""
            global _session_intent, _current_agent, _agent_task
            # Stop any running agent
            if _current_agent and not _current_agent.state.stopped:
                try:
                    _current_agent.stop()
                except Exception:
                    pass
            _session_intent = None
            _current_agent = None
            _agent_task = None
            return [], gr.Button("发送", variant="primary", visible=True), gr.Button("停止执行", variant="stop", visible=False)
        
        clear_btn.click(clear_chat, None, [chatbot, submit_btn, stop_btn], queue=False)

    demo.launch(server_name="127.0.0.1", server_port=7860)


if __name__ == "__main__":
    main()
