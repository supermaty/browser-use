"""
Pipeline orchestration for Yuntu Data Scraper.
Manages the end-to-end flow of a task: intent -> validation -> login -> execution -> post-processing.
"""
import asyncio
import json
from datetime import datetime
from pathlib import Path

from browser_use import Agent, Browser, Tools
from browser_use.llm.google import ChatGoogle

from config import (
    PROFILE_DIR,
    GEMINI_MODEL,
    API_KEY,
    STORAGE_STATE_FILE,
    TASKS_BASE_DIR,
)
from intent import extract_intent, YuntuTask
from login import check_profile_exists, perform_manual_login
from prompt_builder import build_prompt, build_task_summary
from validator import validate_intent
from session import TaskSession
from data_extractor import extract_from_agent_history
from json_exporter import export_agent_output_to_json, export_structured_reports_to_json
from data_models import YuntuDataReport


class TaskPipeline:
    def __init__(self, session: TaskSession, tools: Tools):
        self.session = session
        self.tools = tools

    async def process(self, history: list):
        """
        Executes the task pipeline and yields history updates.
        """
        # Stop any currently running agent before starting a new task
        await self.session.stop_current_agent()
        
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
        previous_intent = self.session.intent
        
        # Extract intent with history context and previous values
        task_intent: YuntuTask = await extract_intent(
            user_message,
            conversation_history=conversation_history,
            previous_intent=previous_intent
        )
        
        # Update session intent with merged values
        self.session.intent = task_intent
        
        # Step 2: Validation
        validation = validate_intent(task_intent)
        if not validation.is_valid:
            history[-1]["content"] = validation.message
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
        final_prompt = build_prompt(task_intent)
        task_summary = build_task_summary(task_intent)
        
        history[-1]["content"] = task_summary
        yield history.copy()
        
        # Step 5: Execution
        task_timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        task_dir = TASKS_BASE_DIR / f"task_{task_timestamp}"
        task_dir.mkdir(parents=True, exist_ok=True)
        self.session.task_dir = task_dir

        browser = None
        try:
            browser = Browser(
                user_data_dir=str(PROFILE_DIR),
                storage_state=str(STORAGE_STATE_FILE) if STORAGE_STATE_FILE.exists() else None,
                downloads_path=str(task_dir),
                headless=False
            )
            (task_dir / "conversations").mkdir(parents=True, exist_ok=True)
            agent = Agent(
                task=final_prompt,
                llm=ChatGoogle(model=GEMINI_MODEL, api_key=API_KEY),
                browser=browser,
                tools=self.tools,
                flash_mode=False,
                file_system_path=str(task_dir),
                save_conversation_path=str(task_dir / "conversations")
            )
            
            # Store agent globally for stop functionality
            self.session.agent = agent
            
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
            self.session.agent_task = agent_task
            
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
                self.session.task_dir = None
                history[-1]["content"] = task_summary + f"\n\n⏹️ **Agent 执行已停止**\n\n---\n**执行日志：**\n{log_text}\n\n---\n您可以输入新的任务重新开始。"
            else:
                history[-1]["content"] = task_summary + f"\n\n✅ Agent 执行完成\n\n---\n**执行日志：**\n{log_text}\n\n---\n"
                yield history.copy()
                
                if result:
                    await self._handle_post_processing(result, history, task_summary, log_text)
            
            yield history.copy()
            
        except Exception as e:
            history[-1]["content"] = f"❌ **任务执行出错**\n\n```\n{str(e)}\n```"
            yield history.copy()
        finally:
            self.session.clear()
            if browser:
                try:
                    await browser.stop()
                except Exception:
                    pass

    async def _handle_post_processing(self, result, history, task_summary, log_text):
        """Handles post-processing: saving JSON outputs and extracting data."""
        output_dir = self.session.task_dir
        if output_dir:
            agent_output_path = output_dir / "agent_output.json"
            try:
                payload = {
                    "final_result": result.final_result(),
                    "extracted_content": result.extracted_content(),
                }
                agent_output_path.write_text(
                    json.dumps(payload, ensure_ascii=False, indent=2),
                    encoding="utf-8",
                )
            except Exception:
                pass
            
            # Copy browseruse_agent_data jsons
            agent_data_dir = output_dir / "browseruse_agent_data"
            if agent_data_dir.exists():
                for _f in agent_data_dir.glob("*.json"):
                    try:
                        (output_dir / _f.name).write_text(_f.read_text(encoding="utf-8"), encoding="utf-8")
                    except Exception:
                        pass

        # Update UI for export step
        history[-1]["content"] = task_summary + f"\n\n✅ Agent 执行完成\n\n---\n**执行日志：**\n{log_text}\n\n---\n📊 正在导出 JSON 数据..."
        # Note: We can't yield here easily because this is a subroutine. 
        # But since we update history content in place and the main loop yields after this returns (or we could pass yield_callback),
        # but here we are just updating the object reference. The caller yields `history.copy()` right after this returns.
        
        try:
            # 1. Save Agent raw JSON output
            _report_name = None
            _brand_name = self.session.intent.brand_name if self.session.intent else None
            _cats = self.session.intent.brand_categories if self.session.intent else None
            if self.session.intent:
                _report_name = self.session.intent.report_name or (
                    _cats[0].report_name_star if _cats and len(_cats) > 0 else None
                )
            
            json_path = export_agent_output_to_json(
                agent_final_result=result.final_result(),
                agent_extracted_content=result.extracted_content(),
                output_dir=self.session.task_dir,
                report_name=_report_name,
                brand_name=_brand_name,
            )

            # 2. Extract structured data
            _trs = self.session.intent.insight_time_ranges if self.session.intent else None
            _dr = _trs[0].insight_date_range if _trs and len(_trs) > 0 else None
            report_names_ordered: list[str] | None = None
            if _cats:
                report_names_ordered = []
                for c in _cats:
                    _star = getattr(c, "report_name_star", None)
                    if _star:
                        report_names_ordered.append(_star)
                    _bid = getattr(c, "report_name_bid", None)
                    if _bid:
                        report_names_ordered.append(_bid)
            
            extract_result = extract_from_agent_history(
                result,
                report_name=_report_name,
                brand_name=_brand_name,
                date_range=_dr,
                report_names=report_names_ordered if report_names_ordered else None,
            )
            
            reports_list: list[YuntuDataReport] = (
                [extract_result] if isinstance(extract_result, YuntuDataReport) else extract_result
            )
            
            structured_path = export_structured_reports_to_json(reports_list, self.session.task_dir)

            self.session.task_dir = None
            
            # 3. Final UI Update
            final_output = result.final_result()
            json_info = (
                f"\n\n📊 **JSON 数据已导出**\n\n"
                f"- Agent 原始输出: `{json_path}`\n"
                f"- 结构化数据: `{structured_path}`\n\n"
            )
            
            if final_output:
                history[-1]["content"] = task_summary + f"\n\n✅ **任务完成**\n\n{final_output}{json_info}"
            else:
                extracted = result.extracted_content()
                if extracted:
                    history[-1]["content"] = task_summary + f"\n\n✅ **任务完成**\n\n提取到的内容：\n{chr(10).join(str(e) for e in extracted if e)}{json_info}"
                else:
                    history[-1]["content"] = task_summary + f"\n\n✅ **任务执行完成**\n\nJSON 已导出，但未提取到明确的文本结果。请检查 JSON 文件。{json_info}"

        except Exception as e:
            import traceback
            error_msg = f"\n\n⚠️ **JSON 导出失败**: {str(e)}\n{traceback.format_exc()}"
            final_output = result.final_result()
            if final_output:
                history[-1]["content"] = f"✅ **任务完成**\n\n{final_output}{error_msg}"
            else:
                history[-1]["content"] = f"✅ 任务执行完成，但 JSON 导出失败。{error_msg}"
