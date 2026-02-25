# pyright: reportMissingImports=false
"""
Yuntu Data Scraper - Gradio Application

A chat-based UI for extracting data from Yuntu (巨量云图) using Browser Use.
Thin UI layer that delegates pipeline orchestration to pipeline.py.
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
from browser_use import ActionResult, Tools

# Add app directory to path for local imports
APP_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(APP_DIR))

from config import DOWNLOADS_PATH
from excel_processor import (
    sum_search_count_from_excel,
    sum_search_count_from_downloads,
)
from session import TaskSession
from pipeline import TaskPipeline

# Global session state
_session = TaskSession()

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
        downloads_path = _session.task_dir if _session.task_dir is not None else DOWNLOADS_PATH
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
    Delegates to TaskPipeline for all orchestration logic.
    """
    pipeline = TaskPipeline(_session, _yuntu_tools)
    async for updated_history in pipeline.process(history):
        yield updated_history


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

            *单品牌分类报告*
            > 从巨量云图获取黑玩复盘数据，报告名：ZY-Heyone12.1-12.31；星图 10万、竞价 20万；爆文加热：近30天；行业搜索洞察：自定义（2025/12/01-2025/12/31）

            *多品牌分类报告*
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
            if _session.agent:
                try:
                    _session.agent.stop()
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
            if _session.agent and not _session.agent.state.stopped:
                try:
                    _session.agent.stop()
                except Exception:
                    pass
            _session.clear()
            return [], gr.Button("发送", variant="primary", visible=True), gr.Button("停止执行", variant="stop", visible=False)
        
        clear_btn.click(clear_chat, None, [chatbot, submit_btn, stop_btn], queue=False)

    demo.launch(server_name="127.0.0.1", server_port=7860)


if __name__ == "__main__":
    main()
