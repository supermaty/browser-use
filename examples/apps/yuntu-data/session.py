"""
Session management for Yuntu Data Scraper.
Encapsulates the state of a task execution session.
"""
import asyncio
from pathlib import Path
from browser_use import Agent
from intent import YuntuTask

class TaskSession:
    """
    Manages the state of a data scraping task session.
    
    Attributes:
        intent (YuntuTask | None): The current task intent.
        agent (Agent | None): The currently running browser agent.
        agent_task (asyncio.Task | None): The asyncio task for the running agent.
        task_dir (Path | None): The directory for the current task's output.
    """
    def __init__(self):
        self.intent: YuntuTask | None = None
        self.agent: Agent | None = None
        self.agent_task: asyncio.Task | None = None
        self.task_dir: Path | None = None

    async def stop_current_agent(self):
        """Stops the currently running agent if it exists."""
        if self.agent and not self.agent.state.stopped:
            try:
                self.agent.stop()
            except Exception:
                pass
        
        if self.agent_task and not self.agent_task.done():
            self.agent_task.cancel()
            try:
                await self.agent_task
            except (asyncio.CancelledError, Exception):
                pass
        
        self.agent = None
        self.agent_task = None

    def clear(self):
        """Resets the session state."""
        self.intent = None
        self.agent = None
        self.agent_task = None
        self.task_dir = None
