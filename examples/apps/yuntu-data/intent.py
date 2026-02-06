import json
from typing import Optional, Literal

from pydantic import BaseModel, Field
from browser_use.llm.google import ChatGoogle
from browser_use.llm.messages import SystemMessage, UserMessage

from config import INTENT_EXTRACTION_SYSTEM_PROMPT, GEMINI_MODEL

# 爆文加热（第五步 达人及内容复盘）：单值，时间类型 + 时间范围
KOL_CONTENT_DATE_TYPE = Literal["实时", "按周", "按月", "近7天", "近30天"]
# 行业搜索洞察（第六步）：一到多值，时间类型 + 时间范围
INSIGHT_DATE_TYPE = Literal["按周", "按月", "近7天", "近30天", "自定义"]


class InsightTimeRange(BaseModel):
    """行业搜索洞察的单个时间范围"""
    date_range_type: INSIGHT_DATE_TYPE
    insight_date_range: Optional[str] = Field(None, description="自定义/按周/按月时的具体日期范围字符串")


class BrandCategory(BaseModel):
    """黑玩品牌案例：单个品牌分类，含星图/竞价报告名及各自消耗金额"""
    category_name: Optional[str] = Field(None, description="品牌分类名称（可选），如「洗护」「美发」")
    report_name_star: Optional[str] = Field(None, description="该分类的星图（达人营销）报告名")
    consumption_amount_star: Optional[float] = Field(None, description="该分类星图报告消耗金额（元），必填")
    report_name_bid: Optional[str] = Field(None, description="该分类的竞价（竞价投放）报告名")
    consumption_amount_bid: Optional[float] = Field(None, description="该分类竞价报告消耗金额（元），必填")


class YuntuTask(BaseModel):
    report_name: Optional[str] = Field(None, description="单个报告名（非黑玩多分类时使用）")
    brand_categories: Optional[list[BrandCategory]] = Field(
        None,
        description="黑玩品牌案例：多个品牌分类，每分类含星图报告名+消耗金额、竞价报告名+消耗金额；与 report_name 二选一",
    )
    # 爆文加热时间（第五步 达人及内容复盘）：单值
    kol_content_date_type: Optional[KOL_CONTENT_DATE_TYPE] = Field(
        None,
        description="爆文加热时间类型：实时｜按周｜按月｜近7天｜近30天；用于第五步达人及内容复盘",
    )
    kol_content_date_range: Optional[str] = Field(
        None,
        description="爆文加热时间范围（按周/按月/自定义时的具体值）；实时/近7天/近30天可为 null",
    )
    # 行业搜索洞察（第六步）：一到多值（行业使用系统默认，不抽取）
    insight_time_ranges: Optional[list[InsightTimeRange]] = Field(
        None,
        description="行业搜索洞察的一到多个时间范围；每项含 date_range_type(按周|按月|近7天|近30天|自定义) 与 insight_date_range",
    )
    brand_name: Optional[str] = Field(None, description="The brand name related to the task")
    consumption_amount_star: Optional[float] = Field(
        None,
        description="星图（达人营销）消耗金额（元）；单报告模式时必填；品牌分类模式时由各分类内 consumption_amount_star 提供",
    )
    consumption_amount_bid: Optional[float] = Field(
        None,
        description="竞价（竞价投放）消耗金额（元）；单报告模式时必填；品牌分类模式时由各分类内 consumption_amount_bid 提供",
    )

async def extract_intent(
    task_description: str,
    conversation_history: list[dict] | None = None,
    previous_intent: YuntuTask | None = None
) -> YuntuTask:
    """
    Extracts the intent from the user's task description using Gemini.
    
    Args:
        task_description: Current user message
        conversation_history: Previous conversation messages for context
        previous_intent: Previously extracted intent values to preserve
        
    Returns:
        YuntuTask with extracted intent fields
    """
    try:
        # Initialize Gemini model
        llm = ChatGoogle(model=GEMINI_MODEL, temperature=0.0)
        
        # Build context from conversation history
        context_parts = []
        if conversation_history:
            context_parts.append("=== 对话历史 ===")
            for msg in conversation_history[-5:]:  # Last 5 messages for context
                role = msg.get("role", "unknown")
                content = msg.get("content", "")
                if role == "user":
                    context_parts.append(f"用户: {content}")
                elif role == "assistant":
                    context_parts.append(f"助手: {content}")
            context_parts.append("")
        
        # Include previous intent values if available
        if previous_intent:
            context_parts.append("=== 已识别的信息 ===")
            if previous_intent.report_name:
                context_parts.append(f"报告名: {previous_intent.report_name}")
            if previous_intent.brand_categories:
                for i, c in enumerate(previous_intent.brand_categories, 1):
                    name = c.category_name or f"分类{i}"
                    star_amt = f" {c.consumption_amount_star}元" if c.consumption_amount_star is not None else ""
                    bid_amt = f" {c.consumption_amount_bid}元" if c.consumption_amount_bid is not None else ""
                    context_parts.append(f"品牌分类-{name}: 星图报告={c.report_name_star}{star_amt}, 竞价报告={c.report_name_bid}{bid_amt}")
            if previous_intent.brand_name:
                context_parts.append(f"品牌: {previous_intent.brand_name}")
            if previous_intent.kol_content_date_type:
                r = previous_intent.kol_content_date_range or "（无具体日期）"
                context_parts.append(f"爆文加热时间: {previous_intent.kol_content_date_type} {r}")
            if previous_intent.insight_time_ranges:
                for i, tr in enumerate(previous_intent.insight_time_ranges, 1):
                    r = tr.insight_date_range or "（无具体日期）"
                    context_parts.append(f"行业搜索洞察-时间范围{i}: {tr.date_range_type} {r}")
            if previous_intent.consumption_amount_star is not None:
                context_parts.append(f"星图（达人营销）消耗金额: {previous_intent.consumption_amount_star} 元")
            if previous_intent.consumption_amount_bid is not None:
                context_parts.append(f"竞价（竞价投放）消耗金额: {previous_intent.consumption_amount_bid} 元")
            context_parts.append("")
            context_parts.append("注意：如果当前消息中没有提供新值，请保留上述已识别的信息。")
            context_parts.append("")
        
        # Build the full prompt
        context_text = "\n".join(context_parts) if context_parts else ""
        
        user_prompt = f"""{context_text}=== 当前用户消息 ===
{task_description}

请从当前消息和对话历史中提取信息，如果之前已经识别过某些字段，请保留它们（除非当前消息提供了新的值）。"""
        
        # Prepare the prompt using proper message types
        messages = [
            SystemMessage(content=INTENT_EXTRACTION_SYSTEM_PROMPT),
            UserMessage(content=user_prompt),
        ]
        
        # Invoke the model
        response = await llm.ainvoke(messages)
        content = response.completion
        
        # Clean up markdown code blocks if present
        if "```json" in content:
            content = content.split("```json")[1].split("```")[0].strip()
        elif "```" in content:
            content = content.split("```")[1].split("```")[0].strip()
            
        data = json.loads(content)
        
        # Merge with previous intent values if available
        if previous_intent:
            merged_data = previous_intent.model_dump()
            for key, value in data.items():
                if value is None or value == "":
                    continue
                if key == "insight_time_ranges":
                    if isinstance(value, list) and len(value) > 0:
                        merged_data[key] = value
                    continue
                if key == "brand_categories":
                    if isinstance(value, list) and len(value) > 0:
                        merged_data[key] = value
                    continue
                merged_data[key] = value
            task = YuntuTask(**merged_data)
        else:
            task = YuntuTask(**data)
        return task
        
    except Exception as e:
        print(f"Error extracting intent: {e}")
        # Return empty task on failure, app logic will handle missing fields
        return YuntuTask(
            report_name=None,
            brand_categories=None,
            kol_content_date_type=None,
            kol_content_date_range=None,
            insight_time_ranges=None,
            brand_name=None,
            consumption_amount_star=None,
            consumption_amount_bid=None,
        )
