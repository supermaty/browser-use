from __future__ import annotations

import json
import re
from typing import Literal

from pydantic import BaseModel, Field
from browser_use.llm.google import ChatGoogle
from browser_use.llm.messages import SystemMessage, UserMessage

from config import INTENT_EXTRACTION_SYSTEM_PROMPT, GEMINI_MODEL


_REPORT_NAME_PATTERN = r"([A-Za-z0-9_\-\u4e00-\u9fff\.]+)"
_AMOUNT_PATTERN = r"([0-9]+(?:\.[0-9]+)?\s*(?:万|w|W|元)?)"
_CH_REPORT = "\u62a5\u544a"
_CH_REPORT_NAME = "\u62a5\u544a\u540d"
_CH_REPORT_NAME2 = "\u62a5\u544a\u540d\u79f0"
_CH_STAR = "\u661f\u56fe"
_CH_BID = "\u7ade\u4ef7"
_CH_SPEND = "\u6d88\u8017"
_CH_AMOUNT = "\u91d1\u989d"
_CH_BUDGET = "\u9884\u7b97"
_CH_COST = "\u82b1\u8d39"
_CH_CATEGORY = "\u5206\u7c7b"
_CH_BRAND_CATEGORY = "\u54c1\u724c\u5206\u7c7b"
_CH_SEGMENT = "\u54c1\u7c7b"
_NAME_STOPWORDS = {
    _CH_REPORT,
    _CH_REPORT_NAME,
    _CH_REPORT_NAME2,
    _CH_STAR,
    _CH_BID,
    _CH_SPEND,
    _CH_AMOUNT,
    _CH_BUDGET,
    _CH_COST,
}


def _parse_amount_token(raw: str | None) -> float | None:
    text = str(raw or "").strip()
    if not text:
        return None
    text = text.replace(",", "").replace("，", "").replace(" ", "")
    multiplier = 1.0
    if text.endswith(("万", "w", "W")):
        multiplier = 10000.0
        text = text[:-1]
    if text.endswith("元"):
        text = text[:-1]
    try:
        return float(text) * multiplier
    except Exception:
        return None


def _clean_name_token(raw: str | None) -> str:
    name = str(raw or "").strip().strip("：:;；,，。")
    if not name:
        return ""
    if name in _NAME_STOPWORDS:
        return ""
    return name


def _extract_amount_in_scope(scope: str) -> float | None:
    m = re.search(_AMOUNT_PATTERN, scope or "", flags=re.IGNORECASE)
    if not m:
        return None
    return _parse_amount_token(m.group(1))


def _parse_category_line(line: str, index: int) -> dict | None:
    text = re.sub(r"\s+", " ", str(line or "")).strip()
    if not text or (_CH_STAR not in text) or (_CH_BID not in text):
        return None

    star = re.search(rf"{_CH_STAR}(?:{_CH_REPORT})?[:：]?\s*{_REPORT_NAME_PATTERN}", text, flags=re.IGNORECASE)
    bid = re.search(rf"{_CH_BID}(?:{_CH_REPORT})?[:：]?\s*{_REPORT_NAME_PATTERN}", text, flags=re.IGNORECASE)
    if not star or not bid:
        return None

    star_name = _clean_name_token(star.group(1))
    bid_name = _clean_name_token(bid.group(1))
    if not star_name or not bid_name:
        return None

    # Prefer local amounts near each report block.
    star_scope = text[star.end() : bid.start()] if bid.start() > star.end() else text[star.end() :]
    bid_scope = text[bid.end() :]
    star_amt = _extract_amount_in_scope(star_scope)
    bid_amt = _extract_amount_in_scope(bid_scope)

    # Fallback: explicit "星图消耗/竞价消耗" style.
    if star_amt is None:
        m = re.search(rf"{_CH_STAR}(?:{_CH_SPEND}|{_CH_COST}|{_CH_BUDGET})?[:：]?\s*{_AMOUNT_PATTERN}", text, flags=re.IGNORECASE)
        if m:
            star_amt = _parse_amount_token(m.group(1))
    if bid_amt is None:
        m = re.search(rf"{_CH_BID}(?:{_CH_SPEND}|{_CH_COST}|{_CH_BUDGET})?[:：]?\s*{_AMOUNT_PATTERN}", text, flags=re.IGNORECASE)
        if m:
            bid_amt = _parse_amount_token(m.group(1))
    if star_amt is None or bid_amt is None:
        return None

    prefix = text[: star.start()].strip()
    prefix = re.sub(rf"^({_CH_BRAND_CATEGORY}|{_CH_CATEGORY}|{_CH_SEGMENT})[:：]?\s*", "", prefix)
    category_name = prefix.strip() or f"分类{index}"

    return {
        "category_name": category_name,
        "report_name_star": star_name,
        "consumption_amount_star": float(star_amt),
        "report_name_bid": bid_name,
        "consumption_amount_bid": float(bid_amt),
    }


def _extract_brand_categories_from_text(task_description: str) -> list[dict]:
    text = str(task_description or "")
    if not text:
        return []
    normalized = text.replace("\uff1b", ";").replace("\r", "\n")
    chunks = [part.strip() for part in re.split(r"[;\n]+", normalized) if part.strip()]

    items: list[dict] = []
    for idx, chunk in enumerate(chunks, 1):
        parsed = _parse_category_line(chunk, idx)
        if parsed:
            items.append(parsed)

    # Global fallback: try parsing the whole sentence if per-line split failed.
    if not items:
        parsed = _parse_category_line(normalized, 1)
        if parsed:
            items.append(parsed)

    # Deduplicate by (star,bid) pair.
    dedup: list[dict] = []
    seen: set[tuple[str, str]] = set()
    for item in items:
        key = (str(item.get("report_name_star") or "").strip(), str(item.get("report_name_bid") or "").strip())
        if not key[0] or not key[1] or key in seen:
            continue
        seen.add(key)
        dedup.append(item)
    return dedup


def _extract_report_names_from_text(task_description: str) -> list[str]:
    text = str(task_description or "")
    if not text:
        return []
    names: list[str] = []
    for m in re.finditer(rf"(?:{_CH_REPORT_NAME}|{_CH_REPORT_NAME2})[:：]\s*([^\n;\uff1b]+)", text):
        block = str(m.group(1) or "").strip()
        for part in re.split(r"[\u3001,\uff0c/|]", block):
            token = _clean_name_token(part)
            if token:
                names.append(token)
    dedup: list[str] = []
    seen: set[str] = set()
    for n in names:
        if n in seen:
            continue
        seen.add(n)
        dedup.append(n)
    return dedup


def _apply_multi_report_fallback(data: dict, task_description: str) -> dict:
    if not isinstance(data, dict):
        return data
    out = dict(data)

    raw_categories = out.get("brand_categories")
    has_valid_categories = isinstance(raw_categories, list) and len(raw_categories) > 0
    if has_valid_categories:
        # Keep LLM output unless key fields are obviously missing.
        complete = True
        for c in raw_categories:
            if not isinstance(c, dict):
                complete = False
                break
            if not (
                c.get("report_name_star")
                and c.get("report_name_bid")
                and c.get("consumption_amount_star") is not None
                and c.get("consumption_amount_bid") is not None
            ):
                complete = False
                break
        if complete:
            return out

    parsed_categories = _extract_brand_categories_from_text(task_description)
    if parsed_categories:
        out["brand_categories"] = parsed_categories
        # In multi-category mode, top-level report_name should not block validation path.
        if len(parsed_categories) > 1:
            out["report_name"] = None
        return out

    # Fallback for "多个报告名" but no category structure: keep first as report_name.
    if not out.get("report_name"):
        names = _extract_report_names_from_text(task_description)
        if names:
            out["report_name"] = names[0]
    return out

# 爆文加热（第五步 达人及内容复盘）：单值，时间类型 + 时间范围
KOL_CONTENT_DATE_TYPE = Literal["实时", "按周", "按月", "近7天", "近30天"]
# 行业搜索洞察（第六步）：一到多值，时间类型 + 时间范围
INSIGHT_DATE_TYPE = Literal["按周", "按月", "近7天", "近30天", "自定义"]


class InsightTimeRange(BaseModel):
    """行业搜索洞察的单个时间范围"""
    date_range_type: INSIGHT_DATE_TYPE
    insight_date_range: str | None = Field(None, description="自定义/按周/按月时的具体日期范围字符串")


class BrandCategory(BaseModel):
    """黑玩品牌案例：单个品牌分类，含星图/竞价报告名及各自消耗金额"""
    category_name: str | None = Field(None, description="品牌分类名称（可选），如「洗护」「美发」")
    report_name_star: str | None = Field(None, description="该分类的星图（达人营销）报告名")
    consumption_amount_star: float | None = Field(None, description="该分类星图报告消耗金额（元），必填")
    report_name_bid: str | None = Field(None, description="该分类的竞价（竞价投放）报告名")
    consumption_amount_bid: float | None = Field(None, description="该分类竞价报告消耗金额（元），必填")


class YuntuTask(BaseModel):
    report_name: str | None = Field(None, description="单个报告名（非黑玩多分类时使用）")
    brand_categories: list[BrandCategory] | None = Field(
        None,
        description="黑玩品牌案例：多个品牌分类，每分类含星图报告名+消耗金额、竞价报告名+消耗金额；与 report_name 二选一",
    )
    # 爆文加热时间（第五步 达人及内容复盘）：单值
    kol_content_date_type: KOL_CONTENT_DATE_TYPE | None = Field(
        None,
        description="爆文加热时间类型：实时｜按周｜按月｜近7天｜近30天；用于第五步达人及内容复盘",
    )
    kol_content_date_range: str | None = Field(
        None,
        description="爆文加热时间范围（按周/按月/自定义时的具体值）；实时/近7天/近30天可为 null",
    )
    # 行业搜索洞察（第六步）：一到多值（行业使用系统默认，不抽取）
    insight_time_ranges: list[InsightTimeRange] | None = Field(
        None,
        description="行业搜索洞察的一到多个时间范围；每项含 date_range_type(按周|按月|近7天|近30天|自定义) 与 insight_date_range",
    )
    brand_name: str | None = Field(None, description="The brand name related to the task")
    consumption_amount_star: float | None = Field(
        None,
        description="星图（达人营销）消耗金额（元）；单报告模式时必填；品牌分类模式时由各分类内 consumption_amount_star 提供",
    )
    consumption_amount_bid: float | None = Field(
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
        print(f"user_prompt: {user_prompt}")
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
        data = _apply_multi_report_fallback(data, task_description)
        print(f"data: {data}")
        
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
