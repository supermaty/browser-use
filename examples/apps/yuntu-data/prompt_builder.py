"""
Prompt assembly and task summary generation for Yuntu Data Scraper.
"""
from intent import YuntuTask
from config import HEYONE_PROMPT_TEMPLATE, GENERIC_PROMPT_TEMPLATE, get_brand_english_name

def build_prompt(task_intent: YuntuTask) -> str:
    """
    Assembles the final prompt for the agent based on the task intent.
    
    Args:
        task_intent: The extracted task intent.
        
    Returns:
        The formatted prompt string.
    """
    if task_intent.brand_name and "黑玩" in task_intent.brand_name:
        template = HEYONE_PROMPT_TEMPLATE
    else:
        template = GENERIC_PROMPT_TEMPLATE
        
    # Process time ranges
    time_ranges = task_intent.insight_time_ranges or []
    insight_time_ranges_formatted = "\n".join(
        f"{i}. {tr.date_range_type}" + (f"（{tr.insight_date_range}）" if tr.insight_date_range else "")
        for i, tr in enumerate(time_ranges, 1)
    )
    first_insight = time_ranges[0] if time_ranges else None
    
    # Process KOL content date
    kol_type = task_intent.kol_content_date_type or "N/A"
    kol_range = task_intent.kol_content_date_range or "（无具体日期）"
    
    # Process brand categories
    categories = task_intent.brand_categories or []
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
    
    # Process brand name
    brand_display = task_intent.brand_name or "未指定"
    brand_name_en = get_brand_english_name(task_intent.brand_name)
    
    # Assemble final prompt
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
    
    return final_prompt

def build_task_summary(task_intent: YuntuTask) -> str:
    """
    Generates a readable summary of the task for the UI.
    
    Args:
        task_intent: The extracted task intent.
        
    Returns:
        Formatted task summary string.
    """
    brand_info = "（使用黑玩专属模板）" if task_intent.brand_name and "黑玩" in task_intent.brand_name else ""
    categories = task_intent.brand_categories or []
    time_ranges = task_intent.insight_time_ranges or []
    
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
    
    if categories:
        report_block = (
            f"- **品牌分类（星图+竞价报告+消耗金额）**:\n"
            + "\n".join(
                f"  - {getattr(c, 'category_name', None) or f'分类{i}'}: 星图={getattr(c, 'report_name_star', '')} {_fmt_amt(getattr(c, 'consumption_amount_star', None))}，竞价={getattr(c, 'report_name_bid', '')} {_fmt_amt(getattr(c, 'consumption_amount_bid', None))}"
                for i, c in enumerate(categories, 1)
            )
            + "\n"
        )
    else:
        report_block = f"- **报告名**: {task_intent.report_name or '未指定'}\n"
        
    task_summary = (
        f"📋 **任务摘要** {brand_info}\n\n"
        f"{report_block}"
        f"- **品牌**: {task_intent.brand_name or '未指定'}\n"
        f"{kol_line}"
        f"{time_ranges_block}"
        f"{consumption_lines}\n"
        f"🚀 正在启动 Agent 执行任务..."
    )
    
    return task_summary
