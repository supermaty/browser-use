"""
Intent validation logic for Yuntu Data Scraper.
"""
from intent import YuntuTask

class ValidationResult:
    def __init__(self, is_valid: bool, message: str = ""):
        self.is_valid = is_valid
        self.message = message

def validate_intent(task_intent: YuntuTask) -> ValidationResult:
    """
    Validates the extracted intent.
    
    Args:
        task_intent: The intent to validate.
        
    Returns:
        ValidationResult indicating success or failure with an error message.
    """
    categories = task_intent.brand_categories or []
    has_single = bool(task_intent.report_name)
    has_categories = bool(categories) and all(
        getattr(c, "report_name_star", None)
        and getattr(c, "report_name_bid", None)
        and getattr(c, "consumption_amount_star", None) is not None
        and getattr(c, "consumption_amount_bid", None) is not None
        for c in categories
    )

    # 1. Check report info existence
    if not has_single and not has_categories:
        return ValidationResult(False, (
            "❌ **无法识别报告信息或消耗金额**\n\n"
            "请任选一种方式提供：\n"
            "1. 单个报告名 + 消耗金额：报告名：******；星图 10 万、竞价 20 万\n"
            "2. 品牌分类（每分类含星图报告名+消耗金额、竞价报告名+消耗金额）：\n"
            "   品牌分类：分类A 星图报告xxx 10万 竞价报告yyy 20万；分类B 星图报告xxx 消耗金额 竞价报告yyy 消耗金额"
        ))

    # 2. Check consumption amount for single report mode
    if has_single and (
        task_intent.consumption_amount_star is None or task_intent.consumption_amount_bid is None
    ):
        missing = []
        if task_intent.consumption_amount_star is None:
            missing.append("星图（达人营销）消耗金额")
        if task_intent.consumption_amount_bid is None:
            missing.append("竞价（竞价投放）消耗金额")
        return ValidationResult(False, (
            "❌ **缺少消耗金额（必填）**\n\n"
            f"请提供：{'、'.join(missing)}。\n\n"
            "示例：星图 10 万、竞价 20 万；或「星图消耗 100000 元，竞价消耗 200000 元」"
        ))

    # 3. Check KOL content date (Step 5) for Heyone brand
    if task_intent.brand_name and "黑玩" in (task_intent.brand_name or ""):
        if not task_intent.kol_content_date_type:
            return ValidationResult(False, (
                "❌ **缺少爆文加热时间（第五步 达人及内容复盘）**\n\n"
                "请提供时间类型，例如：\n"
                "- 爆文加热：近30天\n"
                "- 爆文加热：按月（2025/12/01-2025/12/31）\n\n"
                "支持的类型：实时、按周、按月、近7天、近30天(实时/近7天/近30天可不提供具体时间范围)"
            ))

    # 4. Check Insight time range (Step 6)
    time_ranges = task_intent.insight_time_ranges or []
    if not time_ranges:
        return ValidationResult(False, (
            "❌ **缺少行业搜索洞察的时间范围信息（第六步）**\n\n"
            "请提供至少一个时间范围，例如：\n"
            "- 行业搜索洞察：自定义（2025/12/01-2025/12/31）\n"
            "- 行业搜索洞察：近7天、近30天\n\n"
            "支持的类型：按周、按月、近7天、近30天、自定义"
        ))

    return ValidationResult(True)
