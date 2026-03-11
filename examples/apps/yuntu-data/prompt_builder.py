from __future__ import annotations

import re
from datetime import date, datetime, timedelta
from typing import Any

from intent import YuntuTask
from specs import BrandProfileSpec


def _collect_report_targets(task_intent: YuntuTask) -> list[str]:
    targets: list[str] = []
    if task_intent.report_name:
        targets.append(task_intent.report_name)
    for category in task_intent.brand_categories or []:
        if category.report_name_star:
            targets.append(category.report_name_star)
        if category.report_name_bid:
            targets.append(category.report_name_bid)
    return list(dict.fromkeys([name for name in targets if name]))


def _format_targets(targets: list[str]) -> str:
    if not targets:
        return "(none)"
    return "\n".join(f"- {item}" for item in targets)


def _join_non_empty(values: list[str], sep: str = " > ") -> str:
    cleaned = [str(v).strip() for v in values if str(v).strip()]
    return sep.join(cleaned)


def _infer_touchpoint_from_report_name(report_name: str | None) -> str | None:
    text = (report_name or "").strip()
    if not text:
        return None
    if "竞价" in text:
        return "全部/竞价投放"
    if "星图" in text:
        return "全部/达人营销"
    return None


def _parse_date_token(token: str | None) -> date | None:
    text = (token or "").strip()
    if not text:
        return None
    text = text.replace("/", "-").replace(".", "-")
    text = re.sub(r"\s+", "", text)
    try:
        return datetime.strptime(text, "%Y-%m-%d").date()
    except Exception:
        return None


def _extract_date_pair(date_range: str | None) -> tuple[date, date] | None:
    text = (date_range or "").strip()
    if not text:
        return None
    tokens = re.findall(r"\d{4}[./-]\d{1,2}[./-]\d{1,2}", text)
    if len(tokens) >= 2:
        start = _parse_date_token(tokens[0])
        end = _parse_date_token(tokens[1])
        if start and end and end >= start:
            return start, end
    normalized = text.replace("至", "~").replace("—", "~").replace("–", "~")
    parts = [part.strip() for part in normalized.split("~", 1)]
    if len(parts) != 2:
        return None
    start = _parse_date_token(parts[0])
    end = _parse_date_token(parts[1])
    if not start or not end or end < start:
        return None
    return start, end


def _fmt_date(d: date) -> str:
    return d.strftime("%Y.%m.%d")


def _build_insight_cycles(data: dict[str, Any]) -> tuple[list[dict[str, Any]], str]:
    raw_ranges = data.get("insight_time_ranges") or []
    if not isinstance(raw_ranges, list):
        raw_ranges = []

    cycles: list[dict[str, Any]] = []
    lines: list[str] = []
    for i, item in enumerate(raw_ranges, 1):
        if not isinstance(item, dict):
            continue
        date_range_type = str(item.get("date_range_type") or "").strip()
        insight_date_range = str(item.get("insight_date_range") or "").strip()
        current_insight = {
            "date_range_type": date_range_type,
            "insight_date_range": insight_date_range,
        }
        previous_insight = None
        parsed = _extract_date_pair(insight_date_range)
        if parsed is not None:
            start, end = parsed
            span_days = (end - start).days + 1
            previous_end = start - timedelta(days=1)
            previous_start = previous_end - timedelta(days=span_days - 1)
            previous_insight = {
                "date_range_type": "自定义",
                "insight_date_range": f"{_fmt_date(previous_start)} ~ {_fmt_date(previous_end)}",
            }

        cycles.append(
            {
                "index": i,
                "current_insight": current_insight,
                "previous_insight": previous_insight,
            }
        )

        line = f"{i}. 当前周期: {date_range_type}"
        if insight_date_range:
            line += f"（{insight_date_range}）"
        if previous_insight:
            line += f" | 上一周期: 自定义（{previous_insight['insight_date_range']}）"
        lines.append(line)

    formatted = "\n".join(lines) if lines else ""
    return cycles, formatted


def _intent_context(task_intent: YuntuTask) -> dict[str, Any]:
    data = task_intent.model_dump(mode="python", exclude_none=False)
    if not data.get("kol_content_date_type"):
        data["kol_content_date_type"] = "近30天"
    if not data.get("touchpoint"):
        data["touchpoint"] = _infer_touchpoint_from_report_name(data.get("report_name"))

    insight_cycles, insight_time_ranges_formatted = _build_insight_cycles(data)
    data["insight_cycles"] = insight_cycles
    data["insight_time_ranges_formatted"] = insight_time_ranges_formatted
    if insight_cycles:
        data["current_insight"] = insight_cycles[0].get("current_insight") or {}
    if not isinstance(data.get("current_insight"), dict):
        data["current_insight"] = {
            "date_range_type": "",
            "insight_date_range": "",
        }
    return {"params": data, "context": data}


def _context_get(path: str | None, context: dict[str, Any] | None) -> Any:
    if not context or not path:
        return None
    raw = path.strip()
    if not raw:
        return None
    if not raw.startswith(("params.", "context.")):
        raw = f"context.{raw}"

    current: Any = context
    for part in raw.split("."):
        if part in {"params", "context"}:
            if isinstance(current, dict) and part in current:
                current = current.get(part)
                continue
            return None
        if isinstance(current, dict) and part in current:
            current = current.get(part)
            continue
        return None
    return current


def _resolve_display_value(
    *,
    value: str,
    value_from: str,
    path: list[str],
    path_from: str,
    context: dict[str, Any] | None,
) -> str:
    if value_from:
        resolved = _context_get(value_from, context)
        if isinstance(resolved, list):
            return _join_non_empty([str(v).strip() for v in resolved if str(v).strip()], " / ")
        if resolved is not None and str(resolved).strip():
            return str(resolved).strip()
    if value:
        return value
    if path_from:
        resolved = _context_get(path_from, context)
        if isinstance(resolved, list):
            return _join_non_empty([str(v).strip() for v in resolved if str(v).strip()], " / ")
        if resolved is not None and str(resolved).strip():
            return str(resolved).strip()
    return _join_non_empty(path, " / ")


def _format_step(step, context: dict[str, Any] | None = None) -> str:
    op = getattr(step, "op", "")
    target = (getattr(step, "target", None) or "").strip()
    value = (getattr(step, "value", None) or "").strip()
    value_from = (getattr(step, "value_from", None) or "").strip()
    path = [str(v).strip() for v in (getattr(step, "path", None) or []) if str(v).strip()]
    path_from = (getattr(step, "path_from", None) or "").strip()
    field_keys = [str(v).strip() for v in (getattr(step, "field_keys", None) or []) if str(v).strip()]
    file_keys = [str(v).strip() for v in (getattr(step, "file_keys", None) or []) if str(v).strip()]
    file_key = (getattr(step, "file_key", None) or "").strip()

    if op == "ensure_visible":
        return f"[ensure_visible] `{target}`"
    if op == "click":
        return f"[click] `{target}`"
    if op == "hover":
        if field_keys:
            return f"[hover] {', '.join(field_keys)}"
        return f"[hover] `{target}`"
    if op == "select":
        value_text = _resolve_display_value(
            value=value,
            value_from=value_from,
            path=path,
            path_from=path_from,
            context=context,
        )
        display_target = target or "当前日期类型下拉"
        if not value_text:
            value_text = "(skip if empty)"
        return f"[select] `{display_target}` = `{value_text}`"
    if op == "select_path":
        value_text = _resolve_display_value(
            value=value,
            value_from=value_from,
            path=path,
            path_from=path_from,
            context=context,
        )
        return f"[select_path] `{target}` -> `{value_text}`"
    if op == "input":
        value_text = _resolve_display_value(
            value=value,
            value_from=value_from,
            path=path,
            path_from=path_from,
            context=context,
        )
        return f"[input] `{target}` = `{value_text}`"
    if op == "set_date_range":
        value_text = _resolve_display_value(
            value=value,
            value_from=value_from,
            path=path,
            path_from=path_from,
            context=context,
        )
        if not value_text:
            value_text = "(skip if empty)"
        return f"[set_date_range] `{target or '日期'}` = `{value_text}`"
    if op == "extract_fields":
        return f"[extract] {', '.join(field_keys)}"
    if op == "download":
        keys = file_keys or ([file_key] if file_key else [])
        return f"[download] {', '.join(keys) if keys else '(module default files)'}"
    if op == "wait":
        seconds = getattr(step, "seconds", None)
        return f"[wait] {seconds if seconds is not None else 0.6}s"
    return f"[{op}] `{target}`"


def _format_module_plan(modules: list, title: str, context: dict[str, Any] | None = None) -> str:
    if not modules:
        return f"{title}\n- (none)"

    lines: list[str] = [title]
    for module in modules:
        lines.append(f"- module_key: `{module.module_key}` ({module.label})")
        route_text = _join_non_empty(list(module.route_path or []))
        lines.append(f"  route_path: {route_text if route_text else '(none)'}")

        steps = list(module.interaction_steps or [])
        if steps:
            lines.append("  interaction_steps:")
            for idx, step in enumerate(steps, 1):
                lines.append(f"  {idx}. {_format_step(step, context=context)}")
        else:
            lines.append("  interaction_steps: (none)")

    return "\n".join(lines)


def _format_insight_cycle_plan(context: dict[str, Any] | None) -> str:
    if not isinstance(context, dict):
        return "- (none)"
    params = context.get("params", {})
    if not isinstance(params, dict):
        return "- (none)"
    cycles = params.get("insight_cycles") or []
    if not isinstance(cycles, list) or not cycles:
        return "- (none)"

    lines: list[str] = []
    for cycle in cycles:
        if not isinstance(cycle, dict):
            continue
        idx = cycle.get("index")
        current = cycle.get("current_insight") if isinstance(cycle.get("current_insight"), dict) else {}
        previous = cycle.get("previous_insight") if isinstance(cycle.get("previous_insight"), dict) else None
        cur_type = str(current.get("date_range_type") or "").strip() or "N/A"
        cur_range = str(current.get("insight_date_range") or "").strip() or "N/A"
        lines.append(f"- cycle{idx}_current: type=`{cur_type}`, range=`{cur_range}`")
        if isinstance(previous, dict):
            pre_type = str(previous.get("date_range_type") or "").strip() or "N/A"
            pre_range = str(previous.get("insight_date_range") or "").strip() or "N/A"
            lines.append(f"- cycle{idx}_previous: type=`{pre_type}`, range=`{pre_range}`")
    return "\n".join(lines) if lines else "- (none)"


def build_phase_a_prompt(task_intent: YuntuTask, profile: BrandProfileSpec, brand_name_en: str) -> str:
    brand_name = task_intent.brand_name or "UNKNOWN_BRAND"
    targets = _collect_report_targets(task_intent)
    first_report = targets[0] if targets else (task_intent.report_name or "")
    menu_text = _join_non_empty(list(profile.menu_path or [])) or "营销决策 -> 结案报告"

    return f"""
You are a web operator for Yuntu navigation.

Goal:
- Open Yuntu
- Ensure login state
- Switch brand to `{brand_name}` (alias: `{brand_name_en}`)
- Navigate menu: {menu_text}
- Open first report `{first_report}` in current tab

Preferred way:
1. Use native page actions only (click/input/search/enter) to navigate/open report.
2. After entering report list, click and stay on `升级版报告` tab only (never switch to `历史报告` / `全部报告`).
3. Report-list search must be the table-adjacent search input (NOT top global search).
4. Exact report match required (allow ignoring trailing punctuation only).
5. Open report in current tab.

Report targets:
{_format_targets(targets)}
"""


def build_phase_b_prompt(task_intent: YuntuTask, profile: BrandProfileSpec) -> str:
    targets = _collect_report_targets(task_intent)
    report_modules = [m.module_key for m in profile.module_specs if m.requires_report]
    global_modules = [m.module_key for m in profile.module_specs if not m.requires_report]

    return f"""
You are a deterministic extractor for Yuntu.

Rules:
1. Execute module DSL steps agent-first.
1.1 For each module pass, execute all plain `route_path` nodes in listed order first with native page actions (native click/scroll), then run interaction steps.
1.2 If a route node cannot be clicked natively after scroll retry, fallback to `click_in_content` for that node.
1.3 Never start `[select]` / `[set_date_range]` / `[extract]` / `[download]` if the last `route_path` node for that module is not reached.
1.4 Route-node success evidence must be action-based (native click/scroll effect or `click_in_content_result.clicked=true`); seeing node text alone is NOT success evidence.
1.5 For explicit DSL `[click]` interaction steps, use `click_in_content(target_text=<dsl target>)`; do NOT use raw `click index`.
1.6 If browser warns `Element index ... not available`, do NOT refresh-loop; immediately retry with `click_in_content` for that DSL target.
2. Use native `extract` for `[extract]`.
2.1 For `[hover]`, use native `evaluate` then native `extract` for the same hover field list.
   Hover requirements:
   - Use `[hover]` field semantic text to find the target card/container first.
   - Parse hover field into semantic anchors (split by delimiters like `/`, `-`, `(`, `)`, `、`, whitespace) and match by conjunction, not single-fragment hit.
   - Prefer the container that satisfies the most anchors; for multi-anchor fields, require both primary metric anchor and qualifier anchor before accepting.
   - If only suffix/qualifier anchors match (for example only benchmark phrase), treat as partial match and reject.
   - If field text contains identifier tokens, lock hover to the container with the same identifier tokens.
   - Hover only help/question/popover triggers in that container (`?`, help, question, popover/tooltip trigger), never generic icons (`+`, expand, chart, collapse).
   - Never use broad icon selection like `querySelector('svg')` as hover target.
   - In the target container, choose the nearest tooltip trigger to the semantic anchor text (shortest geometric distance), not the first matched node.
   - `evaluate` must return structured evidence (JSON string): `hover_ok`, `matched_container_text`, `matched_trigger_text`, `tooltip_visible`, `tooltip_text`, `matched_anchors`, `missing_anchors`.
   - Treat hover as success only when `tooltip_visible=true` and `tooltip_text` is non-empty and semantically aligned with the hover field.
   - If evidence check fails, mark this hover field as missing immediately and continue to next DSL step (no retry loop).
   - After evidence passes, run native `extract` and accept only values from the visible tooltip/popup context.
2.2 Use `download_files_by_spec` for `[download]`.
3. Always set `profile_name` to `{profile.profile_name}` in every tool call.
4. If current page is not target report detail, recover with native page actions only.
5. Execute by report target loop in exact order:
   - open one target report from report list (current tab)
   - execute report modules in order: {report_modules}
   - return to report list and continue next target
6. After ALL report targets are completed, execute global modules once in order: {global_modules}
7. `run_module_collection` is DISABLED. Never call it.
8. `extract_fields_by_spec` is DISABLED in report-phase DSL execution. Never call it.
9. `hover_in_content` is DISABLED in report-phase DSL execution. Never call it.
10. Return schema-valid JSON only.
11. After opening each report detail page, run native `extract` once for `计算时间区间` / `计算周期`.
12. Save that period in memory as `current_report_period` and reuse it for downstream calculations.

Report targets:
{_format_targets(targets)}
"""


def build_single_agent_prompt(task_intent: YuntuTask, profile: BrandProfileSpec, brand_name_en: str) -> str:
    brand_name = task_intent.brand_name or "UNKNOWN_BRAND"
    targets = _collect_report_targets(task_intent)
    first_report = targets[0] if targets else (task_intent.report_name or "")
    report_targets_count = len(targets)
    report_modules = [m.module_key for m in profile.module_specs if m.requires_report]
    global_modules = [m.module_key for m in profile.module_specs if not m.requires_report]
    has_report_modules = bool(report_modules)
    report_module_specs = [m for m in profile.module_specs if m.requires_report]
    global_module_specs = [m for m in profile.module_specs if not m.requires_report]
    render_context = _intent_context(task_intent)
    report_plan = _format_module_plan(
        report_module_specs,
        "Report Module DSL Execution Plan",
        context=render_context,
    )
    global_plan = _format_module_plan(
        global_module_specs,
        "Global Module DSL Execution Plan",
        context=render_context,
    )
    insight_cycle_plan = _format_insight_cycle_plan(render_context)
    first_report_line = f"- first_report: `{first_report}`" if has_report_modules else ""
    report_targets_line = f"- report_targets_count: `{report_targets_count}`" if has_report_modules else ""
    pre_report_rule = (
        "1. Before report detail is opened, DO NOT call any custom tool. Use native page actions only."
        if has_report_modules
        else "1. Do not force report-list flow. Start from current/home page and follow DSL route_path directly."
    )
    report_detail_rule = (
        "7. After report detail is opened, use DSL operations as follows:"
        if has_report_modules
        else "7. Use DSL operations directly from current page (no report-detail prerequisite)."
    )
    report_scope_rule = (
        "15. Report search scope rule: after entering report list, click and stay on `升级版报告` tab only. Never switch to `历史报告`/`全部报告` for this task."
        if has_report_modules
        else "15. Do not navigate to `营销决策/结案报告` unless DSL route_path explicitly contains it."
    )
    report_target_loop_rule = (
        "16. Multi-report mode (required): process `report_targets` sequentially in listed order. "
        "For each target: open that report in current tab -> run all report modules in order once -> return to report list -> next target."
        if has_report_modules
        else "16. Multi-report loop is disabled because there is no report-dependent module."
    )
    execution_a_title = "A) Open report list and enter first report detail" if has_report_modules else "A) Open Yuntu home and start DSL routes"
    execution_a_step = (
        ''
    )
    prompt_targets = targets if has_report_modules else []
    params_context = render_context.get("params", {}) if isinstance(render_context, dict) else {}
    insight_cycles = list(params_context.get("insight_cycles") or [])
    insight_ranges_count = len(list(task_intent.insight_time_ranges or []))
    first_cycle = insight_cycles[0] if insight_cycles else {}
    first_current = first_cycle.get("current_insight", {}) if isinstance(first_cycle, dict) else {}
    first_insight_type = first_current.get("date_range_type") if isinstance(first_current, dict) else None
    first_insight_range = first_current.get("insight_date_range") if isinstance(first_current, dict) else None
    insight_time_ranges_formatted = str(params_context.get("insight_time_ranges_formatted") or "N/A")
    inferred_touchpoint = _infer_touchpoint_from_report_name(first_report or task_intent.report_name)

    return f"""
You are a deterministic Yuntu operator. Complete the workflow in ONE run.

Fixed context:
- profile_name: `{profile.profile_name}`
- brand: `{brand_name}` (alias: `{brand_name_en}`)
{first_report_line}
{report_targets_line}
- kol_content_date_type: `{(task_intent.kol_content_date_type or "近30天")}`
- kol_content_date_range: `{(task_intent.kol_content_date_range or "N/A")}`
- touchpoint(auto from report_name): `{(inferred_touchpoint or "N/A")}`
- insight_time_ranges_count: `{insight_ranges_count}`
- current_insight.date_range_type: `{(first_insight_type or "N/A")}`
- current_insight.insight_date_range: `{(first_insight_range or "N/A")}`
- insight_time_ranges_formatted:
{insight_time_ranges_formatted}

Hard rules:
{pre_report_rule}
2. In every tool call, set `profile_name` to `{profile.profile_name}`.
3. Never navigate to `params.*` pseudo values.
4. Never open new tab/window.
5. Do not click unrelated global entries (e.g. message center).
6. `run_module_collection` is DISABLED. Never call it.
6.1 `extract_fields_by_spec` is DISABLED in report-phase DSL execution. Never call it.
6.2 `hover_in_content` is DISABLED in report-phase DSL execution. Never call it.
6.3 `dropdown_options` is DISABLED. Never call it.
{report_detail_rule}
   - `[click]/[ensure_visible]` -> for DSL interaction steps use `click_in_content` (target text from DSL); for plain route_path nodes keep native click/scroll first with `click_in_content` fallback.
   - `[select]/[select_path]` -> `select_in_content`
   - `[set_date_range]` -> native page click-select only (no manual text input)
   - Date rule for `[set_date_range]`:
     a) use calendar/date-picker click-select only; never type or inject input value by script.
     b) after selection, verify date box shows BOTH start and end dates (normalized format match is allowed).
     c) if mismatch, reopen date picker and retry this date step at most once; if still mismatch, mark missing and continue.
     d) for `[set_date_range]`, `path` must stay empty and date type must be handled by `[select]` step.
   - `[extract]/[extract_fields]` -> native `extract`
   - `[hover]` -> native `evaluate` hover tooltip trigger, then native `extract` for the same field list
   - `[download]` -> `download_files_by_spec`
8. Execute DSL strictly in order. One step at a time, one action at a time.
8.0 Do not skip plain `route_path` nodes:
   - In each module pass, process every plain route node in order before the first interaction step.
   - For clickable route nodes: native click first; if target not visible, scroll and retry; fallback to `click_in_content`.
   - For non-clickable route nodes (section/card headings): scroll until visible in main content.
   - If route tail is uncertain, redo route-tail positioning once before `[select]/[set_date_range]/[extract]/[download]`.
8.1 Route node completion evidence is mandatory:
   - A route node is completed only when there is action evidence for that exact node:
     native click changed active state/content OR `click_in_content_result.clicked=true` OR non-clickable heading became visible after scroll.
   - Never treat memory text like "reached xxx page" as completion evidence.
   - Without this evidence, do not run `[select]/[set_date_range]/[extract]/[download]`.
8.2 Stale-index guard:
   - Never use raw `click index` for explicit DSL `[click]` steps.
   - If a raw click fails with index-missing warning, switch to `click_in_content` immediately for the same target; do not repeat refresh attempts.
9. Single-pass policy: each DSL step runs at most once per pass (per cycle label). If failed, mark missing and continue (no loop retry).
10. If page drifts to message/notice center, go back once and continue from next DSL step.
11. Hover rule (generic): find container by field semantics, hover only tooltip/help/question trigger in that container, require visible tooltip, extract from tooltip text only.
12. Never claim values from memory; final values must come from current extract/tool output.
13. Keep model output SHORT to avoid MAX_TOKENS:
   - `thinking` <= 1 short sentence
   - `evaluation_previous_goal` <= 1 short sentence
   - `memory` <= 2 short sentences
   - `next_goal` <= 1 short sentence
   - never restate full DSL/rules in output
{report_scope_rule}
14. Never call file-system tools (`write_file`/`replace_file`/`append_file`/`read_file`) and do not maintain todo/checklist.
15. Module execution is monotonic: once you start module N+1, never go back to module N.
16. For module `行业搜索洞察`, if `params.insight_cycles` is not empty:
   - Iterate cycles in order.
   - For each cycle: run one pass for `current_insight`; if `previous_insight` exists, run one more pass for it.
   - For each pass, explicitly use that pass's `date_range_type` + `insight_date_range` (do not reuse first-cycle values).
   - Keep same field set each pass.
   - Keep same-field results separated by period and write period-split JSON into `module.period_results`
     (labels like `cycle1_current`, `cycle1_previous`).
   - Do not finish task until all labels in `Insight Cycle Concrete Plan` are executed (or marked missing with reason).
17. Completion rule (strict):
   - if report modules exist: all `report_targets` must be completed, and for EACH target all `Report modules` must run once;
   - then all `Global modules` must run once;
   - only then call `done` with a SHORT text.
18. Do not run post-validation/recheck loops after all modules are done.
19. `done` payload rule: do NOT include large nested JSON/data object in `done`. Keep only short text + success flag.
20. After opening each report detail page, immediately run native `extract` once for `计算时间区间` / `计算周期`, and keep it in memory as `current_report_period`.
{report_target_loop_rule}

Execution steps:
A) {execution_a_title}
1. Open 'https://yuntu.oceanengine.com/yuntu_brand/ecom/evaluation_brand/history/distribution?aadvid=1703060939338759'


B) Data extraction/download (Agent executes DSL + tools collect output)
9. If report modules exist:
   - Iterate `report_targets` in listed order.
   - For each report target:
     a) open exact target report in current tab;
     a.1) immediately run native `extract` to capture `计算时间区间` / `计算周期`, then set/update `current_report_period`;
     b) run ALL `Report modules` in order exactly once for this target;
     c) if report modules finished, return to report list and continue next target.
10. After all report targets are done (or if no report modules), run ALL `Global modules` once in order.
11 For global module `行业搜索洞察`, execute by `params.insight_cycles` (current + previous), and keep same-field results separated by period labels in `module.period_results`.
12. If a module has required steps and they fail, mark module as failed/partial, continue next module/target.
13. After finishing all targets + all modules, call `done` with short summary only (no large JSON payload).

Report targets:
{_format_targets(prompt_targets)}

Report modules:
{report_modules}

Global modules:
{global_modules}

{report_plan}

{global_plan}

Insight Cycle Concrete Plan:
{insight_cycle_plan}
"""
