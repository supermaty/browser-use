from __future__ import annotations

import asyncio
import json
import logging
import os
import re
import shutil
import time
from pathlib import Path

from config import DOWNLOADS_PATH
from excel_processor import sum_search_count_from_downloads, sum_search_count_from_excel
from pydantic import BaseModel, ConfigDict
from session import TaskSession
from utils import parse_number

from browser_use import ActionResult, BrowserSession, Tools

from specs import (
    FieldSpec,
    FileSpec,
    InteractionStep,
    ModuleSpec,
    get_module_spec,
    get_profile_spec,
    hydrate_field_spec,
    resolve_profile_name,
)

logger = logging.getLogger(__name__)


def _cap_max_wait_seconds(value: int | None, default: int = 20, upper: int = 20) -> int:
    """Normalize user/model-provided max_wait_seconds into [1, upper]."""
    try:
        if value is None:
            return int(default)
        normalized = int(value)
    except Exception:
        normalized = int(default)
    if normalized < 1:
        normalized = 1
    if normalized > upper:
        normalized = upper
    return normalized


class ExtractFieldsBySpecParams(BaseModel):
    model_config = ConfigDict(extra="forbid")

    profile_name: str = "generic"
    field_keys: list[str] | None = None


class DownloadFilesBySpecParams(BaseModel):
    model_config = ConfigDict(extra="forbid")

    profile_name: str = "generic"
    file_keys: list[str] | None = None
    max_wait_seconds: int | None = None


class OpenReportByNameParams(BaseModel):
    model_config = ConfigDict(extra="forbid")

    profile_name: str = "generic"
    report_name: str
    wait_seconds: int = 5


class NavigateModuleRouteParams(BaseModel):
    model_config = ConfigDict(extra="forbid")

    profile_name: str = "generic"
    module_key: str


class RunModuleCollectionParams(BaseModel):
    model_config = ConfigDict(extra="forbid")

    profile_name: str = "generic"
    module_key: str
    report_name: str | None = None
    touchpoint: str | None = None
    date_type: str | None = None
    date_range: str | None = None
    max_wait_seconds: int | None = None


class OpenYuntuHomeParams(BaseModel):
    model_config = ConfigDict(extra="forbid")

    url: str = "https://yuntu.oceanengine.com"
    wait_seconds: int = 2


class SwitchBrandParams(BaseModel):
    model_config = ConfigDict(extra="forbid")

    profile_name: str = "generic"
    brand_name: str
    brand_aliases: list[str] | None = None


class NavigateReportListParams(BaseModel):
    model_config = ConfigDict(extra="forbid")

    profile_name: str = "generic"


class ClickInContentParams(BaseModel):
    model_config = ConfigDict(extra="forbid")

    profile_name: str = "generic"
    target_text: str
    selectors: list[str] | None = None
    wait_seconds: float = 0.35


class HoverInContentParams(BaseModel):
    model_config = ConfigDict(extra="forbid")

    profile_name: str = "generic"
    target_text: str
    selectors: list[str] | None = None
    wait_seconds: float = 0.6


class SelectInContentParams(BaseModel):
    model_config = ConfigDict(extra="forbid")

    profile_name: str = "generic"
    target_text: str | None = None
    value: str | None = None
    path: list[str] | None = None
    open_target_first: bool = True
    wait_seconds: float = 0.25


DEFAULT_MENU_PATH = ["营销决策", "结案报告"]
DEFAULT_MENU_ALIASES: dict[str, list[str]] = {
    "营销决策": ["营销策略", "营销决策中心"],
    "营销策略": ["营销决策", "营销决策中心"],
    "结案报告": ["投后结案", "结案"],
    "投后结案": ["结案报告", "结案"],
}
DEFAULT_REPORT_LIST_TABS = ["升级版报告"]
DEFAULT_REPORT_LIST_SHORTCUT_LABELS = ["报告列表", "结案报告", "投后结案", "升级版报告"]
DEFAULT_BLOCK_TEXT_TOKENS = ["消息中心", "消息通知", "站内信", "消息", "通知"]
DEFAULT_REPORT_LIKE_TOKENS = ["报告", "流量分析", "人群分析", "转化分析", "营销决策", "投后结案", "结案报告", "画像分析", "5A"]
DEFAULT_CONTENT_CLICK_SELECTORS = [
    "main .ant-tabs-tab",
    "main .ant-menu-item",
    "main .ant-dropdown-menu-item",
    "main .ant-cascader-menu-item",
    "main .ant-select-item-option-content",
    "main .ant-radio-wrapper",
    "main .ant-table-row",
    "main [role='tab']",
    "main [role='button']",
    "main a",
    "main button",
    "main span",
    "main li",
    ".ant-layout-content .ant-tabs-tab",
    ".ant-layout-content .ant-menu-item",
    ".ant-layout-content .ant-dropdown-menu-item",
    ".ant-layout-content .ant-cascader-menu-item",
    ".ant-layout-content [role='tab']",
    ".ant-layout-content [role='button']",
    ".ant-layout-content a",
    ".ant-layout-content button",
    ".ant-layout-content span",
    ".ant-layout-content li",
    "[class*='content'] .ant-tabs-tab",
    "[class*='content'] .ant-menu-item",
    "[class*='content'] [role='tab']",
    "[class*='content'] [role='button']",
    "[class*='content'] a",
    "[class*='content'] button",
    "[class*='content'] span",
    "[class*='content'] li",
]
DEFAULT_CONTENT_HOVER_SELECTORS = [
    "main .evaluation-popper-trigger",
    "main [class*='popper-trigger']",
    "main [class*='tooltip-trigger']",
    "main [class*='popover-trigger']",
    "main .i-icon-help",
    "main [class*='icon-help']",
    "main [class*='question']",
    "main [class*='help']",
    "main [data-popover-id]",
    "main [aria-describedby]",
    ".ant-layout-content .evaluation-popper-trigger",
    ".ant-layout-content [class*='popper-trigger']",
    ".ant-layout-content [class*='tooltip-trigger']",
    ".ant-layout-content [class*='popover-trigger']",
    ".ant-layout-content .i-icon-help",
    ".ant-layout-content [class*='icon-help']",
    ".ant-layout-content [class*='question']",
    ".ant-layout-content [class*='help']",
    ".ant-layout-content [data-popover-id]",
    ".ant-layout-content [aria-describedby]",
    "[class*='content'] .evaluation-popper-trigger",
    "[class*='content'] [class*='popper-trigger']",
    "[class*='content'] [class*='tooltip-trigger']",
    "[class*='content'] [class*='popover-trigger']",
    "[class*='content'] .i-icon-help",
    "[class*='content'] [class*='icon-help']",
    "[class*='content'] [class*='question']",
    "[class*='content'] [class*='help']",
    "[class*='content'] [data-popover-id]",
    "[class*='content'] [aria-describedby]",
]


def _sanitize_hover_selectors(raw_selectors: list[str] | None) -> list[str]:
    """
    Keep only hover-trigger-like selectors from model-provided inputs.
    This prevents overly broad selectors like `main span/div` from causing
    accidental hover on non-tooltip icons (e.g. `+`).
    """
    if not raw_selectors:
        return []
    allow_tokens = (
        "question",
        "help",
        "tooltip",
        "popover",
        "popper",
        "i-icon-help",
        "icon-help",
        "evaluation-popper-trigger",
        "aria-describedby",
        "data-popover-id",
    )
    out: list[str] = []
    seen: set[str] = set()
    for item in raw_selectors:
        text = str(item or "").strip()
        if not text:
            continue
        lower = text.lower()
        if not any(tok in lower for tok in allow_tokens):
            continue
        if lower in seen:
            continue
        seen.add(lower)
        out.append(text)
        if len(out) >= 40:
            break
    return out


def _derive_hover_anchor_text(target_text: str) -> str:
    """
    Convert long benchmark-like field text into a concise nearby anchor label
    for hover targeting.
    Example:
    - "A3流转率行业TOP5%品牌均值(不选触点)" -> "A3流转率"
    """
    raw = str(target_text or "").strip()
    if not raw:
        return raw
    compact = re.sub(r"\s+", "", raw)
    compact = re.sub(r"[（(].*?[）)]", "", compact)
    if not compact:
        return raw

    lower = compact.lower()
    benchmark_like = ("top5" in lower) or ("top 5" in lower) or ("行业" in compact) or ("均值" in compact)
    if not benchmark_like:
        return raw

    metric = re.search(r"(率|占比|比例|人数|人群|金额|次数|规模|roi|ROI)", compact)
    if metric:
        anchor = compact[: metric.end()].strip()
        if len(anchor) >= 2:
            return anchor
    return raw


def _terminal_log(message: str) -> None:
    text = str(message)
    print(text, flush=True)


def _short_text(value: object, limit: int = 120) -> str:
    text = str(value) if value is not None else ""
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) <= limit:
        return text
    return f"{text[:limit - 3]}..."


def _normalize_match_text(value: str | None) -> str:
    text = str(value or "")
    text = re.sub(r"\s+", "", text).lower()
    text = re.sub(r"[\.。,_\-—–·:：/\\|（）()\[\]【】]+", "", text)
    return text


def _infer_touchpoint_from_report_name(report_name: str | None) -> str | None:
    text = (report_name or "").strip()
    if not text:
        return None
    if "竞价" in text:
        return "全部/竞价投放"
    if "星图" in text:
        return "全部/达人营销"
    return None


def _is_text_match_like(target: str | None, candidate: str | None) -> bool:
    target_norm = _normalize_match_text(target)
    candidate_norm = _normalize_match_text(candidate)
    if not target_norm or not candidate_norm:
        return False
    return target_norm in candidate_norm or candidate_norm in target_norm


def _is_select_label_match(target: str | None, candidate: str | None) -> bool:
    """
    Stricter label match for select controls.
    Short labels like "触点" must match exactly (ignoring punctuation/space),
    otherwise controls such as "营销触点" are considered mismatch.
    """
    target_norm = _normalize_match_text(target)
    candidate_norm = _normalize_match_text(candidate)
    if not target_norm or not candidate_norm:
        return False
    if target_norm == candidate_norm:
        return True
    if len(target_norm) <= 4:
        return False
    return target_norm in candidate_norm or candidate_norm in target_norm


def _log_extraction_terminal(
    *,
    profile_name: str,
    requested_keys: list[str],
    status: str,
    page_url: str,
    items: list[dict] | None,
    error: str | None = None,
) -> None:
    rows = items or []
    ok_count = sum(1 for item in rows if item.get("status") == "ok" and item.get("value") is not None)
    _terminal_log(
        f"[extract_fields_by_spec] profile={profile_name} status={status} ok={ok_count}/{len(rows)} "
        f"requested={requested_keys} page={page_url}"
    )
    if error:
        _terminal_log(f"[extract_fields_by_spec] error={error}")
    for item in rows:
        _terminal_log(
            "[extract_fields_by_spec] "
            f"field={item.get('field_key')} status={item.get('status')} source={item.get('source')} "
            f"value={item.get('value')} raw_text={_short_text(item.get('raw_text'))} "
            f"error={_short_text(item.get('error'))}"
        )


def _log_download_terminal(
    *,
    profile_name: str,
    requested_keys: list[str],
    status: str,
    items: list[dict] | None,
    error: str | None = None,
) -> None:
    rows = items or []
    ok_count = sum(1 for item in rows if item.get("status") == "ok")
    _terminal_log(
        f"[download_files_by_spec] profile={profile_name} status={status} ok={ok_count}/{len(rows)} "
        f"requested={requested_keys}"
    )
    if error:
        _terminal_log(f"[download_files_by_spec] error={error}")
    for item in rows:
        _terminal_log(
            "[download_files_by_spec] "
            f"file_key={item.get('file_key')} status={item.get('status')} "
            f"file_name={item.get('file_name')} file_path={item.get('file_path')} "
            f"selector={_short_text(item.get('clicked_selector'))} error={_short_text(item.get('error'))}"
        )


def _profile_menu_path(profile) -> list[str]:
    return list(getattr(profile, "menu_path", []) or DEFAULT_MENU_PATH)


def _profile_menu_aliases(profile) -> dict[str, list[str]]:
    aliases = getattr(profile, "menu_aliases", None)
    if isinstance(aliases, dict) and aliases:
        return aliases
    return DEFAULT_MENU_ALIASES


def _profile_report_list_tabs(profile) -> list[str]:
    return list(getattr(profile, "report_list_tab_candidates", []) or DEFAULT_REPORT_LIST_TABS)


def _profile_report_shortcut_labels(profile) -> list[str]:
    return list(getattr(profile, "report_list_shortcut_labels", []) or DEFAULT_REPORT_LIST_SHORTCUT_LABELS)


def _profile_block_text_tokens(profile) -> list[str]:
    return list(getattr(profile, "blocked_text_tokens", []) or DEFAULT_BLOCK_TEXT_TOKENS)


def _profile_report_like_tokens(profile) -> list[str]:
    return list(getattr(profile, "report_like_tokens", []) or DEFAULT_REPORT_LIKE_TOKENS)


def _json_loads_if_possible(value) -> dict | list:
    if isinstance(value, (dict, list)):
        return value
    if isinstance(value, str):
        try:
            return json.loads(value)
        except Exception:
            return {"raw": value}
    return {"raw": str(value)}


def _benchmark_anchor_from_spec(spec: FieldSpec) -> str:
    """
    Derive benchmark/value-anchor suffix from field label itself.
    Example:
    - "A3流转率行业TOP5%品牌均值(选触点)" -> "行业TOP5%品牌均值"
    """
    source = " ".join(
        [spec.field_key or "", spec.label or "", " ".join(spec.alt_labels or [])]
    ).strip()
    if not source:
        return ""
    text = re.sub(r"\s+", "", source)
    text = re.sub(r"[（(].*?[）)]", "", text)
    if not text:
        return ""
    m = re.search(r"(率|占比|比例|人数|人群|金额|次数|规模|roi|ROI)", text)
    if not m:
        return ""
    anchor = text[m.end() :]
    anchor = re.sub(r"^[：:，,、\-_/\\|]+", "", anchor)
    anchor = anchor.strip()
    return anchor if len(anchor) >= 2 else ""


def _has_benchmark_value_after_anchor(text: str, anchor: str, expects_percent: bool) -> bool:
    if not text or not anchor:
        return False
    compact_text = _normalize_match_text(text)
    compact_anchor = _normalize_match_text(anchor)
    if not compact_text or not compact_anchor:
        return False
    idx = compact_text.find(compact_anchor)
    if idx < 0:
        return False
    tail = compact_text[idx + len(compact_anchor) :]
    if expects_percent:
        return bool(re.search(r"\d+%", tail))
    return bool(re.search(r"\d+", tail))


def _pick_numeric_token_from_text(
    raw_text: str,
    expects_percent: bool,
    rejects_percent: bool,
    label_positions: list[int] | None = None,
) -> str | None:
    tokens: list[tuple[int, str]] = []
    for item in re.finditer(r"-?\d[\d,]*(?:\.\d+)?\s*%?", raw_text):
        token = item.group(0).strip()
        if not token:
            continue
        has_pct = "%" in token
        if expects_percent and not has_pct:
            continue
        if rejects_percent and has_pct:
            continue
        start = item.start()
        end = item.end()
        prev_char = raw_text[start - 1] if start > 0 else ""
        next_char = raw_text[end] if end < len(raw_text) else ""
        # Skip alnum-adjacent fragments such as A3 / v2.
        if re.match(r"[A-Za-z]", prev_char) or re.match(r"[A-Za-z]", next_char):
            continue
        if re.fullmatch(r"\d", re.sub(r"[^\d]", "", token)):
            continue
        score = 0
        if expects_percent and has_pct:
            score += 180
        if rejects_percent and not has_pct:
            score += 160
        digits = re.sub(r"[^\d]", "", token)
        score += min(len(digits) * 12, 160)
        if label_positions:
            nearest = min(abs(start - pos) for pos in label_positions)
            score += max(0, int(120 - nearest * 0.35))
        tokens.append((score, token))
    if not tokens:
        return None
    tokens.sort(key=lambda x: x[0], reverse=True)
    return tokens[0][1]


def _normalize_value(spec: FieldSpec, raw_text: str | None, status: str) -> tuple[float | int | str | None, str]:
    if not raw_text:
        return None, status

    if spec.value_type == "text":
        return raw_text.strip(), "ok"

    text = raw_text.strip()
    combined = " ".join(
        [
            spec.field_key or "",
            spec.label or "",
            " ".join(spec.alt_labels or []),
        ]
    ).lower()
    expects_percent = bool(spec.regex and "%" in spec.regex) or any(token in combined for token in ("率", "占比", "比例", "ratio", "rate"))
    rejects_percent = not expects_percent and any(
        token in combined
        for token in (
            "金额",
            "人数",
            "次数",
            "规模",
            "曝光",
            "消耗",
            "amount",
            "number",
            "count",
            "users",
            "size",
        )
    )

    semantic_text = " ".join([spec.field_key or "", spec.label or "", " ".join(spec.alt_labels or [])]).lower()
    benchmark_anchor = _benchmark_anchor_from_spec(spec)
    benchmark_anchor_norm = _normalize_match_text(benchmark_anchor) if benchmark_anchor else ""
    is_benchmark_field = bool(benchmark_anchor_norm)
    a_stage_match = re.search(r"a\s*([1-5])", semantic_text)
    expected_a_stage = f"a{a_stage_match.group(1)}" if a_stage_match else ""
    is_stage_percent_field = bool(expected_a_stage and expects_percent)
    is_stage_count_field = bool(expected_a_stage and rejects_percent)
    is_new_customer_field = any(token in semantic_text for token in ("拉新", "新客"))
    label_terms = [
        term.strip().lower()
        for term in [spec.field_key, spec.label, *(spec.alt_labels or [])]
        if str(term or "").strip()
    ]
    text_lower = text.lower()
    if is_benchmark_field:
        compact_text = _normalize_match_text(text)
        idx = compact_text.find(benchmark_anchor_norm)
        if idx < 0:
            return None, "parse_failed"
        tail = compact_text[idx + len(benchmark_anchor_norm) :]
        if expects_percent:
            if not re.search(r"\d+%", tail):
                return None, "parse_failed"
        elif not re.search(r"\d+", tail):
            return None, "parse_failed"
    label_positions = [text_lower.find(term) for term in label_terms if term and text_lower.find(term) >= 0]

    if spec.regex:
        matches = list(re.finditer(spec.regex, text))
        if matches:
            picked = None
            for item in matches:
                group_val = item.group(1)
                prefix = text[max(0, item.start() - 8) : item.start()].lower()
                if "top" in prefix and "top5" in text.lower() and "top5" not in spec.label.lower():
                    continue
                picked = group_val
                break
            text = picked if picked is not None else matches[0].group(1)
    else:
        numeric_tokens = list(re.finditer(r"-?\d[\d,]*(?:\.\d+)?\s*%?", text))
        if numeric_tokens:
            candidates: list[tuple[int, str, str]] = []
            for token_match in numeric_tokens:
                token = token_match.group(0).strip()
                has_pct = "%" in token
                if expects_percent and not has_pct:
                    continue
                if rejects_percent and has_pct:
                    continue

                start = token_match.start()
                end = token_match.end()
                prev_char = text[start - 1] if start > 0 else ""
                next_char = text[end] if end < len(text) else ""
                # Drop alnum-adjacent fragments like A3 / v2.
                if re.match(r"[A-Za-z]", prev_char) or re.match(r"[A-Za-z]", next_char):
                    continue

                plain = re.sub(r"[,\s%]", "", token)
                if rejects_percent:
                    if re.fullmatch(r"0\d{3,}", plain):
                        continue
                    if re.fullmatch(r"20\d{2}", plain):
                        continue
                    if len(re.sub(r"[^\d]", "", plain)) <= 1:
                        continue

                score = 0
                score += 120 if (has_pct and expects_percent) else 0
                score -= 420 if (has_pct and rejects_percent) else 0
                score -= 300 if (expects_percent and not has_pct) else 0
                if label_positions:
                    nearest = min(abs(start - pos) for pos in label_positions)
                    score += max(0, int(260 - nearest * 0.9))
                if rejects_percent:
                    digits = re.sub(r"[^\d]", "", plain)
                    score += min(len(digits) * 10, 120)
                    if re.fullmatch(r"\d{4}", digits):
                        score -= 180

                window = text_lower[max(0, start - 28) : min(len(text_lower), end + 36)]
                if is_new_customer_field:
                    if ("拉新" in window) or ("新客" in window):
                        score += 260
                    if ("投后" in window) or ("5a" in window):
                        score += 180
                    else:
                        score -= 120
                    if re.search(r"转粉|粉丝|机会|o机会|a1|a2|a4|a5", window):
                        score -= 520
                    if re.search(r"转化|未匹配|非本次活动", window) and not re.search(r"拉新|新客", window):
                        score -= 900
                if is_stage_percent_field and ("流转" in window):
                    score += 140
                if is_stage_count_field and ("流转" in window or "人群" in window or "人数" in window):
                    score += 120
                if expected_a_stage:
                    if expected_a_stage in window:
                        score += 280
                    elif re.search(r"a[1-5]", window):
                        score -= 760
                    else:
                        score -= 180
                if is_benchmark_field:
                    window_norm = _normalize_match_text(window)
                    if benchmark_anchor_norm and benchmark_anchor_norm in window_norm:
                        score += 320
                    else:
                        score -= 420
                    if is_stage_percent_field and expected_a_stage and expected_a_stage not in window:
                        score -= 420

                candidates.append((score, token, window))

            if candidates:
                candidates.sort(key=lambda x: x[0], reverse=True)
                best_score, best_token, best_window = candidates[0]
                # Keep strict confidence only for highly ambiguous fields; do not over-block stable count fields.
                if (is_benchmark_field or is_new_customer_field) and best_score < 80:
                    return None, "parse_failed"
                if expected_a_stage:
                    if expected_a_stage not in best_window and expected_a_stage not in text_lower:
                        return None, "parse_failed"
                    if re.search(r"a[1-5]", best_window) and expected_a_stage not in best_window:
                        return None, "parse_failed"
                text = best_token
            elif expects_percent or rejects_percent:
                return None, "parse_failed"

    parsed = parse_number(text)
    if parsed is None:
        fallback_token = _pick_numeric_token_from_text(
            raw_text,
            expects_percent=expects_percent,
            rejects_percent=rejects_percent,
            label_positions=label_positions,
        )
        if fallback_token:
            parsed = parse_number(fallback_token)
    if parsed is None:
        return None, "parse_failed"
    return parsed, "ok"


def _maybe_enrich_raw_text(spec: FieldSpec, raw_text: str | None, matched_label: str | None = None) -> str | None:
    if not raw_text:
        return raw_text
    text = raw_text.strip()
    if not text:
        return text

    labels = [matched_label, spec.label, *spec.alt_labels]
    labels = [item.strip() for item in labels if item and item.strip()]
    lowered = text.lower()
    if any(label.lower() in lowered for label in labels):
        return text

    # If the extracted text is mostly a bare numeric token, prepend the label for readability.
    numeric_like = bool(re.fullmatch(r"-?\d[\d,]*(?:\.\d+)?\s*%?", text))
    short_text = len(text) <= 16
    if numeric_like or short_text:
        prefix = labels[0] if labels else (spec.label or spec.field_key)
        return f"{prefix} {text}".strip()

    return text


def _is_hover_like_field(spec: FieldSpec) -> bool:
    if spec.mode == "hover":
        return True
    text = " ".join(
        [
            spec.field_key or "",
            spec.label or "",
            " ".join(spec.alt_labels or []),
            spec.regex or "",
        ]
    ).lower()
    tokens = ("top5", "均值", "问号", "tooltip", "提示", "行业top", "行业 top")
    return any(tok in text for tok in tokens)


def _resolve_field_specs(profile_name: str, field_keys: list[str] | None) -> list[FieldSpec]:
    profile = get_profile_spec(profile_name)
    field_map = {spec.field_key: spec for spec in profile.field_specs}
    alias_map: dict[str, FieldSpec] = {}
    for spec in profile.field_specs:
        keys = [spec.field_key, spec.label, *(spec.alt_labels or [])]
        for key in keys:
            norm = str(key or "").strip()
            if norm:
                alias_map[norm] = spec
    if field_keys:
        selected: list[FieldSpec] = []
        seen: set[str] = set()
        for key in field_keys:
            norm = str(key or "").strip()
            if not norm:
                continue
            spec = field_map.get(norm) or alias_map.get(norm)
            if spec is None:
                continue
            if spec.field_key in seen:
                continue
            seen.add(spec.field_key)
            selected.append(spec)
    else:
        selected = profile.field_specs
    return [hydrate_field_spec(spec) for spec in selected]


def _resolve_file_specs(profile_name: str, file_keys: list[str] | None) -> list[FileSpec]:
    profile = get_profile_spec(profile_name)
    file_map = {spec.file_key: spec for spec in profile.file_specs}
    if not file_keys:
        return []
    return [file_map[key] for key in file_keys if key in file_map]


def _build_file_anchor_requirements(profile) -> dict[str, list[str]]:
    requirements: dict[str, list[str]] = {}
    for module in getattr(profile, "module_specs", []) or []:
        anchors: list[str] = []

        def push_anchor(value: str | None) -> None:
            text = str(value or "").strip()
            if not text:
                return
            if text.startswith(("params.", "module.", "context.")):
                return
            if re.fullmatch(r"-?\d+(?:[\.,]\d+)?%?", text):
                return
            anchors.append(text)

        for item in getattr(module, "route_path", []) or []:
            push_anchor(item)

        for step in getattr(module, "interaction_steps", []) or []:
            op = str(getattr(step, "op", "") or "")
            if op in {"click", "ensure_visible", "select", "select_path", "input", "set_date_range"}:
                push_anchor(getattr(step, "target", None))
                push_anchor(getattr(step, "value", None))
                for v in getattr(step, "path", []) or []:
                    push_anchor(v)
            if op != "download":
                continue
            keys: list[str] = []
            fk = str(getattr(step, "file_key", "") or "").strip()
            if fk:
                keys.append(fk)
            keys.extend([str(v).strip() for v in (getattr(step, "file_keys", []) or []) if str(v).strip()])
            if not keys:
                keys = [str(v).strip() for v in (getattr(module, "file_keys", []) or []) if str(v).strip()]
            if not keys:
                continue
            merged = list(dict.fromkeys([str(module.label or "").strip(), *anchors]))
            merged = [m for m in merged if m]
            for key in keys:
                requirements[key] = merged[-8:]
    return requirements


def _build_file_preconditions(profile) -> dict[str, dict]:
    preconditions: dict[str, dict] = {}
    for module in getattr(profile, "module_specs", []) or []:
        select_pairs: list[tuple[str, str]] = []
        click_targets: list[str] = []

        def _push_click(target: str | None) -> None:
            t = str(target or "").strip()
            if not t:
                return
            click_targets.append(t)

        for step in getattr(module, "interaction_steps", []) or []:
            op = str(getattr(step, "op", "") or "")
            if op in {"click", "ensure_visible"}:
                _push_click(getattr(step, "target", None))
                continue

            if op in {"select", "select_path"}:
                target = str(getattr(step, "target", "") or "").strip()
                value = str(getattr(step, "value", "") or "").strip()
                value_from = str(getattr(step, "value_from", "") or "").strip()
                if target and value and not value_from and not value.startswith(("params.", "module.", "context.")):
                    # Normalize DSL path-like values such as "全部/全部" to terminal value ("全部")
                    # for robust precondition verification.
                    value_parts = _split_select_values(value)
                    normalized_value = value_parts[-1] if value_parts else value
                    select_pairs.append((target, normalized_value))
                path = [str(v).strip() for v in (getattr(step, "path", []) or []) if str(v).strip()]
                if target and path:
                    # For cascader path, enforce terminal value.
                    select_pairs.append((target, path[-1]))
                continue

            if op != "download":
                continue

            keys: list[str] = []
            fk = str(getattr(step, "file_key", "") or "").strip()
            if fk:
                keys.append(fk)
            keys.extend([str(v).strip() for v in (getattr(step, "file_keys", []) or []) if str(v).strip()])
            if not keys:
                keys = [str(v).strip() for v in (getattr(module, "file_keys", []) or []) if str(v).strip()]
            if not keys:
                continue

            dedup_select = list(dict.fromkeys([pair for pair in select_pairs if pair[0] and pair[1]]))
            dedup_click = list(dict.fromkeys([c for c in click_targets if c]))
            record = {
                "module_key": getattr(module, "module_key", None),
                "module_label": getattr(module, "label", None),
                "route_path": list(getattr(module, "route_path", []) or []),
                "select_pairs": [{"target": t, "value": v} for t, v in dedup_select[-6:]],
                "click_targets": dedup_click[-6:],
            }
            for key in keys:
                preconditions[key] = record
    return preconditions


def _resolve_effective_profile_name(session: TaskSession, requested_profile_name: str | None) -> str:
    # Prefer profile inferred from current task intent to avoid LLM hallucinated names
    # (e.g. "brand_ecom") causing navigation drift.
    try:
        if getattr(session, "intent", None) is not None:
            return resolve_profile_name(session.intent)  # type: ignore[arg-type]
    except Exception:
        pass
    requested = str(requested_profile_name or "").strip()
    return requested or "generic"


def _recent_downloaded_files(download_dir: Path) -> list[Path]:
    if not download_dir.exists():
        return []

    files = [path for path in download_dir.glob("*") if path.is_file()]
    files = [path for path in files if path.suffix.lower() not in {".crdownload", ".tmp", ".part"}]
    files.sort(key=lambda path: path.stat().st_mtime, reverse=True)
    return files


def _download_file_meta(download_dir: Path) -> dict[Path, float]:
    meta: dict[Path, float] = {}
    if not download_dir.exists():
        return meta
    for path in download_dir.rglob("*"):
        if not path.is_file():
            continue
        if path.suffix.lower() in {".crdownload", ".tmp", ".part"}:
            continue
        try:
            meta[path] = path.stat().st_mtime
        except OSError:
            continue
    return meta


def _wait_for_new_file(download_dir: Path, before: dict[Path, float], timeout_seconds: int, keywords: list[str]) -> Path | None:
    deadline = time.time() + timeout_seconds
    normalized = [key.lower() for key in keywords]

    while time.time() < deadline:
        current_meta = _download_file_meta(download_dir)
        new_or_updated = [
            path
            for path, mtime in current_meta.items()
            if path not in before or mtime > before.get(path, 0.0) + 0.2
        ]
        new_or_updated.sort(key=lambda p: current_meta.get(p, 0.0), reverse=True)

        if normalized:
            filtered = [path for path in new_or_updated if any(keyword in path.name.lower() for keyword in normalized)]
        else:
            filtered = new_or_updated

        if filtered:
            return filtered[0]

        time.sleep(0.4)

    return None


def _snapshot_download_for_file_key(path: Path, download_dir: Path, file_key: str, name_prefix: str | None = None) -> Path:
    """
    Create a stable per-file_key snapshot to avoid ambiguity when the website exports
    same-name files for different modules.
    """
    try:
        if not path.exists() or not path.is_file():
            return path
        raw_key = str(file_key or "").strip() or "download"
        # Prefer DSL node before [download] as filename prefix.
        # If unavailable, fallback to module-like prefix from file_key.
        module_prefix = re.sub(r"_file(?:_\d+)?$", "", raw_key)
        base = str(name_prefix or "").strip() or module_prefix or raw_key
        # Keep Chinese and general unicode chars; only strip filesystem-illegal chars.
        base = re.sub(r"\s+", "_", base)
        safe_key = re.sub(r"[<>:\"/\\|?*\x00-\x1f]+", "_", base).strip(" .") or "download"
        stamp = time.strftime("%Y%m%d_%H%M%S")
        target = download_dir / f"{safe_key}_{stamp}{path.suffix}"
        counter = 1
        while target.exists():
            target = download_dir / f"{safe_key}_{stamp}_{counter}{path.suffix}"
            counter += 1
        shutil.copy2(path, target)
        return target
    except Exception:
        return path


def _field_semantic_flags(spec: FieldSpec) -> tuple[bool, bool]:
    combined = " ".join(
        [
            spec.field_key or "",
            spec.label or "",
            " ".join(spec.alt_labels or []),
        ]
    ).lower()
    expects_percent = bool(spec.regex and "%" in spec.regex) or any(
        token in combined for token in ("率", "占比", "比例", "ratio", "rate")
    )
    rejects_percent = not expects_percent and any(
        token in combined
        for token in (
            "金额",
            "人数",
            "次数",
            "规模",
            "曝光",
            "消耗",
            "amount",
            "number",
            "count",
            "users",
            "size",
        )
    )
    return expects_percent, rejects_percent


def _expand_label_candidates(spec: FieldSpec) -> list[str]:
    seeds = [spec.label, spec.field_key, *(spec.alt_labels or [])]
    cleaned = [str(v).strip() for v in seeds if str(v or "").strip()]
    if not cleaned:
        return []

    variants: set[str] = set(cleaned)
    replace_rules = [
        ("占比", "比例"),
        ("比例", "占比"),
        ("人数", "人群"),
        ("人群", "人数"),
        ("流转人数", "流转人群"),
        ("流转人群", "流转人数"),
        ("新客", "拉新"),
        ("拉新", "新客"),
    ]
    for label in list(variants):
        stripped = re.sub(r"[（(].*?[）)]", "", label).strip()
        if stripped and stripped != label:
            variants.add(stripped)
        if re.search(r"top\s*5|top5", label, flags=re.IGNORECASE):
            variants.update(
                {
                    "行业TOP5%品牌均值",
                    "行业TOP5品牌均值",
                    "TOP5%品牌均值",
                    "TOP5品牌均值",
                    "行业TOP5%均值",
                }
            )
        for old, new in replace_rules:
            if old in label:
                variants.add(label.replace(old, new))
        # Common sentence prefixes in report narratives.
        for prefix in ("本次活动", "本期活动", "投后", "活动"):
            if not label.startswith(prefix):
                variants.add(f"{prefix}{label}")

    return [v for v in variants if v]


async def _extract_field_by_text_pair_fallback(browser_session: BrowserSession, spec: FieldSpec) -> dict:
    page = await browser_session.must_get_current_page()
    expects_percent, rejects_percent = _field_semantic_flags(spec)
    labels = _expand_label_candidates(spec)
    if not labels:
        return {"status": "not_found", "source": "none", "raw_text": None, "matched_label": None, "error": "label_not_found"}

    js = r"""
(input) => {
  const labels = Array.isArray(input?.labels) ? input.labels : [];
  const expectsPercent = !!input?.expects_percent;
  const rejectsPercent = !!input?.rejects_percent;
  const out = { status: "not_found", source: "none", raw_text: null, matched_label: null, error: "label_not_found" };

  const clean = (v) => String(v ?? "").replace(/\s+/g, " ").trim();
  const softNorm = (v) =>
    clean(v)
      .toLowerCase()
      .replace(/[\.。]/g, "")
      .replace(/[\s_\-—–·:：\/\\|（）()\[\]【】]+/g, "");

  function isUsableScopeRoot(el) {
    if (!el || el === document.body || el === document.documentElement) return false;
    const rect = el.getBoundingClientRect();
    if (!rect || rect.width <= 0 || rect.height <= 0) return false;
    if (rect.width < Math.min(window.innerWidth * 0.45, 520)) return false;
    if (rect.height < 180) return false;
    const cls = clean(el.className || "").toLowerCase();
    if (/header|topbar|navbar|sider|sidebar|menu|toolbar|notice|message/.test(cls)) return false;
    if (el.closest("header,nav,aside,[role='banner']")) return false;
    return true;
  }

  function collectScopedRoots() {
    const roots = [];
    const pushRoot = (el) => {
      if (!isUsableScopeRoot(el)) return;
      if (roots.some((r) => r === el || r.contains(el) || el.contains(r))) return;
      roots.push(el);
    };

    try {
      const last = window.__opt_last_interaction;
      if (
        last &&
        typeof last === "object" &&
        Number.isFinite(last.ts) &&
        Date.now() - Number(last.ts) <= 120000 &&
        Number.isFinite(last.x) &&
        Number.isFinite(last.y) &&
        (last.url || "") === `${location.pathname}${location.search}`
      ) {
        const atPoint = document.elementFromPoint(Number(last.x), Number(last.y));
        if (atPoint) {
          let cur = atPoint;
          for (let i = 0; i < 9 && cur; i++) {
            pushRoot(cur);
            cur = cur.parentElement;
          }
          const tabsRoot = atPoint.closest(".ant-tabs,[role='tablist'],.ant-tabs-nav");
          if (tabsRoot) {
            const activePane = tabsRoot.closest(".ant-tabs")
              ? tabsRoot.closest(".ant-tabs").querySelector(".ant-tabs-tabpane-active,[role='tabpanel']")
              : null;
            if (activePane) pushRoot(activePane);
          }
        }
      }
    } catch (_) {
      // ignore
    }

    const selectors = [
      "[role='tabpanel']",
      ".ant-tabs-tabpane",
      ".ant-tabs-content-holder",
      ".ant-layout-content",
      "main",
      ".panel",
      ".module",
      "section",
      "article",
      "#root",
    ];
    for (const sel of selectors) {
      try {
        const nodes = Array.from(document.querySelectorAll(sel));
        for (const node of nodes) pushRoot(node);
      } catch (_) {}
      if (roots.length >= 4) break;
    }

    if (!roots.length) {
      const fallback = document.querySelector("main,.ant-layout-content,#root");
      if (fallback) roots.push(fallback);
    }
    return roots.slice(0, 4);
  }

  function collectTokenCandidates(text) {
    const source = String(text || "");
    const regex = /-?\d[\d,]*(?:\.\d+)?\s*%?/g;
    const out = [];
    let match;
    while ((match = regex.exec(source)) !== null) {
      const token = clean(match[0]);
      const idx = Number(match.index || 0);
      const hasPct = token.includes("%");
      const plain = token.replace(/[,\s%]/g, "");
      const prev = source[idx - 1] || "";
      const next = source[idx + token.length] || "";

      if (expectsPercent && !hasPct) continue;
      if (rejectsPercent && hasPct) continue;
      // Avoid matching inline alnum fragments like A3 / v2 / x12.
      if (/[A-Za-z]/.test(prev) || /[A-Za-z]/.test(next)) continue;
      // Avoid report/date suffix noise such as "_0224", "2026-02-14", etc.
      if (rejectsPercent) {
        if (/^0\d{3,}$/.test(plain)) continue;
        if (/^20\d{2}$/.test(plain)) continue;
        if (plain.length <= 1) continue;
      }

      let score = 0;
      if (hasPct) score += expectsPercent ? 120 : -480;
      else score += expectsPercent ? -320 : 60;
      if (/^\d{4}$/.test(plain) && rejectsPercent) score -= 220;
      if (rejectsPercent) {
        const digitsOnly = plain.replace(/[^\d]/g, "");
        score += Math.min(digitsOnly.length * 12, 120);
      }
      out.push({ token, score });
    }
    return out.sort((a, b) => b.score - a.score);
  }

  const bodyText = clean(document.body?.innerText || "");
  const scopedRoots = collectScopedRoots();
  const localTexts = scopedRoots
    .map((root) => clean(root?.innerText || root?.textContent || ""))
    .filter((t) => !!t);
  if (!bodyText && !localTexts.length) {
    out.error = "empty_body_text";
    return out;
  }
  const localLinesMerged = localTexts.join("\n");
  const localLines = localLinesMerged.split(/[\r\n]+/).map(clean).filter(Boolean).slice(0, 2600);
  const bodyLines = bodyText.split(/[\r\n]+/).map(clean).filter(Boolean).slice(0, 3500);
  // Scope-first rule:
  // - Scan recent tab/content context first.
  // - Fall back to full page when scoped roots miss the value.
  const linePools = [];
  if (localLines.length) linePools.push(localLines);
  // Keep full-page fallback for cases where scoped root only contains tab headers.
  if (bodyLines.length) linePools.push(bodyLines);

  let best = null;

  for (const label of labels) {
    const labelText = clean(label);
    const labelSoft = softNorm(labelText);
    if (!labelText || !labelSoft) continue;
    for (let poolIndex = 0; poolIndex < linePools.length; poolIndex++) {
      const lines = linePools[poolIndex];
      for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        // Long descriptive lines are common in report insights; keep a wider cap.
        if (!line || line.length > 420) continue;
        const lineSoft = softNorm(line);
        if (!lineSoft.includes(labelSoft)) continue;

        let candidateTail = line;
        const idx = line.indexOf(labelText);
        if (idx >= 0) candidateTail = line.slice(idx + labelText.length);
        // Allow value on same line or the next two lines (common in chart cards).
        const next1 = i + 1 < lines.length ? lines[i + 1] : "";
        const next2 = i + 2 < lines.length ? lines[i + 2] : "";
        const next3 = i + 3 < lines.length ? lines[i + 3] : "";
        const next4 = i + 4 < lines.length ? lines[i + 4] : "";
        const tokenPools = [
          { text: candidateTail, bias: 420 },
          { text: line, bias: 300 },
          { text: next1, bias: 140 },
          { text: next2, bias: 60 },
          { text: next3, bias: 30 },
          { text: next4, bias: 10 },
        ];

        for (const pool of tokenPools) {
          const picks = collectTokenCandidates(pool.text);
          if (!picks.length) continue;
          const picked = picks[0];
          const candidate = {
            token: picked.token,
            label: labelText,
            raw: `${labelText} ${picked.token}`.trim(),
            score: picked.score + pool.bias - poolIndex * 160 - i * 0.05,
          };
          if (/拉新|新客/.test(labelText)) {
            const lc = softNorm(pool.text || "");
            if (/转粉|粉丝|机会|o机会|a1|a2|a4|a5/.test(lc)) candidate.score -= 420;
            if (/拉新|新客/.test(lc)) candidate.score += 220;
          }
          if (!best || candidate.score > best.score) {
            best = candidate;
          }
        }
      }
    }
  }

  if (best) {
    out.status = "ok";
    out.source = "text_pair";
    out.raw_text = best.raw;
    out.matched_label = best.label;
    out.error = null;
    return out;
  }
  out.error = "label_or_value_not_found_in_text";
  return out;
}
"""
    raw = await page.evaluate(
        js,
        {
            "labels": labels,
            "expects_percent": expects_percent,
            "rejects_percent": rejects_percent,
        },
    )
    parsed = _json_loads_if_possible(raw)
    if not isinstance(parsed, dict):
        return {"status": "error", "source": "none", "raw_text": None, "matched_label": None, "error": f"invalid_js_result:{raw}"}
    return parsed

async def _extract_single_field(browser_session: BrowserSession, spec: FieldSpec) -> dict:
    page = await browser_session.must_get_current_page()
    evaluate_spec = spec
    if spec.mode == "auto" and not _is_hover_like_field(spec):
        # Non-hover metrics in auto mode should not trigger hover exploration/click fallback,
        # otherwise extraction failures can accidentally click global UI (e.g. message center).
        evaluate_spec = spec.model_copy(update={"mode": "direct"})
    js = r"""
(input) => {
  const spec = input || {};
  const out = { status: "not_found", source: "none", raw_text: null, error: null };

  const clean = (v) => String(v ?? "").replace(/\s+/g, " ").trim();

  function expandLabels(rawList) {
    const base = Array.from(new Set((rawList || []).map((v) => clean(v)).filter(Boolean)));
    const rules = [
      ["占比", "比例"],
      ["比例", "占比"],
      ["人数", "人群"],
      ["人群", "人数"],
      ["流转人数", "流转人群"],
      ["流转人群", "流转人数"],
      ["新客", "拉新"],
      ["拉新", "新客"],
    ];
    const out = new Set(base);
    for (const label of base) {
      const stripped = clean(label.replace(/[（(].*?[）)]/g, ""));
      if (stripped && stripped !== label) out.add(stripped);
      if (/top\s*5|top5/i.test(label)) {
        out.add("行业TOP5%品牌均值");
        out.add("行业TOP5品牌均值");
        out.add("TOP5%品牌均值");
        out.add("TOP5品牌均值");
        out.add("行业TOP5%均值");
      }
      for (const [oldText, newText] of rules) {
        if (label.includes(oldText)) out.add(label.replace(oldText, newText));
      }
    }
    return Array.from(out);
  }

  function textOf(node) {
    if (!node) return "";
    return clean(node.innerText || node.textContent || "");
  }

  function softNorm(v) {
    return clean(v)
      .toLowerCase()
      .replace(/[\.。]/g, "")
      .replace(/[\s_\-—–·:：\/\\|（）()\[\]【】]+/g, "");
  }

  function semanticKeysFromLabel(labelText) {
    const raw = clean(labelText || "");
    if (!raw) return [];
    const compact = raw.replace(/[\s:：()（）\[\]【】]/g, "");
    const out = new Set();
    if (compact) out.add(compact.toLowerCase());

    let reduced = compact;
    const removeTokens = [
      "行业",
      "品牌",
      "均值",
      "top5%",
      "top5",
      "不选触点",
      "选触点",
      "比例",
      "占比",
      "率",
      "人数",
      "人群",
      "金额",
      "次数",
      "规模",
      "数值",
      "值",
    ];
    for (const tok of removeTokens) {
      reduced = reduced.replace(new RegExp(tok, "ig"), "");
    }
    reduced = clean(reduced);
    if (reduced && reduced.length >= 2) out.add(reduced.toLowerCase());

    const aLike = compact.match(/a\d+/ig) || [];
    for (const item of aLike) out.add(String(item).toLowerCase());

    if (compact.includes("拉新") || compact.includes("新客")) {
      out.add("拉新");
      out.add("新客");
    }
    if (compact.includes("转化")) out.add("转化");

    return Array.from(out).filter((v) => !!v && v.length >= 2);
  }

  function semanticContextScore(contextText, labelText) {
    const ctx = softNorm(contextText || "");
    if (!ctx) return 0;
    const keys = semanticKeysFromLabel(labelText);
    if (!keys.length) return 0;
    let hit = 0;
    for (const key of keys) {
      const nk = softNorm(key);
      if (!nk) continue;
      if (ctx.includes(nk)) hit += 1;
    }
    let score = hit * 180 - (keys.length - hit) * 40;
    if (/非本次活动转化|未匹配到/.test(ctx) && /转化/.test(String(labelText || ""))) score -= 420;
    if (/拉新|新客/.test(String(labelText || ""))) {
      if (/转粉|粉丝|机会|o机会|a1|a2|a4|a5/.test(ctx)) score -= 520;
      if (/拉新|新客/.test(ctx)) score += 220;
    }
    return score;
  }

  function firstText(selectors, rootNode, labelNode) {
    if (!selectors || !selectors.length) return null;
    const scope = rootNode || document;
    const candidates = [];
    const labelRect = labelNode ? labelNode.getBoundingClientRect() : null;
    for (const sel of selectors) {
      try {
        const nodes = Array.from(scope.querySelectorAll(sel));
        for (const node of nodes) {
          if (!isVisible(node)) continue;
          const token = getNumericToken(textOf(node));
          if (!token) continue;
          let score = 2000;
          if (labelRect) {
            const r = node.getBoundingClientRect();
            const distance = Math.hypot(
              r.left + r.width / 2 - (labelRect.left + labelRect.width / 2),
              r.top + r.height / 2 - (labelRect.top + labelRect.height / 2),
            );
            if (distance > 180) continue;
            score -= distance;
          }
          candidates.push({ token, score });
        }
      } catch (_) {}
    }
    if (!candidates.length) return null;
    const picked = candidates.sort((a, b) => b.score - a.score)[0];
    return picked ? picked.token : null;
  }

  const labels = expandLabels([spec.label, spec.field_key, ...(Array.isArray(spec.alt_labels) ? spec.alt_labels : [])]);
  const keyText = clean(`${spec.field_key || ""} ${spec.label || ""} ${(spec.alt_labels || []).join(" ")}`).toLowerCase();
  const benchmarkAnchor = clean(spec.benchmark_anchor || "");
  const benchmarkAnchorSoft = softNorm(benchmarkAnchor);
  const expectsPercent = /%/.test(String(spec.regex || "")) || /率|占比|比例|ratio|rate/.test(keyText);
  const rejectsPercent =
    !expectsPercent &&
    /金额|人数|次数|规模|曝光|消耗|amount|number|count|users|size/.test(keyText);
  const semanticSeed = softNorm(`${spec.field_key || ""} ${spec.label || ""} ${(spec.alt_labels || []).join(" ")}`);
  const expectedStageMatch = semanticSeed.match(/a([1-5])/);
  const expectedStage = expectedStageMatch ? `a${expectedStageMatch[1]}` : "";

  function detectStage(text) {
    const m = softNorm(text || "").match(/a([1-5])/);
    return m ? `a${m[1]}` : "";
  }

  function isVisible(node) {
    if (!node) return false;
    const rect = node.getBoundingClientRect();
    const style = window.getComputedStyle(node);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  }

  function getNumericToken(text) {
    const source = String(text || "");
    const regex = /-?\d[\d,]*(?:\.\d+)?\s*%?/g;
    const picks = [];
    let match;
    while ((match = regex.exec(source)) !== null) {
      const token = clean(match[0]);
      const idx = Number(match.index || 0);
      const prev = source[idx - 1] || "";
      const next = source[idx + token.length] || "";
      const hasPct = token.includes("%");
      const plain = token.replace(/[,\s%]/g, "");

      if (expectsPercent && !hasPct) continue;
      if (rejectsPercent && hasPct) continue;
      // Skip alnum-adjacent fragments like A3 / v2.
      if (/[A-Za-z]/.test(prev) || /[A-Za-z]/.test(next)) continue;
      // Skip common report/date suffix noise for count-like fields.
      if (rejectsPercent) {
        if (/^0\d{3,}$/.test(plain)) continue;
        if (/^20\d{2}$/.test(plain)) continue;
        if (plain.length <= 1) continue;
      }

      let score = 0;
      if (hasPct) score += expectsPercent ? 120 : -420;
      else score += expectsPercent ? -300 : 60;
      if (/^\d{4}$/.test(plain) && rejectsPercent) score -= 180;
      if (rejectsPercent) {
        const digitsOnly = plain.replace(/[^\d]/g, "");
        score += Math.min(digitsOnly.length * 10, 120);
      }
      picks.push({ token, score });
    }
    if (!picks.length) return null;
    picks.sort((a, b) => b.score - a.score);
    return picks[0].token;
  }

  function isUsableScopeRoot(el) {
    if (!el || el === document.body || el === document.documentElement) return false;
    const rect = el.getBoundingClientRect();
    if (!rect || rect.width <= 0 || rect.height <= 0) return false;
    if (rect.width < Math.min(window.innerWidth * 0.45, 520)) return false;
    if (rect.height < 180) return false;
    const cls = clean(el.className || "").toLowerCase();
    if (/header|topbar|navbar|sider|sidebar|menu|toolbar|notice|message/.test(cls)) return false;
    if (el.closest("header,nav,aside,[role='banner']")) return false;
    return true;
  }

  function collectScopedRoots() {
    const roots = [];
    const pushRoot = (el) => {
      if (!isUsableScopeRoot(el)) return;
      if (roots.some((r) => r === el || r.contains(el) || el.contains(r))) return;
      roots.push(el);
    };

    try {
      const last = window.__opt_last_interaction;
      if (
        last &&
        typeof last === "object" &&
        Number.isFinite(last.ts) &&
        Date.now() - Number(last.ts) <= 120000 &&
        Number.isFinite(last.x) &&
        Number.isFinite(last.y) &&
        (last.url || "") === `${location.pathname}${location.search}`
      ) {
        const atPoint = document.elementFromPoint(Number(last.x), Number(last.y));
        if (atPoint) {
          let cur = atPoint;
          for (let i = 0; i < 9 && cur; i++) {
            pushRoot(cur);
            cur = cur.parentElement;
          }
          const tabsRoot = atPoint.closest(".ant-tabs,[role='tablist'],.ant-tabs-nav");
          if (tabsRoot) {
            const activePane = tabsRoot.closest(".ant-tabs")
              ? tabsRoot.closest(".ant-tabs").querySelector(".ant-tabs-tabpane-active,[role='tabpanel']")
              : null;
            if (activePane) pushRoot(activePane);
          }
        }
      }
    } catch (_) {}

    const selectors = [
      "[role='tabpanel']",
      ".ant-tabs-tabpane",
      ".ant-tabs-content-holder",
      ".ant-layout-content",
      "main",
      ".panel",
      ".module",
      "section",
      "article",
      "#root",
    ];
    for (const sel of selectors) {
      try {
        const nodes = Array.from(document.querySelectorAll(sel));
        for (const node of nodes) pushRoot(node);
      } catch (_) {}
      if (roots.length >= 4) break;
    }

    if (!roots.length) {
      const fallback = document.querySelector("main,.ant-layout-content,#root");
      if (fallback) roots.push(fallback);
    }
    return roots.slice(0, 4);
  }

  function findLabelCandidates(labelList, scopeRoot = null, limit = 6) {
    if (!labelList || !labelList.length) return [];
    const nodes = scopeRoot ? Array.from(scopeRoot.querySelectorAll("*")) : Array.from(document.querySelectorAll("body *"));
    const hits = [];
    const seen = new Set();
    for (const node of nodes) {
      const t = textOf(node);
      if (!t || t.length > 360) continue;
      const n = t.toLowerCase();
      const ns = softNorm(t);
      for (const label of labelList) {
        const target = label.toLowerCase();
        const targetSoft = softNorm(label);
        const labelStage = detectStage(label);
        const directHit = n.includes(target) || target.includes(n);
        const softHit = targetSoft && (ns.includes(targetSoft) || targetSoft.includes(ns));
        if (!directHit && !softHit) continue;
        const nodeStage = detectStage(t);
        if (expectedStage) {
          if (labelStage && labelStage !== expectedStage) continue;
          if (nodeStage && nodeStage !== expectedStage) continue;
          if (!nodeStage) {
            const localCtx = textOf(node.closest("div,li,td,th,tr,section,article") || node.parentElement || node);
            const ctxStages = Array.from(softNorm(localCtx).matchAll(/a([1-5])/g)).map((m) => `a${m[1]}`);
            if (ctxStages.length && !ctxStages.includes(expectedStage)) continue;
          }
        }
        let score = -t.length;
        if (n === target || ns === targetSoft) score += 10000;
        else if (n.startsWith(target)) score += 2400;
        else if (softHit) score += 1400;
        else if (target.includes(n)) score += 600;
        else score += 800;
        if (expectedStage && nodeStage === expectedStage) score += 900;
        if (/[：:]$/.test(t)) score += 180;
        if (/金额|人数|比例|占比|率|roi|a\d/i.test(t)) score += 120;
        const key = `${label}@@${t}@@${score}`;
        if (seen.has(key)) continue;
        seen.add(key);
        hits.push({ node, matchedLabel: label, score });
      }
    }
    hits.sort((a, b) => b.score - a.score);
    return hits.slice(0, Math.max(1, limit));
  }

  function findStrictContainer(labelNode, labelText) {
    if (!labelNode) return null;
    const label = clean(labelText || "").toLowerCase();
    const classHint = /(card|item|metric|stat|row|cell|panel|block|module|wrap|content|box|table|list|col|ant-statistic|ant-card|ant-table-row)/i;
    let best = null;
    let bestScore = -1e9;
    let current = labelNode;
    for (let i = 0; i < 8 && current; i++) {
      current = current.parentElement;
      if (!current) break;
      if (!isVisible(current)) continue;
      const t = textOf(current);
      if (!t || t.length > 1200) continue;
      if (label && !t.toLowerCase().includes(label)) continue;
      const token = getNumericToken(t);
      const hasToken = !!token;
      const cls = clean(current.className || "").toLowerCase();
      const tag = (current.tagName || "").toLowerCase();
      const childCount = current.querySelectorAll("*").length;
      let score = -t.length - childCount * 1.2 - i * 40;
      if (classHint.test(cls)) score += 520;
      if (["tr", "td", "li", "section", "article"].includes(tag)) score += 180;
      if (hasToken) score += 360;
      else score -= 180;
      if (score > bestScore) {
        best = current;
        bestScore = score;
      }
    }
    return best;
  }

  function valueNearLabel(labelNode, labelText, container, maxDistance = 260) {
    if (!labelNode || !container) return null;
    const label = clean(labelText || "").toLowerCase();
    const labelRect = labelNode.getBoundingClientRect();
    const labelCx = labelRect.left + labelRect.width / 2;
    const labelCy = labelRect.top + labelRect.height / 2;

    const geometric = [];
    const candidateNodes = Array.from(
      container.querySelectorAll(
        "[class*='value'],[class*='num'],[class*='count'],[data-test*='value'],[data-test*='num'],strong,b,em,.number,.metric-value,span,div,p,td,th"
      )
    );
    for (const node of candidateNodes) {
      if (!isVisible(node)) continue;
      const t = textOf(node);
      if (!t || t.length > 140) continue;
      if (label && t.toLowerCase() === label) continue;
      const token = getNumericToken(t);
      if (!token) continue;
      const rect = node.getBoundingClientRect();
      if (!rect || rect.width <= 0 || rect.height <= 0) continue;
      const cx = rect.left + rect.width / 2;
      const cy = rect.top + rect.height / 2;
      const dx = cx - labelCx;
      const dy = cy - labelCy;
      const distance = Math.hypot(dx, dy);
      if (distance > maxDistance) continue;
      let score = -distance;
      if (Math.abs(dy) < 60) score += 220;
      if (dx > -20) score += 120;
      if (dx < 0 && Math.abs(dx) < 40) score += 80;
      if (token.includes("%")) score += expectsPercent ? 140 : -300;
      else if (expectsPercent) score -= 160;
      if (/top\s*5/i.test(t) && !/top\s*5/i.test(label)) score -= 260;
      if (/非本次活动转化|未匹配到/.test(t) && rejectsPercent) score -= 260;
      const localCtx = textOf(node.closest("div,li,td,th,tr,section,article") || node.parentElement || node);
      score += semanticContextScore(localCtx, labelText);
      geometric.push({ token, score });
    }
    geometric.sort((a, b) => b.score - a.score);
    if (geometric.length && geometric[0].score >= 120) return geometric[0].token;
    return null;
  }

  function findHoverNode(selectors, labelNode, labelText, strictContainer) {
    const isPlusLike = (node) => {
      const t = textOf(node);
      const title = clean(node?.getAttribute?.("title") || "");
      const aria = clean(node?.getAttribute?.("aria-label") || "");
      const cls = clean(node?.className || "");
      const joined = `${t} ${title} ${aria} ${cls}`.toLowerCase();
      if (t === "+" || t === "＋") return true;
      return /plus|add|新增|创建|添加|展开|收起|icon-jia|jiahao|\bi-add\b/.test(joined);
    };
    const hintLikeScore = (node) => {
      const t = textOf(node);
      const title = clean(node?.getAttribute?.("title") || "");
      const aria = clean(node?.getAttribute?.("aria-label") || "");
      const cls = clean(node?.className || "");
      const joined = `${t} ${title} ${aria} ${cls}`.toLowerCase();
      if (t === "?" || t === "？") return 240;
      if (/question|help|tip|tooltip|提示|问号|问/.test(joined)) return 180;
      if (/icon-help|question-circle|i-icon-help|anticon-question/.test(joined)) return 160;
      return -120;
    };
    const nearScore = (node) => {
      if (!labelNode) return 0;
      const nr = node.getBoundingClientRect();
      const lr = labelNode.getBoundingClientRect();
      const nx = nr.left + nr.width / 2;
      const ny = nr.top + nr.height / 2;
      const lx = lr.left + lr.width / 2;
      const ly = lr.top + lr.height / 2;
      return -Math.hypot(nx - lx, ny - ly);
    };
    const findAnchorNode = () => {
      if (!benchmarkAnchorSoft) return null;
      const scope =
        strictContainer ||
        labelNode?.closest("div,li,td,th,tr,section,article,main,[class*='card'],[class*='panel'],[class*='module']") ||
        document.body;
      const nodes = Array.from(scope.querySelectorAll("span,div,p,label,td,th,strong,b,em"));
      let best = null;
      let bestScore = -1e9;
      for (const node of nodes) {
        if (!isVisible(node)) continue;
        const t = textOf(node);
        if (!t || t.length > 180) continue;
        const ns = softNorm(t);
        if (!ns || !ns.includes(benchmarkAnchorSoft)) continue;
        const score = 1200 - t.length + nearScore(node);
        if (score > bestScore) {
          best = node;
          bestScore = score;
        }
      }
      return best;
    };
    const anchorNode = findAnchorNode();
    if (anchorNode) {
      const anchorContainer =
        anchorNode.closest("div,li,td,th,tr,section,article") || anchorNode.parentElement || strictContainer || document.body;
      const preferred = [];
      const anchorNear = (node) => {
        const nr = node.getBoundingClientRect();
        const ar = anchorNode.getBoundingClientRect();
        const nx = nr.left + nr.width / 2;
        const ny = nr.top + nr.height / 2;
        const ax = ar.left + ar.width / 2;
        const ay = ar.top + ar.height / 2;
        return -Math.hypot(nx - ax, ny - ay);
      };
      const anchorSelectors = [
        ".evaluation-popper-trigger",
        "[class*='popper-trigger']",
        "[class*='popover-trigger']",
        "[class*='tooltip-trigger']",
        ".i-icon-help",
        "[class*='icon-help']",
        "[class*='question']",
        "[class*='help']",
        "[data-popover-id]",
        "[aria-describedby]",
        "svg",
        "i",
      ];
      for (const sel of anchorSelectors) {
        try {
          for (const node of Array.from(anchorContainer.querySelectorAll(sel))) {
            if (!isVisible(node)) continue;
            if (isPlusLike(node)) continue;
            const score = 2600 + anchorNear(node) + hintLikeScore(node);
            preferred.push({ node, score });
          }
        } catch (_) {}
      }
      if (preferred.length) {
        preferred.sort((a, b) => b.score - a.score);
        return preferred[0].node;
      }
    }
    const candidates = [];
    if (selectors && selectors.length) {
      for (const sel of selectors) {
        try {
          const nodes = Array.from(document.querySelectorAll(sel));
          for (const node of nodes) {
            if (!isVisible(node)) continue;
            if (isPlusLike(node)) continue;
            const semantic = hintLikeScore(node);
            candidates.push({ node, score: nearScore(node) + 1000 + semantic });
          }
        } catch (_) {}
      }
    }
    if (labelNode) {
      const container = strictContainer || labelNode.closest("div,li,td,th,tr,section,article") || labelNode.parentElement;
      if (container) {
        const nodes = Array.from(
          container.querySelectorAll(
            "[class*='question'],[class*='help'],[class*='tip'],[class*='tooltip'],svg,i,[aria-label*='提示'],[aria-label*='问'],[title],.iconfont"
          )
        );
        for (const node of nodes) {
          if (!isVisible(node)) continue;
          if (isPlusLike(node)) continue;
          const t = textOf(node);
          let score = nearScore(node) + 600 + hintLikeScore(node);
          if (t === "?" || t === "？") score += 120;
          candidates.push({ node, score });
        }
      }
    }
    if (labelText) {
      const token = clean(labelText);
      const global = Array.from(document.querySelectorAll("[title],[aria-label],i,svg,span,div"));
      for (const node of global) {
        if (!isVisible(node)) continue;
        if (isPlusLike(node)) continue;
        const title = clean(node.getAttribute("title") || "");
        const aria = clean(node.getAttribute("aria-label") || "");
        const t = textOf(node);
        const joined = `${title} ${aria} ${t}`.toLowerCase();
        if (!joined) continue;
        if (!(joined.includes("top5") || joined.includes("提示") || joined.includes(token.toLowerCase()) || t === "?" || t === "？")) continue;
        candidates.push({ node, score: nearScore(node) + 100 + hintLikeScore(node) });
      }
    }
    if (!candidates.length) return null;
    candidates.sort((a, b) => b.score - a.score);
    return candidates[0].node;
  }

  function tooltipText(tooltipSelectors, hoverNode) {
    if (tooltipSelectors && tooltipSelectors.length) {
      for (const sel of tooltipSelectors) {
        try {
          const node = document.querySelector(sel);
          const t = textOf(node);
          if (t) return t;
        } catch (_) {}
      }
    }
    if (hoverNode) {
      const title = clean(hoverNode.getAttribute("title") || "");
      if (title) return title;
      const aria = clean(hoverNode.getAttribute("aria-label") || "");
      if (aria) return aria;
    }
    const genericTips = Array.from(document.querySelectorAll("[role='tooltip'],.ant-tooltip-inner,.ant-popover-inner,.tippy-content,.tooltip,.el-tooltip__popper"))
      .map((node) => textOf(node))
      .filter((t) => !!t);
    const hoverPointTips = [];
    if (hoverNode) {
      try {
        const r = hoverNode.getBoundingClientRect();
        const x = r.left + r.width / 2;
        const y = r.top + r.height / 2;
        const stack = document.elementsFromPoint(x, y) || [];
        for (const node of stack) {
          const t = textOf(node);
          if (!t || t.length > 260) continue;
          if (!/-?\d/.test(t) && !/top\s*5|均值|行业|流转/.test(t)) continue;
          hoverPointTips.push(t);
        }
      } catch (_) {}
    }
    const floatingTips = Array.from(document.querySelectorAll("body *"))
      .filter((node) => {
        if (!isVisible(node)) return false;
        const style = window.getComputedStyle(node);
        if (!["fixed", "absolute", "sticky"].includes(style.position)) return false;
        const t = textOf(node);
        if (!t || t.length > 280) return false;
        if (!/-?\d/.test(t) && !/top\s*5|均值|行业|流转/.test(t)) return false;
        return true;
      })
      .map((node) => textOf(node));
    const allTips = genericTips.concat(hoverPointTips).concat(floatingTips);
    const hasBenchmarkValueAfterAnchor = (text) => {
      if (!benchmarkAnchorSoft) return true;
      const ts = softNorm(text || "");
      const idx = ts.indexOf(benchmarkAnchorSoft);
      if (idx < 0) return false;
      const tail = ts.slice(idx + benchmarkAnchorSoft.length);
      return /\d+%?/.test(tail);
    };
    const semanticSeed = softNorm((matchedLabel || spec.label || spec.field_key || ""));
    const stageMatch = semanticSeed.match(/a([1-5])/);
    const expectedStage = stageMatch ? `a${stageMatch[1]}` : "";
    const requiresFlowOnly = !expectedStage && semanticSeed.includes("流转");
    const requiresTop5Strict = semanticSeed.includes("top5");
    const requiresTop5Loose = !requiresTop5Strict && (semanticSeed.includes("均值") || semanticSeed.includes("行业"));
    const filteredTips = allTips.filter((t) => {
      const ts = softNorm(t);
      const hasPct = /-?\d[\d,]*(?:\.\d+)?\s*%/.test(t);
      if (expectsPercent && !hasPct) return false;
      if (rejectsPercent && hasPct) return false;
      if (expectedStage && !new RegExp(expectedStage).test(ts)) return false;
      if (expectedStage) {
        const allStages = Array.from(ts.matchAll(/a([1-5])/g)).map((m) => `a${m[1]}`);
        if (allStages.length && allStages.some((s) => s !== expectedStage)) return false;
      }
      if (requiresFlowOnly && !/流转/.test(ts)) return false;
      if (requiresTop5Strict && !(/top\s*5|top5/.test(ts))) return false;
      if (requiresTop5Loose && !(/top\s*5|top5|行业|均值/.test(ts))) return false;
      if (!hasBenchmarkValueAfterAnchor(t)) return false;
      return true;
    });
    if (expectsPercent && !filteredTips.length) return null;
    const best = filteredTips
      .map((t) => {
        const ts = softNorm(t);
        const hasPercent = /-?\d[\d,]*(?:\.\d+)?\s*%/.test(t);
        return {
          t,
          score:
            (/-?\d/.test(t) ? 1000 : 0) +
            (hasPercent ? 420 : 0) +
            (/top\s*5|行业|均值/i.test(t) ? 160 : 0) +
            (expectedStage && new RegExp(expectedStage).test(ts) ? 260 : 0) +
            (/流转/.test(ts) ? 80 : 0) -
            (expectedStage ? (Array.from(ts.matchAll(/a([1-5])/g)).some((m) => `a${m[1]}` !== expectedStage) ? 520 : 0) : 0) -
            t.length,
        };
      })
      .sort((a, b) => b.score - a.score)[0];
    if (best) return best.t;
    return null;
  }

  const scopedRoots = collectScopedRoots();
  let labelCandidates = [];
  for (const root of scopedRoots) {
    const m = findLabelCandidates(labels, root, 6);
    if (m && m.length) {
      labelCandidates = m;
      break;
    }
  }
  // Scope-first rule:
  // - Prefer recent tab/content context.
  // - Fall back to global label search when scoped roots miss labels.
  if (!labelCandidates.length) labelCandidates = findLabelCandidates(labels, null, 8);
  const primary = labelCandidates[0] || null;
  const labelNode = primary ? primary.node : null;
  const matchedLabel = (primary ? primary.matchedLabel : (spec.label || "")) || "";
  out.matched_label = matchedLabel || null;
  if (!labelNode) out.error = "label_not_found";

  let strictContainer = null;
  let scopedContainer = null;
  let workingContainer = null;
  if (labelNode) {
    strictContainer = findStrictContainer(labelNode, matchedLabel);
    const scopedContainerByLabel = scopedRoots.find((root) => root && root.contains && root.contains(labelNode)) || null;
    if (scopedContainerByLabel && strictContainer && !scopedContainerByLabel.contains(strictContainer)) {
      strictContainer = null;
    }
    scopedContainer = scopedContainerByLabel;
    const looseContainer =
      scopedContainer
        ? scopedContainer
        : (labelNode.closest("div,li,td,th,tr,section,article,main,[class*='card'],[class*='panel'],[class*='module']") || document.body);
    if (!strictContainer) out.error = "strict_container_not_found";
    workingContainer = strictContainer || scopedContainer || looseContainer;
  }

  let directText = null;
  let pickedLabelNode = labelNode;
  let pickedLabel = matchedLabel;
  let pickedStrictContainer = strictContainer;

  if (spec.mode === "direct" || spec.mode === "auto") {
    for (const cand of labelCandidates) {
      const ln = cand.node;
      const ml = cand.matchedLabel || spec.label || "";
      let sc = findStrictContainer(ln, ml);
      const scopedByLabel = scopedRoots.find((root) => root && root.contains && root.contains(ln)) || null;
      if (scopedByLabel && sc && !scopedByLabel.contains(sc)) sc = null;
      const lc = scopedByLabel
        ? scopedByLabel
        : (ln.closest("div,li,td,th,tr,section,article,main,[class*='card'],[class*='panel'],[class*='module']") || document.body);
      const wc = sc || scopedByLabel || lc;
      if (!wc) continue;
      let dt = firstText(spec.value_selectors || [], wc, ln);
      if (!dt) {
        const wideField = /转化|金额|人数|roi/i.test(String(ml || ""));
        const maxDistance = scopedByLabel ? (wideField ? 560 : 420) : (wideField ? 420 : 260);
        dt = valueNearLabel(ln, ml, wc, maxDistance);
      }
      if (dt) {
        directText = dt;
        pickedLabelNode = ln;
        pickedLabel = ml;
        pickedStrictContainer = sc;
        break;
      }
    }
  }

  if (directText && (spec.mode === "direct" || spec.mode === "auto")) {
    out.status = "ok";
    out.source = "direct";
    out.raw_text = directText;
    out.matched_label = pickedLabel || null;
    out.error = null;
    return out;
  }

  if (spec.mode === "direct") {
    if (!out.error) out.error = "value_not_found";
    return out;
  }

  const hoverNode = findHoverNode(spec.hover_selectors || [], pickedLabelNode || labelNode, pickedLabel || matchedLabel, pickedStrictContainer || strictContainer);
  const effectiveHoverNode = hoverNode || pickedLabelNode || labelNode;
  if (!effectiveHoverNode) {
    if (spec.mode === "hover") {
      out.status = "hover_failed";
      out.error = "hover_target_not_found";
    } else {
      out.status = "not_found";
    }
    return out;
  }

  try {
    effectiveHoverNode.scrollIntoView({ block: "center", inline: "center" });
    effectiveHoverNode.dispatchEvent(new MouseEvent("mouseenter", { bubbles: true }));
    effectiveHoverNode.dispatchEvent(new MouseEvent("mouseover", { bubbles: true }));
    effectiveHoverNode.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));
  } catch (e) {
    out.status = "hover_failed";
    out.error = String(e);
    return out;
  }

  return new Promise((resolve) => {
    setTimeout(() => {
      const text = tooltipText(spec.tooltip_selectors || [], effectiveHoverNode);
      if (text) {
        out.status = "ok";
        out.source = "hover";
        out.raw_text = text;
        out.matched_label = matchedLabel || null;
        out.error = null;
      } else {
        out.status = "not_found";
        if (!out.error) out.error = "tooltip_not_found";
      }
      resolve(out);
    }, 900);
  });
}
"""

    evaluate_input = evaluate_spec.model_dump(mode="json")
    evaluate_input["benchmark_anchor"] = _benchmark_anchor_from_spec(spec)
    raw = await page.evaluate(js, evaluate_input)
    parsed = _json_loads_if_possible(raw)
    if not isinstance(parsed, dict):
        parsed = {
            "status": "error",
            "source": "none",
            "raw_text": None,
            "error": f"invalid_js_result:{raw}",
        }

    hover_like_field = _is_hover_like_field(spec)
    if evaluate_spec.mode in {"hover", "auto"} and hover_like_field and parsed.get("status") in {
        "hover_failed",
        "not_found",
        "parse_failed",
    }:
        fallback_hover = await _extract_with_mouse_hover(browser_session, spec)
        if fallback_hover.get("status") == "ok":
            parsed = fallback_hover

    allow_text_pair_fallback = not hover_like_field
    if allow_text_pair_fallback and parsed.get("status") in {"not_found", "hover_failed", "parse_failed"}:
        text_pair = await _extract_field_by_text_pair_fallback(browser_session, spec)
        if text_pair.get("status") == "ok":
            parsed = text_pair

    value, normalized_status = _normalize_value(spec, parsed.get("raw_text"), parsed.get("status", "error"))
    parsed["status"] = "ok" if normalized_status == "ok" else normalized_status
    if parsed["status"] == "ok":
        parsed["error"] = None
    parsed["raw_text"] = _maybe_enrich_raw_text(spec, parsed.get("raw_text"), parsed.get("matched_label"))

    source = parsed.get("source", "none")
    if source not in {"direct", "hover", "auto", "none"}:
        source = "direct" if parsed.get("status") == "ok" else "none"

    return {
        "field_key": spec.field_key,
        "status": parsed.get("status", "error"),
        "source": source,
        "value": value,
        "raw_text": parsed.get("raw_text"),
        "error": parsed.get("error"),
    }


async def _extract_with_mouse_hover(browser_session: BrowserSession, spec: FieldSpec) -> dict:
    page = await browser_session.must_get_current_page()
    expects_percent, rejects_percent = _field_semantic_flags(spec)
    benchmark_anchor = _benchmark_anchor_from_spec(spec)
    js_points = r"""
(input) => {
  const spec = input || {};
  const clean = (v) => String(v ?? "").replace(/\s+/g, " ").trim();
  const softNorm = (v) =>
    clean(v)
      .toLowerCase()
      .replace(/[\.。]/g, "")
      .replace(/[\s_\-—–·:：\/\\|（）()\[\]【】]+/g, "");
  const expandLabels = (rawList) => {
    const base = Array.from(new Set((rawList || []).map(clean).filter(Boolean)));
    const out = new Set(base);
    const rules = [
      ["占比", "比例"],
      ["比例", "占比"],
      ["人数", "人群"],
      ["人群", "人数"],
      ["流转人数", "流转人群"],
      ["流转人群", "流转人数"],
      ["新客", "拉新"],
      ["拉新", "新客"],
    ];
    for (const label of base) {
      const stripped = clean(label.replace(/[（(].*?[）)]/g, ""));
      if (stripped && stripped !== label) out.add(stripped);
      if (/top\s*5|top5/i.test(label)) {
        out.add("行业TOP5%品牌均值");
        out.add("行业TOP5品牌均值");
        out.add("TOP5%品牌均值");
        out.add("TOP5品牌均值");
        out.add("行业TOP5%均值");
      }
      for (const [oldText, newText] of rules) {
        if (label.includes(oldText)) out.add(label.replace(oldText, newText));
      }
    }
    return Array.from(out);
  };
  const labels = expandLabels([spec.field_key, spec.label, ...(Array.isArray(spec.alt_labels) ? spec.alt_labels : [])]);
  const semanticSeed = softNorm(`${spec.field_key || ""} ${spec.label || ""} ${(Array.isArray(spec.alt_labels) ? spec.alt_labels.join(" ") : "")}`);
  const expectedStageMatch = semanticSeed.match(/a([1-5])/);
  const expectedStage = expectedStageMatch ? `a${expectedStageMatch[1]}` : "";
  const benchmarkAnchor = clean(spec.benchmark_anchor || "");
  const benchmarkAnchorSoft = softNorm(benchmarkAnchor);
  const strictBenchmark = !!benchmarkAnchorSoft;
  const detectStage = (text) => {
    const m = softNorm(text || "").match(/a([1-5])/);
    return m ? `a${m[1]}` : "";
  };
  const textOf = (node) => clean(node?.innerText || node?.textContent || "");
  const visible = (el) => {
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  const isPlusLike = (node) => {
    const t = textOf(node);
    const title = clean(node?.getAttribute?.("title") || "");
    const aria = clean(node?.getAttribute?.("aria-label") || "");
    const cls = clean(node?.className || "");
    const joined = `${t} ${title} ${aria} ${cls}`.toLowerCase();
    if (t === "+" || t === "＋") return true;
    return /plus|add|新增|创建|添加|展开|收起|icon-jia|jiahao|\bi-add\b/.test(joined);
  };
  const hintLikeScore = (node) => {
    const t = textOf(node);
    const title = clean(node?.getAttribute?.("title") || "");
    const aria = clean(node?.getAttribute?.("aria-label") || "");
    const cls = clean(node?.className || "");
    const joined = `${t} ${title} ${aria} ${cls}`.toLowerCase();
    if (t === "?" || t === "？") return 220;
    if (/question|help|tip|tooltip|提示|问号|问/.test(joined)) return 160;
    if (/icon-help|question-circle|i-icon-help|anticon-question/.test(joined)) return 140;
    return -140;
  };
  const isHintTrigger = (node) => {
    const t = textOf(node);
    const title = clean(node?.getAttribute?.("title") || "");
    const aria = clean(node?.getAttribute?.("aria-label") || "");
    const cls = clean(node?.className || "");
    const joined = `${t} ${title} ${aria} ${cls}`.toLowerCase();
    if (t === "?" || t === "？") return true;
    return /question|help|tip|tooltip|提示|问号|问|popper-trigger|popover-trigger|tooltip-trigger|i-icon-help|evaluation-popper-trigger/.test(joined);
  };

  const isUsableScopeRoot = (el) => {
    if (!el || el === document.body || el === document.documentElement) return false;
    const rect = el.getBoundingClientRect();
    if (!rect || rect.width <= 0 || rect.height <= 0) return false;
    if (rect.width < Math.min(window.innerWidth * 0.45, 520)) return false;
    if (rect.height < 160) return false;
    const cls = clean(el.className || "").toLowerCase();
    if (/header|topbar|navbar|sider|sidebar|menu|toolbar|notice|message/.test(cls)) return false;
    if (el.closest("header,nav,aside,[role='banner']")) return false;
    return true;
  };

  const collectScopedRoots = () => {
    const roots = [];
    const pushRoot = (el) => {
      if (!isUsableScopeRoot(el)) return;
      if (roots.some((r) => r === el || r.contains(el) || el.contains(r))) return;
      roots.push(el);
    };
    try {
      const last = window.__opt_last_interaction;
      if (
        last &&
        typeof last === "object" &&
        Number.isFinite(last.ts) &&
        Date.now() - Number(last.ts) <= 120000 &&
        Number.isFinite(last.x) &&
        Number.isFinite(last.y) &&
        (last.url || "") === `${location.pathname}${location.search}`
      ) {
        const atPoint = document.elementFromPoint(Number(last.x), Number(last.y));
        if (atPoint) {
          let cur = atPoint;
          for (let i = 0; i < 9 && cur; i++) {
            pushRoot(cur);
            cur = cur.parentElement;
          }
          const tabsRoot = atPoint.closest(".ant-tabs,[role='tablist'],.ant-tabs-nav");
          if (tabsRoot) {
            const activePane = tabsRoot.closest(".ant-tabs")
              ? tabsRoot.closest(".ant-tabs").querySelector(".ant-tabs-tabpane-active,[role='tabpanel']")
              : null;
            if (activePane) pushRoot(activePane);
          }
        }
      }
    } catch (_) {}
    return roots.slice(0, 4);
  };

  const findLabelNode = () => {
    const scopedRoots = collectScopedRoots();
    const nodes = scopedRoots.length
      ? scopedRoots.flatMap((root) => Array.from(root.querySelectorAll("*")))
      : Array.from(document.querySelectorAll("body *"));
    let best = null;
    let bestScore = -1e9;
    for (const node of nodes) {
      const t = textOf(node);
      if (!t || t.length > 220) continue;
      const n = t.toLowerCase();
      const ns = softNorm(t);
      for (const label of labels) {
        const key = label.toLowerCase();
        const keySoft = softNorm(label);
        const labelStage = detectStage(label);
        if (!(n.includes(key) || key.includes(n) || (keySoft && (ns.includes(keySoft) || keySoft.includes(ns))))) continue;
        const nodeStage = detectStage(t);
        if (expectedStage) {
          if (labelStage && labelStage !== expectedStage) continue;
          if (nodeStage && nodeStage !== expectedStage) continue;
          if (!nodeStage) {
            const localCtx = textOf(node.closest("div,li,td,th,tr,section,article") || node.parentElement || node);
            const ctxStages = Array.from(softNorm(localCtx).matchAll(/a([1-5])/g)).map((m) => `a${m[1]}`);
            if (ctxStages.length && !ctxStages.includes(expectedStage)) continue;
          }
        }
        let score = -t.length;
        if (n === key || ns === keySoft) score += 5000;
        else if (n.startsWith(key)) score += 1200;
        else if (keySoft && ns.includes(keySoft)) score += 900;
        if (expectedStage && nodeStage === expectedStage) score += 700;
        if (score > bestScore) {
          best = node;
          bestScore = score;
        }
      }
    }
    return best;
  };

  const labelNode = findLabelNode();
  const labelRect = labelNode ? labelNode.getBoundingClientRect() : null;
  const lx = labelRect ? (labelRect.left + labelRect.width / 2) : 0;
  const ly = labelRect ? (labelRect.top + labelRect.height / 2) : 0;
  const near = (node) => {
    if (!labelRect) return 0;
    const r = node.getBoundingClientRect();
    const x = r.left + r.width / 2;
    const y = r.top + r.height / 2;
    return -Math.hypot(x - lx, y - ly);
  };

  const findAnchorNode = () => {
    if (!benchmarkAnchorSoft) return null;
    const scopedRoots = collectScopedRoots();
    const scopeNodes = scopedRoots.length
      ? scopedRoots.flatMap((root) => Array.from(root.querySelectorAll("*")))
      : Array.from(document.querySelectorAll("body *"));
    let best = null;
    let bestScore = -1e9;
    for (const node of scopeNodes) {
      if (!visible(node)) continue;
      const t = textOf(node);
      if (!t || t.length > 180) continue;
      const ns = softNorm(t);
      if (!ns || !ns.includes(benchmarkAnchorSoft)) continue;
      const score = 1000 - t.length + near(node);
      if (score > bestScore) {
        best = node;
        bestScore = score;
      }
    }
    return best;
  };
  const anchorNode = findAnchorNode();

  const push = (acc, node, baseScore = 0, kind = "hint") => {
    if (!node || !visible(node)) return;
    if (isPlusLike(node)) return;
    const hintScore = hintLikeScore(node);
    const hintTrigger = isHintTrigger(node);
    if (strictBenchmark && kind === "anchor" && !hintTrigger && hintScore < -20) return;
    if (hintScore <= -120 && !hintTrigger) return;
    const r = node.getBoundingClientRect();
    if (!r || r.width <= 0 || r.height <= 0) return;
    acc.push({
      x: Math.round(r.left + r.width / 2),
      y: Math.round(r.top + r.height / 2),
      score: baseScore + near(node) + hintScore + (hintTrigger ? 260 : 0),
      kind,
    });
  };

  const all = [];
  for (const sel of (spec.hover_selectors || [])) {
    try {
      for (const n of Array.from(document.querySelectorAll(sel))) {
        push(all, n, 1000, "hint");
      }
    } catch (_) {}
  }

  if (anchorNode) {
    const anchorContainer =
      anchorNode.closest("div,li,td,th,tr,section,article") || anchorNode.parentElement || document.body;
    const anchorSelectors = [
      ".evaluation-popper-trigger",
      "[class*='popper-trigger']",
      "[class*='popover-trigger']",
      "[class*='tooltip-trigger']",
      ".i-icon-help",
      "[class*='icon-help']",
      "[class*='question']",
      "[class*='help']",
      "[data-popover-id]",
      "[aria-describedby]",
      "[role='button']",
      "button",
    ];
    for (const sel of anchorSelectors) {
      try {
        for (const n of Array.from(anchorContainer.querySelectorAll(sel))) {
          push(all, n, 3200, "anchor");
        }
      } catch (_) {}
    }
  }

  if (labelNode) {
    const container = labelNode.closest("div,li,td,th,tr,section,article") || labelNode.parentElement;
    if (container) {
      const nodes = container.querySelectorAll(
        ".evaluation-popper-trigger,[class*='popper-trigger'],[class*='popover-trigger'],[class*='tooltip-trigger'],[class*='question'],[class*='help'],[class*='tip'],[class*='tooltip'],.i-icon-help,[class*='icon-help'],[data-popover-id],[aria-describedby],[role='button'],button,[title],[aria-label]"
      );
      for (const n of Array.from(nodes)) {
        const t = textOf(n);
        if (t && t.length > 8) continue;
        push(all, n, 700, "hint");
      }
    }
    if (!strictBenchmark) {
      // Fallback: hover the label itself and nearby area on the right side.
      const lr = labelNode.getBoundingClientRect();
      if (lr && lr.width > 0 && lr.height > 0) {
        all.push({ x: Math.round(lr.left + lr.width / 2), y: Math.round(lr.top + lr.height / 2), score: 180, kind: "fallback" });
        all.push({ x: Math.round(lr.right + 10), y: Math.round(lr.top + lr.height / 2), score: 120, kind: "fallback" });
        all.push({ x: Math.round(lr.right + 18), y: Math.round(lr.top + lr.height / 2), score: 80, kind: "fallback" });
      }
    }
  }

  if (!strictBenchmark) {
    for (const n of Array.from(document.querySelectorAll("[title],[aria-label],i,svg,span"))) {
      if (!visible(n)) continue;
      if (isPlusLike(n)) continue;
      const t = textOf(n);
      const title = clean(n.getAttribute("title") || "");
      const aria = clean(n.getAttribute("aria-label") || "");
      const joined = `${t} ${title} ${aria}`.toLowerCase();
      if (!joined) continue;
      const labelHit = labels.some((label) => {
        const ls = softNorm(label);
        const js = softNorm(joined);
        return !!ls && (js.includes(ls) || ls.includes(js));
      });
      if (!(joined.includes("top5") || joined.includes("提示") || labelHit || t === "?" || t === "？")) continue;
      push(all, n, 100, "hint");
    }
  }

  const anchorOnly = all.filter((p) => p && p.kind === "anchor");
  const hasHint = all.some((p) => p && p.kind === "hint" && Number(p.score || 0) > -120);
  const ranked = (
    anchorOnly.length
      ? anchorOnly
      : (hasHint ? all.filter((p) => p.kind === "hint" || p.kind === "anchor") : all)
  ).sort((a, b) => b.score - a.score);
  const unique = [];
  const seen = new Set();
  for (const p of ranked) {
    const key = `${Math.round(p.x / 4)}:${Math.round(p.y / 4)}`;
    if (seen.has(key)) continue;
    seen.add(key);
    unique.push({ x: p.x, y: p.y });
    if (unique.length >= (strictBenchmark ? 6 : 10)) break;
  }
  return unique;
}
"""
    hover_input = spec.model_dump(mode="json")
    hover_input["benchmark_anchor"] = benchmark_anchor
    points_raw = await page.evaluate(js_points, hover_input)
    points = _json_loads_if_possible(points_raw)
    if not isinstance(points, list) or not points:
        return {"status": "hover_failed", "source": "none", "raw_text": None, "error": "hover_target_not_found"}

    js_tip = r"""
(input) => {
  const selectors = Array.isArray(input?.tooltip_selectors) ? input.tooltip_selectors : [];
  const expectsPercent = !!input?.expects_percent;
  const rejectsPercent = !!input?.rejects_percent;
  const pointX = Number.isFinite(input?.point_x) ? Number(input.point_x) : null;
  const pointY = Number.isFinite(input?.point_y) ? Number(input.point_y) : null;
  const clean = (v) => String(v ?? "").replace(/\s+/g, " ").trim();
  const softNorm = (v) =>
    clean(v)
      .toLowerCase()
      .replace(/[\.。]/g, "")
      .replace(/[\s_\-—–·:：\/\\|（）()\[\]【】]+/g, "");
  const visible = (el) => {
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  const textOf = (n) => clean(n?.innerText || n?.textContent || "");
  const labels = [input?.field_key, input?.label, ...(Array.isArray(input?.alt_labels) ? input.alt_labels : [])]
    .map((v) => clean(v))
    .filter(Boolean);
  const semanticKeys = (() => {
    const out = new Set();
    for (const label of labels) {
      const compact = softNorm(label);
      if (!compact) continue;
      out.add(compact);
      const stages = compact.match(/a[1-5]/g) || [];
      for (const s of stages) out.add(s);
      if (compact.includes("流转")) out.add("流转");
      if (compact.includes("行业")) out.add("行业");
      if (compact.includes("均值")) out.add("均值");
      if (compact.includes("top5")) out.add("top5");
    }
    return Array.from(out).filter((v) => v.length >= 2);
  })();
  const semanticSeed = semanticKeys.join(" ");
  const stageMatch = semanticSeed.match(/a([1-5])/);
  const expectedStage = stageMatch ? `a${stageMatch[1]}` : "";
  const requiresFlowOnly = !expectedStage && /流转/.test(semanticSeed);
  const requiresTop5Strict = /top5/.test(semanticSeed);
  const requiresTop5Loose = !requiresTop5Strict && /行业|均值/.test(semanticSeed);

  const items = [];
  for (const sel of selectors) {
    try {
      for (const node of Array.from(document.querySelectorAll(sel))) {
        if (!visible(node)) continue;
        const t = textOf(node);
        if (!t) continue;
        items.push(t);
      }
    } catch (_) {}
  }
  for (const node of Array.from(document.querySelectorAll("[role='tooltip'],.ant-tooltip-inner,.ant-popover-inner,.tippy-content,.tooltip,.el-tooltip__popper"))) {
    if (!visible(node)) continue;
    const t = textOf(node);
    if (!t) continue;
    items.push(t);
  }
  if (Number.isFinite(pointX) && Number.isFinite(pointY)) {
    try {
      const stack = document.elementsFromPoint(pointX, pointY) || [];
      for (const node of stack) {
        if (!visible(node)) continue;
        const t = textOf(node);
        if (!t || t.length > 260) continue;
        if (!/-?\d/.test(t) && !/top\s*5|均值|行业|流转/.test(t)) continue;
        items.push(t);
      }
    } catch (_) {}
  }
  // Floating overlays not covered by common tooltip selectors.
  for (const node of Array.from(document.querySelectorAll("body *"))) {
    if (!visible(node)) continue;
    const style = window.getComputedStyle(node);
    if (!["fixed", "absolute", "sticky"].includes(style.position)) continue;
    const t = textOf(node);
    if (!t || t.length > 280) continue;
    if (!/-?\d/.test(t) && !/top\s*5|均值|行业|流转/.test(t)) continue;
    const rect = node.getBoundingClientRect();
    if (!rect || rect.width < 30 || rect.height < 14) continue;
    items.push(t);
  }
  const filtered = items.filter((t) => {
    const ts = softNorm(t);
    const hasPercent = /-?\d[\d,]*(?:\.\d+)?\s*%/.test(t);
    if (expectsPercent && !hasPercent) return false;
    if (rejectsPercent && hasPercent) return false;
    if (expectedStage && !new RegExp(expectedStage).test(ts)) return false;
    if (expectedStage) {
      const allStages = Array.from(ts.matchAll(/a([1-5])/g)).map((m) => `a${m[1]}`);
      if (allStages.length && allStages.some((s) => s !== expectedStage)) return false;
    }
    if (requiresFlowOnly && !/流转/.test(ts)) return false;
    if (requiresTop5Strict && !/top\s*5|top5/.test(ts)) return false;
    if (requiresTop5Loose && !/top\s*5|top5|行业|均值/.test(ts)) return false;
    if (semanticKeys.length) {
      const hit = semanticKeys.some((k) => ts.includes(k));
      if (!hit && !!expectedStage) return false;
    }
    return true;
  });
  if (!filtered.length) return null;
  const best = filtered
    .map((t) => {
      const hasPercent = /-?\d[\d,]*(?:\.\d+)?\s*%/.test(t);
      return {
        t,
        score:
          (/-?\d/.test(t) ? 1200 : 0) +
          (hasPercent ? 420 : 0) +
          (/top\s*5/i.test(t) ? 220 : 0) +
          (/行业|均值|流转/i.test(t) ? 180 : 0) +
          (expectedStage ? (new RegExp(expectedStage).test(softNorm(t)) ? 180 : 0) : 0) -
          (/转粉|粉丝|机会|o机会|a4|a5/i.test(t) ? 520 : 0) -
          (expectedStage ? (Array.from(softNorm(t).matchAll(/a([1-5])/g)).some((m) => `a${m[1]}` !== expectedStage) ? 520 : 0) : 0) -
          t.length,
      };
    })
    .sort((a, b) => b.score - a.score)[0];
  return best ? best.t : null;
}
"""

    mouse = await page.mouse
    probe_dwells = (0.45, 0.85, 1.25) if benchmark_anchor else (0.28, 0.55, 0.9)
    max_points = 6 if benchmark_anchor else len(points)
    for idx, point in enumerate(points):
        if idx >= max_points:
            break
        x = int(point.get("x", 0))
        y = int(point.get("y", 0))
        if x <= 0 or y <= 0:
            continue
        try:
            await mouse.move(x=x, y=y, steps=3)
            # Some popovers are delayed; probe the same point with increasing dwell time.
            for dwell in probe_dwells:
                await asyncio.sleep(dwell)
                tip_raw = await page.evaluate(
                    js_tip,
                    {
                        **spec.model_dump(mode="json"),
                        "expects_percent": expects_percent,
                        "rejects_percent": rejects_percent,
                        "point_x": x,
                        "point_y": y,
                    },
                )
                tip_text = tip_raw if isinstance(tip_raw, str) else str(tip_raw) if tip_raw else ""
                if not tip_text:
                    # Keep hover stable for benchmark tooltip fields; tiny jitter only for generic fields.
                    if not benchmark_anchor:
                        await mouse.move(x=x + 1, y=y + 1, steps=1)
                    continue
                tip_norm = re.sub(r"\s+", " ", tip_text).strip().lower()
                has_number = bool(re.search(r"-?\d[\d,]*(?:\.\d+)?\s*%?", tip_text))
                if expects_percent and not re.search(r"-?\d[\d,]*(?:\.\d+)?\s*%", tip_text):
                    continue
                # For TOP5-like fields, require stronger semantic hints in hovered text.
                label_seed = " ".join(
                    [
                        spec.field_key or "",
                        spec.label or "",
                        " ".join(spec.alt_labels or []),
                    ]
                ).lower()
                is_top5_like = ("top5" in label_seed) or ("top 5" in label_seed) or ("行业top" in label_seed)
                stage_match = re.search(r"a\s*([1-5])", label_seed)
                expected_stage = f"a{stage_match.group(1)}" if stage_match else ""
                is_flow_like = ("流转" in label_seed) and not expected_stage
                if is_top5_like and not re.search(r"top\s*5|top5|行业|均值", tip_norm):
                    continue
                if expected_stage and not re.search(re.escape(expected_stage), tip_norm):
                    continue
                if expected_stage and re.search(r"a[1-5]", tip_norm):
                    stages = set(re.findall(r"a[1-5]", tip_norm))
                    if expected_stage not in stages or any(s != expected_stage for s in stages):
                        continue
                if is_flow_like and not re.search(r"流转", tip_norm):
                    continue
                if benchmark_anchor and not _has_benchmark_value_after_anchor(tip_text, benchmark_anchor, expects_percent):
                    # This is usually the card summary text, not the benchmark tooltip value.
                    # Keep searching other hover candidates.
                    continue
                if not has_number:
                    continue
                return {"status": "ok", "source": "hover", "raw_text": tip_text, "error": None}
        except Exception:
            continue

    return {"status": "hover_failed", "source": "none", "raw_text": None, "error": "hover_target_not_found"}


async def _click_download_trigger(
    browser_session: BrowserSession,
    spec: FileSpec,
    download_context: dict | None = None,
) -> dict:
    page = await browser_session.must_get_current_page()
    js = r"""
(input) => {
  const selectors = Array.isArray(input?.selectors) ? input.selectors : [];
  const expectedKeywords = Array.isArray(input?.expected_keywords) ? input.expected_keywords.map((x) => String(x || "").toLowerCase()) : [];
  const preferNearLastInteraction = input?.prefer_near_last_interaction !== false;
  const lastInteractionMaxDistance = Number.isFinite(input?.last_interaction_max_distance) ? Number(input.last_interaction_max_distance) : 460;
  const contextHints = Array.isArray(input?.context_hints) ? input.context_hints.map((x) => String(x || "").toLowerCase()) : [];
  const hardContextHints = Array.isArray(input?.hard_context_hints) ? input.hard_context_hints.map((x) => String(x || "").toLowerCase()) : [];
  const anchorHints = Array.isArray(input?.anchor_hints) ? input.anchor_hints.map((x) => String(x || "").toLowerCase()) : [];
  const anchorMaxDistance = Number.isFinite(input?.anchor_max_distance) ? Number(input.anchor_max_distance) : (anchorHints.length ? 260 : 0);
  const minContextHits = Number.isFinite(input?.min_context_hits) ? Number(input.min_context_hits) : (contextHints.length >= 4 ? 2 : contextHints.length ? 1 : 0);
  const minHardContextHits = Number.isFinite(input?.min_hard_context_hits) ? Number(input.min_hard_context_hits) : (hardContextHints.length >= 3 ? 2 : hardContextHints.length ? 1 : 0);
  const minAnchorHits = Number.isFinite(input?.min_anchor_hits) ? Number(input.min_anchor_hits) : (anchorHints.length ? 1 : 0);
  const contextExcludeHints = Array.isArray(input?.context_exclude_hints) ? input.context_exclude_hints.map((x) => String(x || "").toLowerCase()) : [];
  const out = { clicked: false, selector: null, error: null, picked_text: null };

  const clean = (v) => String(v ?? "").replace(/\s+/g, " ").trim();
  const norm = (v) => clean(v).toLowerCase();

  const isVisible = (el) => {
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  const isLikelyGlobalNav = (node) => {
    if (!node) return false;
    if (node.closest("header,aside,nav,[role='banner']")) return true;
    let cur = node;
    for (let depth = 0; depth < 5 && cur; depth++) {
      const cls = clean(cur.className || "").toLowerCase();
      if (cls && /(header|topbar|navbar|toolbar|message|notice|user|avatar)/.test(cls)) {
        const r = cur.getBoundingClientRect();
        if (r && r.top < 260) return true;
      }
      cur = cur.parentElement;
    }
    return false;
  };
  const distanceToLast = (node, last) => {
    try {
      const r = node.getBoundingClientRect();
      const cx = r.left + r.width / 2;
      const cy = r.top + r.height / 2;
      const dx = cx - Number(last.x || 0);
      const dy = cy - Number(last.y || 0);
      return Math.sqrt(dx * dx + dy * dy);
    } catch (_) {
      return 1e9;
    }
  };
  const last = window.__opt_last_interaction || null;
  const canUseLast =
    !!last &&
    Date.now() - Number(last.ts || 0) <= 20000 &&
    String(last.url || "") === `${location.pathname}${location.search}` &&
    Number.isFinite(Number(last.x)) &&
    Number.isFinite(Number(last.y));

  const contextText = (el) => {
    const root =
      el?.closest(
        "[role='tabpanel'],.ant-tabs-tabpane,.ant-card,.ant-table-wrapper,.ant-table,.panel,.module,section,article,.ant-layout-content,main"
      ) || el?.closest("div,li,td,th,tr") || document.body;
    // Keep context local; avoid parent-wide text pollution that can mix multiple modules.
    const text = `${root?.innerText || root?.textContent || ""}`;
    return norm(text).slice(0, 1200);
  };

  const pageText = norm(document.body?.innerText || "").slice(0, 3000);

  const isNodeAllowedForFile = (el) => {
    const localCtx = contextText(el);
    const ctx = `${localCtx} ${pageText}`;
    if (!ctx) return true;
    if (contextExcludeHints.length) {
      const bad = contextExcludeHints.some((kw) => kw && ctx.includes(kw));
      if (bad) return false;
    }
    if (contextHints.length) {
      const hitCount = contextHints.filter((kw) => kw && localCtx.includes(kw)).length;
      if (hitCount < Math.max(0, minContextHits)) return false;
    }
    if (hardContextHints.length) {
      const hardHitCount = hardContextHints.filter((kw) => kw && localCtx.includes(kw)).length;
      if (hardHitCount < Math.max(0, minHardContextHits)) return false;
    }
    if (anchorHints.length) {
      const anchorHitCount = anchorHints.filter((kw) => kw && localCtx.includes(kw)).length;
      if (anchorHitCount < Math.max(0, minAnchorHits)) return false;
    }

    if (expectedKeywords.length) {
      const localHit = expectedKeywords.some((kw) => kw && localCtx.includes(kw));
      if (!localHit) return false;
    }
    return true;
  };

  const nodeCenter = (node) => {
    const r = node.getBoundingClientRect();
    return { x: r.left + r.width / 2, y: r.top + r.height / 2 };
  };
  const lastInteraction = (() => {
    const raw = window.__opt_last_interaction;
    if (!raw || typeof raw !== "object") return null;
    const now = Date.now();
    if (!Number.isFinite(raw.ts) || now - Number(raw.ts) > 180000) return null;
    const currentUrl = `${location.pathname}${location.search}`;
    if ((raw.url || "") !== currentUrl) return null;
    if (!Number.isFinite(raw.x) || !Number.isFinite(raw.y)) return null;
    return { x: Number(raw.x), y: Number(raw.y), text: String(raw.text || "").toLowerCase() };
  })();

  const isUsableScopeRoot = (el) => {
    if (!el || el === document.body || el === document.documentElement) return false;
    const rect = el.getBoundingClientRect();
    if (!rect || rect.width <= 0 || rect.height <= 0) return false;
    if (rect.width < Math.min(window.innerWidth * 0.45, 520)) return false;
    if (rect.height < 180) return false;
    const cls = clean(el.className || "").toLowerCase();
    if (/header|topbar|navbar|sider|sidebar|menu|toolbar|notice|message/.test(cls)) return false;
    if (el.closest("header,nav,aside,[role='banner']")) return false;
    return true;
  };

  const lastScopeRoot = (() => {
    if (!lastInteraction) return null;
    try {
      const atPoint = document.elementFromPoint(lastInteraction.x, lastInteraction.y);
      if (!atPoint) return null;
      let cur = atPoint;
      let best = null;
      for (let i = 0; i < 10 && cur; i++) {
        if (isUsableScopeRoot(cur)) {
          const hasDownloadLike = !!cur.querySelector(
            "[class*='download'],[class*='export'],button[title*='下载'],button[aria-label*='下载'],button[title*='导出'],button[aria-label*='导出'],[data-test*='下载'],[data-testid*='下载'],[data-test*='导出'],[data-testid*='导出']"
          );
          if (hasDownloadLike) {
            best = cur;
            break;
          }
          if (!best) best = cur;
        }
        cur = cur.parentElement;
      }
      return best;
    } catch (_) {
      return null;
    }
  })();

  const candidates = [];
  const seenNodes = new Set();
  const anchorPoints = (() => {
    if (!anchorHints.length) return [];
    const out = [];
    for (const node of Array.from(document.querySelectorAll("a,button,span,div,li,[role='tab'],[role='button'],[role='menuitem'],h1,h2,h3,h4,p"))) {
      if (!isVisible(node)) continue;
      const t = clean(node.innerText || node.textContent || "");
      if (!t || t.length > 120) continue;
      const n = norm(t);
      const hitCount = anchorHints.filter((kw) => kw && n.includes(kw)).length;
      if (hitCount <= 0) continue;
      const c = nodeCenter(node);
      out.push({ x: c.x, y: c.y, hitCount, text: t });
    }
    out.sort((a, b) => b.hitCount - a.hitCount);
    return out.slice(0, 18);
  })();
  const pushNode = (el, sourceSelector) => {
    if (!el || seenNodes.has(el)) return;
    seenNodes.add(el);
    candidates.push({ el, sourceSelector });
  };

  for (const sel of selectors) {
    try {
      for (const node of Array.from(document.querySelectorAll(sel))) pushNode(node, sel);
    } catch (_) {}
  }
  for (const node of Array.from(document.querySelectorAll("button,a,[role='button'],li,span,div"))) {
    pushNode(node, "fallback:text(download/export)");
  }

  const ranked = candidates
    .map((el) => {
      const node = el.el;
      const t = clean(node.innerText || node.textContent || "");
      const title = clean(node.getAttribute("title") || "");
      const aria = clean(node.getAttribute("aria-label") || "");
      const cls = clean(node.className || "");
      const testid = clean(node.getAttribute("data-testid") || node.getAttribute("data-test") || "");
      const joined = `${t} ${title} ${aria} ${cls} ${testid}`.toLowerCase();
      if (!isVisible(node)) return null;
      if (!joined) return null;
      if (!/下载|导出|download|export|xiazai|daochu/.test(joined)) return null;
      if ((t || "").length > 80) return null;
      if (!isNodeAllowedForFile(node)) return null;
      if (lastScopeRoot && !lastScopeRoot.contains(node)) return null;

      let score = 0;
      if (el.sourceSelector !== "fallback:text(download/export)") score += 1600;
      if (/下载|导出/.test(t)) score += 300;
      if (/download|export/.test(joined)) score += 120;
      const ctx = contextText(node);
      if (contextHints.length) {
        const hitCount = contextHints.filter((kw) => kw && ctx.includes(kw)).length;
        score += hitCount * 240;
      }
      if (hardContextHints.length) {
        const hardHitCount = hardContextHints.filter((kw) => kw && ctx.includes(kw)).length;
        score += hardHitCount * 420;
      }
      if (contextExcludeHints.length) {
        const badCount = contextExcludeHints.filter((kw) => kw && ctx.includes(kw)).length;
        score -= badCount * 1200;
      }

      if (preferNearLastInteraction && lastInteraction) {
        const c = nodeCenter(node);
        const dLast = Math.hypot(c.x - lastInteraction.x, c.y - lastInteraction.y);
        score += Math.max(0, 2600 - dLast * 4.2);
        if (lastInteractionMaxDistance != null && dLast > lastInteractionMaxDistance) {
          // Do not hard-reject: keep as fallback with strong penalty.
          score -= Math.min(2200, (dLast - lastInteractionMaxDistance) * 2.4);
        }
      }
      if (lastScopeRoot) {
        // Stronger bias to keep click inside the current working scope.
        if (lastScopeRoot.contains(node)) score += 1800;
        else score -= 1400;
      }
      if (anchorHints.length) {
        if (!anchorPoints.length) return null;
        const c = nodeCenter(node);
        let minAnchorDist = Number.POSITIVE_INFINITY;
        let hitBoost = 0;
        for (const ap of anchorPoints) {
          const d = Math.hypot(c.x - ap.x, c.y - ap.y);
          if (d < minAnchorDist) minAnchorDist = d;
          hitBoost = Math.max(hitBoost, Number(ap.hitCount || 0));
        }
        if (Number.isFinite(anchorMaxDistance) && anchorMaxDistance > 0 && minAnchorDist > anchorMaxDistance) {
          // Prefer near anchors but allow fallback candidates in same module.
          score -= Math.min(1800, (minAnchorDist - anchorMaxDistance) * 2.6);
        }
        score += Math.max(0, 2200 - minAnchorDist * 6.2);
        score += hitBoost * 260;
      }

      return {
        node,
        sourceSelector: el.sourceSelector,
        score,
        t,
      };
    })
    .filter(Boolean)
    .sort((a, b) => b.score - a.score);

  if (ranked.length) {
    const clickSameTab = (node) => {
      const clickable =
        node?.closest("a,button,[role='button'],li,.ant-dropdown-menu-item,.ant-btn,div,span") || node;
      if (!clickable) return false;
      const anchor =
        clickable.tagName === "A" ? clickable : clickable.closest("a[href]");
      if (anchor) {
        try {
          anchor.setAttribute("target", "_self");
          anchor.removeAttribute("rel");
        } catch (_) {}
        try {
          anchor.dispatchEvent(
            new MouseEvent("click", {
              bubbles: true,
              cancelable: true,
              button: 0,
              ctrlKey: false,
              metaKey: false,
              shiftKey: false,
              altKey: false,
            })
          );
          return true;
        } catch (_) {}
      }
      clickable.click();
      return true;
    };
    try {
      const picked = ranked[0].node;
      const clickTarget =
        picked.closest("button,a,[role='button'],li,.ant-dropdown-menu-item,.ant-btn,div,span") || picked;
      clickTarget.scrollIntoView({ block: "center", inline: "center" });
      clickSameTab(clickTarget);
      out.clicked = true;
      out.selector = ranked[0].sourceSelector;
      out.picked_text = ranked[0].t || null;
      return out;
    } catch (e) {
      out.error = String(e);
      return out;
    }
  }

  out.error = "download_trigger_not_found";
  return out;
}
"""
    raw = await page.evaluate(
        js,
        {
            "selectors": spec.trigger_selectors,
            "expected_keywords": spec.expected_name_keywords,
            "prefer_near_last_interaction": True,
            "last_interaction_max_distance": (download_context or {}).get("last_interaction_max_distance", 460),
            "context_hints": (download_context or {}).get("context_hints", []),
            "hard_context_hints": (download_context or {}).get("hard_context_hints", []),
            "anchor_hints": (download_context or {}).get("anchor_hints", []),
            "min_context_hits": (download_context or {}).get("min_context_hits"),
            "min_hard_context_hits": (download_context or {}).get("min_hard_context_hits"),
            "min_anchor_hits": (download_context or {}).get("min_anchor_hits"),
            "anchor_max_distance": (download_context or {}).get("anchor_max_distance"),
            "context_exclude_hints": (download_context or {}).get("context_exclude_hints", []),
        },
    )
    parsed = _json_loads_if_possible(raw)
    if not isinstance(parsed, dict):
        return {"clicked": False, "selector": None, "error": f"invalid_js_result:{raw}"}
    return parsed


async def _confirm_export_dialog_if_needed(browser_session: BrowserSession) -> dict:
    page = await browser_session.must_get_current_page()
    js = r"""
() => {
  const out = { handled: false, option_clicked: null, confirm_clicked: null, error: null };
  const clean = (v) => String(v ?? "").replace(/\s+/g, " ").trim();

  const isVisible = (el) => {
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };

  const overlayRoots = Array.from(
    document.querySelectorAll(
      ".ant-modal-root,.ant-modal-wrap,.ant-modal,.ant-dropdown,.ant-dropdown-menu,.ant-popover,[role='dialog']"
    )
  ).filter(isVisible);

  if (!overlayRoots.length) return out;

  const clickByRegex = (roots, patternList) => {
    const pool = [];
    for (const root of roots) {
      const nodes = Array.from(root.querySelectorAll("button,a,[role='button'],li,span,div"));
      for (const node of nodes) {
        if (!isVisible(node)) continue;
        const t = clean(node.innerText || node.textContent || "");
        if (!t || t.length > 60) continue;
        pool.push({ node, t });
      }
    }
    for (const p of patternList) {
      const hit = pool.find((x) => p.test(x.t));
      if (hit) {
        const target =
          hit.node.closest("button,a,[role='button'],li,.ant-dropdown-menu-item,.ant-btn,div,span") || hit.node;
        target.scrollIntoView({ block: "center", inline: "center" });
        target.click();
        return hit.t;
      }
    }
    return null;
  };

  const optionText = clickByRegex(overlayRoots, [/全部内容/, /全部维度/, /全部数据/, /^全部$/, /all content/i, /^all$/i]);
  if (optionText) {
    out.handled = true;
    out.option_clicked = optionText;
  }

  const confirmText = clickByRegex(overlayRoots, [/确定/, /导出/, /下载/, /确认/, /^ok$/i, /完成/]);
  if (confirmText) {
    out.handled = true;
    out.confirm_clicked = confirmText;
  }

  return out;
}
"""
    raw = await page.evaluate(js)
    parsed = _json_loads_if_possible(raw)
    if not isinstance(parsed, dict):
        return {"handled": False, "option_clicked": None, "confirm_clicked": None, "error": f"invalid_js_result:{raw}"}
    return parsed


async def _extract_fields_by_specs(browser_session: BrowserSession, specs: list[FieldSpec]) -> list[dict]:
    outputs: list[dict] = []
    for spec in specs:
        try:
            item = await _extract_single_field(browser_session, spec)
        except Exception as exc:  # noqa: BLE001
            item = {
                "field_key": spec.field_key,
                "status": "error",
                "source": "none",
                "value": None,
                "raw_text": None,
                "error": str(exc),
            }
        outputs.append(item)
    return outputs


async def _download_files_by_specs(
    browser_session: BrowserSession,
    specs: list[FileSpec],
    download_dir: Path,
    max_wait_seconds: int | None,
    download_context: dict | None = None,
) -> list[dict]:
    outputs: list[dict] = []

    for spec in specs:
        before = _download_file_meta(download_dir)
        try:
            click_result = await _click_download_trigger(browser_session, spec, download_context=download_context)
        except Exception as exc:  # noqa: BLE001
            outputs.append(
                {
                    "file_key": spec.file_key,
                    "status": "error",
                    "file_path": None,
                    "file_name": None,
                    "clicked_selector": None,
                    "error": str(exc),
                }
            )
            continue

        if not click_result.get("clicked"):
            # Retry once after a short wait; some pages render export actions late.
            await asyncio.sleep(0.45)
            try:
                click_result = await _click_download_trigger(browser_session, spec, download_context=download_context)
            except Exception:
                pass
        if not click_result.get("clicked"):
            # Final retry with relaxed context constraints in the same page/module scope.
            relaxed_context = dict(download_context or {})
            relaxed_context.update(
                {
                    "context_hints": [],
                    "hard_context_hints": [],
                    "anchor_hints": [],
                    "min_context_hits": 0,
                    "min_hard_context_hits": 0,
                    "min_anchor_hits": 0,
                    "anchor_max_distance": 0,
                    "last_interaction_max_distance": 820,
                }
            )
            await asyncio.sleep(0.35)
            try:
                click_result = await _click_download_trigger(browser_session, spec, download_context=relaxed_context)
            except Exception:
                pass
        if not click_result.get("clicked"):
            outputs.append(
                {
                    "file_key": spec.file_key,
                    "status": "not_found",
                    "file_path": None,
                    "file_name": None,
                    "clicked_selector": click_result.get("selector"),
                    "error": click_result.get("error"),
                }
            )
            continue

        timeout = _cap_max_wait_seconds(max_wait_seconds or spec.timeout_seconds or 20)
        try:
            await asyncio.sleep(0.35)
            await _confirm_export_dialog_if_needed(browser_session)
            await asyncio.sleep(0.45)
        except Exception:
            pass

        path = _wait_for_new_file(
            download_dir=download_dir,
            before=before,
            timeout_seconds=timeout,
            keywords=spec.expected_name_keywords,
        )

        if path is None:
            outputs.append(
                {
                    "file_key": spec.file_key,
                    "status": "timeout",
                    "file_path": None,
                    "file_name": None,
                    "clicked_selector": click_result.get("selector"),
                    "error": "no_new_download_detected",
                }
            )
            continue

        stored_path = _snapshot_download_for_file_key(path, download_dir, spec.file_key, name_prefix=spec.name_prefix)
        outputs.append(
            {
                "file_key": spec.file_key,
                "status": "ok",
                "file_path": str(stored_path),
                "file_name": stored_path.name,
                "clicked_selector": click_result.get("selector"),
                "error": None,
            }
        )

    return outputs


async def _check_download_preconditions(
    browser_session: BrowserSession,
    precondition: dict | None,
    runtime_hints: list[str] | None = None,
) -> dict:
    if not isinstance(precondition, dict) or not precondition:
        return {"ok": True, "missing": [], "evidence": {}}

    page = await browser_session.must_get_current_page()
    js = r"""
(input) => {
  const clean = (v) => String(v ?? "").replace(/\s+/g, " ").trim();
  const norm = (v) =>
    clean(v)
      .toLowerCase()
      .replace(/[\.。]/g, "")
      .replace(/[\s_\-—–·:：\/\\|（）()\[\]【】]+/g, "");

  const isVisible = (el) => {
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };

  const bodyText = norm(document.body?.innerText || "");
  const lines = (document.body?.innerText || "")
    .split(/\n+/)
    .map((x) => norm(x))
    .filter(Boolean)
    .slice(0, 1800);

  const result = { ok: true, missing: [], evidence: { select_pairs: [], click_targets: [] } };
  const runtimeHints = Array.isArray(input?.runtime_hints) ? input.runtime_hints.map((x) => norm(x)) : [];
  const runtimeJoined = runtimeHints.join(" ");

  const valueVariants = (v) => {
    const raw = clean(v || "");
    const base = norm(raw);
    const out = new Set([base]);
    if (!base) return Array.from(out);
    for (const token of raw.split(/(?:\/|／|->|→)/)) {
      const nv = norm(token);
      if (nv) out.add(nv);
    }
    if (base.includes("全部全部")) out.add("全部");
    if (base === "投后") out.add("投后复盘");
    if (base === "全部") {
      out.add("全部触达人数");
      out.add("全部触点");
    }
    return Array.from(out);
  };

  const pairList = Array.isArray(input?.select_pairs) ? input.select_pairs : [];
  for (const pair of pairList) {
    const target = norm(pair?.target || "");
    const value = norm(pair?.value || "");
    if (!target || !value) continue;
    const vals = valueVariants(value);
    const lineHit = lines.some((line) =>
      vals.some((vv) => {
        const pairRegex = new RegExp(`${target}.{0,22}${vv}|${vv}.{0,22}${target}`);
        return pairRegex.test(line) || (line.includes(target) && line.includes(vv));
      })
    );
    const bodyHit = vals.some((vv) => {
      const pairRegex = new RegExp(`${target}.{0,22}${vv}|${vv}.{0,22}${target}`);
      return pairRegex.test(bodyText);
    });
    const hintHit = vals.some((vv) => runtimeJoined.includes(target) && runtimeJoined.includes(vv));
    const ok = lineHit || bodyHit || hintHit;
    result.evidence.select_pairs.push({
      target: pair?.target || null,
      value: pair?.value || null,
      ok,
      line_hit: lineHit,
      body_hit: bodyHit,
      hint_hit: hintHit,
    });
    if (!ok) result.missing.push(`select:${pair?.target || ""}=${pair?.value || ""}`);
  }

  const clickTargets = Array.isArray(input?.click_targets) ? input.click_targets : [];
  for (const targetRaw of clickTargets) {
    const target = norm(targetRaw || "");
    if (!target) continue;
    let visibleHit = false;
    for (const node of Array.from(document.querySelectorAll("a,button,span,div,li,[role='tab'],[role='button'],[role='menuitem']"))) {
      if (!isVisible(node)) continue;
      const t = norm(node.innerText || node.textContent || "");
      if (!t) continue;
      if (t.includes(target) || target.includes(t)) {
        visibleHit = true;
        break;
      }
    }
    const hintHit = runtimeJoined.includes(target);
    const ok = visibleHit || hintHit;
    result.evidence.click_targets.push({ target: targetRaw, ok, visible_hit: visibleHit, hint_hit: hintHit });
    if (!ok) result.missing.push(`click:${targetRaw}`);
  }

  result.ok = result.missing.length === 0;
  return result;
}
"""
    raw = await page.evaluate(
        js,
        {
            "select_pairs": precondition.get("select_pairs", []),
            "click_targets": precondition.get("click_targets", []),
            "runtime_hints": runtime_hints or [],
        },
    )
    parsed = _json_loads_if_possible(raw)
    if not isinstance(parsed, dict):
        return {"ok": False, "missing": ["invalid_precondition_check_result"], "evidence": {"raw": raw}}
    return parsed

async def _click_text_step(
    browser_session: BrowserSession,
    label: str,
    selectors: list[str],
    *,
    allow_reverse_contains: bool = True,
    avoid_global_nav: bool = False,
    require_clickable: bool = False,
    max_text_len: int | None = None,
    block_text_tokens: list[str] | None = None,
    strict_selectors_only: bool = False,
    prefer_near_last_interaction: bool = True,
    last_interaction_max_distance: int = 560,
) -> dict:
    page = await browser_session.must_get_current_page()
    js = r"""
(input) => {
  const label = (input && input.label ? input.label : "").trim();
  const selectors = (input && Array.isArray(input.selectors)) ? input.selectors : [];
  const allowReverseContains = input?.allow_reverse_contains !== false;
  const avoidGlobalNav = input?.avoid_global_nav === true;
  const requireClickable = input?.require_clickable === true;
  const strictSelectorsOnly = input?.strict_selectors_only === true;
  const preferNearLastInteraction = input?.prefer_near_last_interaction !== false;
  const lastInteractionMaxDistance = Number.isFinite(Number(input?.last_interaction_max_distance))
    ? Number(input.last_interaction_max_distance)
    : 560;
  const maxTextLen = Number.isFinite(Number(input?.max_text_len)) ? Number(input?.max_text_len) : null;
  const blockTextTokens = Array.isArray(input?.block_text_tokens) ? input.block_text_tokens : [];
  const out = { clicked: false, selector: null, text: null, matched_label: null, error: null };

  const clean = (v) => String(v ?? "").replace(/\s+/g, " ").trim();
  const norm = (v) => clean(v).toLowerCase();
  const softNorm = (v) =>
    norm(v)
      .replace(/[\.。]/g, "")
      .replace(/[\s_\-—–·:：\/\\|（）()\[\]【】]+/g, "");
  const isPureNumberLike = (v) => /^-?\d+(?:[\.,]\d+)?%?$/.test(clean(v || ""));

  const isVisible = (el) => {
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  const isLikelyGlobalNav = (node) => {
    if (!node) return false;
    if (node.closest("header,aside,nav,[role='banner']")) return true;
    let cur = node;
    for (let depth = 0; depth < 5 && cur; depth++) {
      const cls = clean(cur.className || "").toLowerCase();
      if (cls && /(header|topbar|navbar|toolbar|message|notice|user|avatar)/.test(cls)) {
        const r = cur.getBoundingClientRect();
        if (r && r.top < 260) return true;
      }
      cur = cur.parentElement;
    }
    return false;
  };
  const getLastInteraction = () => {
    try {
      const raw = window.__opt_last_interaction;
      if (!raw || !Number.isFinite(raw.x) || !Number.isFinite(raw.y)) return null;
      return {
        x: Number(raw.x),
        y: Number(raw.y),
        url: String(raw.url || ""),
      };
    } catch (_) {
      return null;
    }
  };

  const unique = new Set();
  const candidates = [];
  const addNodes = (nodes, sourceSelector) => {
    for (const node of nodes) {
      if (!node || unique.has(node)) continue;
      unique.add(node);
      const text = clean(node.innerText || node.textContent || "");
      if (!text || !isVisible(node)) continue;
      candidates.push({ node, text, sourceSelector });
    }
  };

  if (selectors.length) {
    for (const sel of selectors) {
      try {
        addNodes(Array.from(document.querySelectorAll(sel)), sel);
      } catch (_) {}
    }
  }

  if (!candidates.length && !strictSelectorsOnly) {
    addNodes(
      Array.from(document.querySelectorAll("[role='menuitem'],.ant-menu-item,.ant-tabs-tab,a,button,li,span")),
      "fallback:candidate"
    );
  }
  if (!candidates.length && strictSelectorsOnly) {
    out.error = "selector_scoped_target_not_found";
    return out;
  }

  const target = norm(label);
  const targetSoft = softNorm(label);
  const isTop5LikeTarget = /top\s*5|top5|均值|行业/.test(targetSoft);
  const stageMatch = targetSoft.match(/a([1-5])/);
  const expectedStage = stageMatch ? `a${stageMatch[1]}` : "";
  const detectStage = (text) => {
    const m = softNorm(text || "").match(/a([1-5])/);
    return m ? `a${m[1]}` : "";
  };
  const blockedNorm = blockTextTokens.map((v) => norm(v)).filter(Boolean);
  const isBlockedTarget = blockedNorm.some((tok) => target.includes(tok) || tok.includes(target));
  if (!target) {
    out.error = "empty_label";
    return out;
  }

  const scored = candidates
    .map((item) => {
      const t = norm(item.text);
      const ts = softNorm(item.text);
      if (!t || !ts) return null;
      if (maxTextLen && item.text.length > maxTextLen) return null;

      if (!/^[-\d\.,%]+$/.test(targetSoft) && isPureNumberLike(item.text)) return null;

      const exact = ts === targetSoft || t === target || t.replace(/[:：]/g, "") === target;
      const starts = ts.startsWith(targetSoft);
      const contains = ts.includes(targetSoft);
      const reverseContains = targetSoft.length >= 3 && targetSoft.includes(ts);
      if (!(exact || starts || contains || (allowReverseContains && reverseContains))) return null;
      if (targetSoft.length <= 2 && !(exact || starts)) return null;
      if (!isBlockedTarget && blockedNorm.some((tok) => t.includes(tok))) return null;

      const lengthPenalty = Math.abs(ts.length - targetSoft.length) * 20;
      let score = -item.text.length - lengthPenalty;
      if (exact) score += 12000;
      else if (starts) score += 3200;
      else if (contains) score += 1400;
      else if (reverseContains) score += 200;
      if (/[：:]$/.test(item.text)) score += 240;
      if (/[a-z0-9]/i.test(targetSoft) && !/[a-z0-9]/i.test(ts)) score -= 80;
      if (avoidGlobalNav) {
        const inGlobalNav = isLikelyGlobalNav(item.node);
        if (inGlobalNav) return null;
      }
      if (requireClickable) {
        const clickable = item.node.closest(
          "a,button,[role='button'],[role='menuitem'],[role='tab'],li,.ant-menu-item,.ant-menu-submenu-title,.ant-tabs-tab,.ant-dropdown-menu-item"
        );
        if (!clickable) return null;
      }
      if (preferNearLastInteraction) {
        const last = getLastInteraction();
        if (last) {
          const rect = item.node.getBoundingClientRect();
          const cx = rect.left + rect.width / 2;
          const cy = rect.top + rect.height / 2;
          const dx = cx - last.x;
          const dy = cy - last.y;
          const dist = Math.sqrt(dx * dx + dy * dy);
          if (dist <= lastInteractionMaxDistance) {
            score += Math.max(0, 1600 - dist * 2.2);
          } else {
            score -= Math.min(600, (dist - lastInteractionMaxDistance) * 0.2);
          }
        }
      }
      return { ...item, score };
    })
    .filter(Boolean)
    .sort((a, b) => b.score - a.score);

  if (!scored.length) {
    out.error = "menu_target_not_found";
    return out;
  }

  const best = scored[0];
  const clickSameTab = (node) => {
    const clickable =
      node?.closest("a,button,[role='button'],li,.ant-btn,div,span") || node;
    if (!clickable) return false;
    const anchor =
      clickable.tagName === "A" ? clickable : clickable.closest("a[href]");
    if (anchor) {
      try {
        const href = String(anchor.getAttribute("href") || "").trim();
        const url = new URL(href || anchor.href || "", location.href);
        // Keep clicks within current origin to avoid accidental jump to docs/notice center.
        if (url && url.origin && url.origin !== location.origin) return false;
      } catch (_) {}
      try {
        anchor.setAttribute("target", "_self");
        anchor.removeAttribute("rel");
      } catch (_) {}
      try {
        anchor.dispatchEvent(
          new MouseEvent("click", {
            bubbles: true,
            cancelable: true,
            button: 0,
            ctrlKey: false,
            metaKey: false,
            shiftKey: false,
            altKey: false,
          })
        );
        return true;
      } catch (_) {}
    }
    clickable.click();
    return true;
  };
  try {
    best.node.scrollIntoView({ block: "center", inline: "center" });
    const didClick = clickSameTab(best.node);
    if (!didClick) {
      out.error = "external_navigation_blocked";
      return out;
    }
    const r = best.node.getBoundingClientRect();
    window.__opt_last_interaction = {
      kind: "click",
      text: clean(best.text || ""),
      x: r.left + r.width / 2,
      y: r.top + r.height / 2,
      ts: Date.now(),
      url: `${location.pathname}${location.search}`,
    };
    out.clicked = true;
    out.selector = best.sourceSelector || null;
    out.text = best.text;
    return out;
  } catch (e) {
    out.error = String(e);
    return out;
  }
}
"""

    raw = await page.evaluate(
        js,
        {
            "label": label,
            "selectors": selectors,
            "allow_reverse_contains": allow_reverse_contains,
            "avoid_global_nav": avoid_global_nav,
            "require_clickable": require_clickable,
            "max_text_len": max_text_len,
            "block_text_tokens": block_text_tokens or [],
            "strict_selectors_only": strict_selectors_only,
            "prefer_near_last_interaction": prefer_near_last_interaction,
            "last_interaction_max_distance": last_interaction_max_distance,
        },
    )
    parsed = _json_loads_if_possible(raw)
    if not isinstance(parsed, dict):
        return {"clicked": False, "selector": None, "text": None, "error": f"invalid_js_result:{raw}"}
    return parsed


async def _hover_text_step(
    browser_session: BrowserSession,
    label: str,
    selectors: list[str],
    *,
    semantic_text: str | None = None,
    allow_reverse_contains: bool = True,
    avoid_global_nav: bool = True,
    max_text_len: int | None = None,
    block_text_tokens: list[str] | None = None,
    strict_selectors_only: bool = False,
    wait_seconds: float = 0.6,
) -> dict:
    page = await browser_session.must_get_current_page()
    js = r"""
(input) => {
  const label = (input && input.label ? input.label : "").trim();
  const semanticText = String(input?.semantic_text || label || "").trim();
  const selectors = (input && Array.isArray(input.selectors)) ? input.selectors : [];
  const allowReverseContains = input?.allow_reverse_contains !== false;
  const avoidGlobalNav = input?.avoid_global_nav !== false;
  const strictSelectorsOnly = input?.strict_selectors_only === true;
  const maxTextLen = Number.isFinite(Number(input?.max_text_len)) ? Number(input?.max_text_len) : null;
  const blockTextTokens = Array.isArray(input?.block_text_tokens) ? input.block_text_tokens : [];
  const out = { hovered: false, selector: null, text: null, matched_label: null, error: null, x: null, y: null };

  const clean = (v) => String(v ?? "").replace(/\s+/g, " ").trim();
  const norm = (v) => clean(v).toLowerCase();
  const softNorm = (v) =>
    norm(v)
      .replace(/[\.。]/g, "")
      .replace(/[\s_\-—–·:：\/\\|（）()\[\]【】]+/g, "");
  const isVisible = (el) => {
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  const isLikelyGlobalNav = (node) => {
    if (!node) return false;
    if (node.closest("header,aside,nav,[role='banner']")) return true;
    let cur = node;
    for (let depth = 0; depth < 5 && cur; depth++) {
      const cls = clean(cur.className || "").toLowerCase();
      if (cls && /(header|topbar|navbar|toolbar|message|notice|user|avatar)/.test(cls)) {
        const r = cur.getBoundingClientRect();
        if (r && r.top < 260) return true;
      }
      cur = cur.parentElement;
    }
    return false;
  };
  const isPlusLike = (node) => {
    const t = clean(node?.innerText || node?.textContent || "");
    const title = clean(node?.getAttribute?.("title") || "");
    const aria = clean(node?.getAttribute?.("aria-label") || "");
    const cls = clean(node?.className || "");
    const ancestorCls = clean(node?.closest?.("[class]")?.className || "");
    const joined = `${t} ${title} ${aria} ${cls} ${ancestorCls}`.toLowerCase();
    if (t === "+" || t === "＋") return true;
    return /plus|add|新增|创建|添加|展开|收起|icon-jia|jiahao|\bi-add\b/.test(joined);
  };
  const isHintTrigger = (node) => {
    const t = clean(node?.innerText || node?.textContent || "");
    const title = clean(node?.getAttribute?.("title") || "");
    const aria = clean(node?.getAttribute?.("aria-label") || "");
    const cls = clean(node?.className || "");
    const ancestorCls = clean(node?.closest?.("[class]")?.className || "");
    const joined = `${t} ${title} ${aria} ${cls} ${ancestorCls}`.toLowerCase();
    if (t === "?" || t === "？") return true;
    if (node?.hasAttribute?.("data-popover-id")) return true;
    if (node?.hasAttribute?.("aria-describedby")) return true;
    return /question|question-circle|anticon-question|help|tip|tooltip|提示|问号|问|i-icon-help|icon-help|evaluation-popper-trigger|popper-trigger|popover-trigger|tooltip-trigger/.test(joined);
  };
  const hintLike = (node) => {
    const t = clean(node?.innerText || node?.textContent || "");
    const title = clean(node?.getAttribute?.("title") || "");
    const aria = clean(node?.getAttribute?.("aria-label") || "");
    const cls = clean(node?.className || "");
    const joined = `${t} ${title} ${aria} ${cls}`.toLowerCase();
    if (t === "?" || t === "？") return 220;
    if (/question|question-circle|anticon-question|help|tip|tooltip|提示|问号|问|i-icon-help|icon-help/.test(joined)) return 180;
    return -160;
  };

  const target = norm(label);
  const targetSoft = softNorm(label);
  const semanticSoft = softNorm(semanticText || label);
  const isTop5LikeTarget = /top\s*5|top5|均值|行业/.test(targetSoft);
  const semanticNeedsBenchmark = /top\s*5|top5|均值|行业/.test(semanticSoft);
  const semanticNeedsFlow = /流转/.test(semanticSoft);
  const stageMatch = (targetSoft || semanticSoft || "").match(/a([1-5])/);
  const expectedStage = stageMatch ? `a${stageMatch[1]}` : "";
  const detectStage = (text) => {
    const m = softNorm(text || "").match(/a([1-5])/);
    return m ? `a${m[1]}` : "";
  };
  const blockedNorm = blockTextTokens.map((v) => norm(v)).filter(Boolean);
  if (!target) {
    out.error = "empty_label";
    return out;
  }

  const benchmarkSemanticHit = (text) => {
    const s = softNorm(text || "");
    if (!s) return 0;
    let hit = 0;
    if (/top\s*5|top5/.test(s)) hit += 1;
    if (s.includes("行业")) hit += 1;
    if (s.includes("均值")) hit += 1;
    if (semanticNeedsFlow && s.includes("流转")) hit += 1;
    return hit;
  };

  const labelNodes = Array.from(document.querySelectorAll("main *,.ant-layout-content *,[class*='content'] *"))
    .filter((node) => {
      if (!isVisible(node)) return false;
      const text = clean(node.innerText || node.textContent || "");
      if (!text) return false;
      if (maxTextLen && text.length > maxTextLen) return false;
      const n = norm(text);
      const ns = softNorm(text);
      if (isTop5LikeTarget && !/top\s*5|top5|均值|行业/.test(ns)) return false;
      const reverseContains = allowReverseContains && ns.length >= 4 && targetSoft.includes(ns);
      const reverseTargetContains = target.includes(n) && n.length >= 4;
      if (!(n.includes(target) || reverseTargetContains || ns.includes(targetSoft) || reverseContains)) return false;
      if (expectedStage) {
        const nodeStage = detectStage(text);
        if (nodeStage && nodeStage !== expectedStage) return false;
        if (!nodeStage) {
          const holder =
            node.closest("div,li,td,th,tr,section,article,[class*='card'],[class*='container']") ||
            node.parentElement ||
            node;
          const holderText = clean(holder?.innerText || holder?.textContent || "");
          const holderSoft = softNorm(holderText);
          const holderStages = Array.from(holderSoft.matchAll(/a([1-5])/g)).map((m) => `a${m[1]}`);
          if (holderStages.length && !holderStages.includes(expectedStage)) return false;
        }
      }
      if (semanticNeedsBenchmark) {
        const holder =
          node.closest("div,li,td,th,tr,section,article,[class*='card'],[class*='container']") ||
          node.parentElement ||
          node;
        const holderText = clean(holder?.innerText || holder?.textContent || "");
        const semHit = benchmarkSemanticHit(`${text} ${holderText}`);
        if (semHit < 2) return false;
      }
      if (blockedNorm.some((tok) => n.includes(tok))) return false;
      return true;
    })
    .slice(0, 120);

  const preferredContainers = (() => {
    if (!labelNodes.length) return [];
    const scored = [];
    const seenContainers = new Set();
    for (const ln of labelNodes.slice(0, 36)) {
      const container =
        ln.closest("div,li,td,th,tr,section,article,[class*='card'],[class*='container']") ||
        ln.parentElement ||
        document.body;
      if (!container || seenContainers.has(container)) continue;
      seenContainers.add(container);
      if (!isVisible(container)) continue;
      const text = clean(container.innerText || container.textContent || "");
      if (!text) continue;
      const ns = softNorm(text);
      let score = 0;
      if (targetSoft && ns.includes(targetSoft)) score += 320;
      if (semanticNeedsFlow && ns.includes("流转")) score += 180;
      if (semanticNeedsBenchmark) {
        if (/top\s*5|top5|均值|行业/.test(ns)) score += 320;
        else score -= 260;
      }
      if (expectedStage) {
        const stages = Array.from(ns.matchAll(/a([1-5])/g)).map((m) => `a${m[1]}`);
        if (stages.length && !stages.includes(expectedStage)) score -= 520;
        if (stages.includes(expectedStage)) score += 260;
      }
      // Prefer card-like containers with moderate text size.
      score += Math.max(0, 260 - Math.min(text.length, 260));
      scored.push({ container, score });
    }
    return scored
      .sort((a, b) => b.score - a.score)
      .slice(0, 4)
      .map((x) => x.container);
  })();

  const isInPreferredContainer = (node) => {
    if (!preferredContainers.length) return true;
    return preferredContainers.some((c) => c === node || c.contains(node));
  };

  const candidates = [];
  const seen = new Set();
  const addCandidate = (node, source, restrictPreferred = false) => {
    if (!node || seen.has(node)) return;
    if (!isVisible(node)) return;
    if (avoidGlobalNav && isLikelyGlobalNav(node)) return;
    if (restrictPreferred && !isInPreferredContainer(node)) return;
    if (isPlusLike(node)) return;
    if (!isHintTrigger(node)) return;
    const rect = node.getBoundingClientRect();
    if (!rect || rect.width <= 0 || rect.height <= 0) return;
    seen.add(node);
    const text = clean(node.innerText || node.textContent || "");
    const scoreHint = hintLike(node);
    candidates.push({ node, source, text, rect, scoreHint });
  };

  const anchorHintSelectors = [
    ".evaluation-popper-trigger",
    "[class*='popper-trigger']",
    "[class*='popover-trigger']",
    "[class*='tooltip-trigger']",
    ".i-icon-help",
    "[class*='icon-help']",
    "[class*='question']",
    "[class*='help']",
    "[data-popover-id]",
    "[aria-describedby]",
  ];
  if (labelNodes.length) {
    for (const ln of labelNodes.slice(0, 24)) {
      const container =
        ln.closest("div,li,td,th,tr,section,article,[class*='card'],[class*='container']") ||
        ln.parentElement ||
        document.body;
      if (!container) continue;
      for (const sel of anchorHintSelectors) {
        try {
          for (const n of Array.from(container.querySelectorAll(sel))) addCandidate(n, "anchor_hint", true);
        } catch (_) {}
      }
    }
  }

  for (const sel of selectors) {
    try {
      for (const n of Array.from(document.querySelectorAll(sel))) addCandidate(n, sel, true);
    } catch (_) {}
  }
  if (!candidates.length) {
    for (const sel of selectors) {
      try {
        for (const n of Array.from(document.querySelectorAll(sel))) addCandidate(n, sel, false);
      } catch (_) {}
    }
  }
  if (!candidates.length && !strictSelectorsOnly) {
    for (const n of Array.from(document.querySelectorAll("main [class*='question'],main [class*='help'],main [class*='tip'],main [class*='tooltip'],main [data-popover-id],main [aria-describedby],main .i-icon-help,main .evaluation-popper-trigger,.ant-layout-content [class*='question'],.ant-layout-content [class*='help'],.ant-layout-content [class*='tip'],.ant-layout-content [class*='tooltip'],.ant-layout-content [data-popover-id],.ant-layout-content [aria-describedby],.ant-layout-content .i-icon-help,.ant-layout-content .evaluation-popper-trigger"))) {
      addCandidate(n, "fallback:hover_trigger", true);
    }
  }
  if (!candidates.length && !strictSelectorsOnly) {
    for (const n of Array.from(document.querySelectorAll("main [class*='question'],main [class*='help'],main [class*='tip'],main [class*='tooltip'],main [data-popover-id],main [aria-describedby],main .i-icon-help,main .evaluation-popper-trigger,.ant-layout-content [class*='question'],.ant-layout-content [class*='help'],.ant-layout-content [class*='tip'],.ant-layout-content [class*='tooltip'],.ant-layout-content [data-popover-id],.ant-layout-content [aria-describedby],.ant-layout-content .i-icon-help,.ant-layout-content .evaluation-popper-trigger"))) {
      addCandidate(n, "fallback:hover_trigger", false);
    }
  }
  if (!candidates.length) {
    out.error = "hover_target_not_found";
    return out;
  }

  const scored = candidates
    .map((item) => {
      let score = item.scoreHint;
      if (item.scoreHint <= -120) return null;
      if (labelNodes.length) {
        let nearest = 1e9;
        let nearestLabel = null;
        for (const ln of labelNodes) {
          const lr = ln.getBoundingClientRect();
          const cx = item.rect.left + item.rect.width / 2;
          const cy = item.rect.top + item.rect.height / 2;
          const lx = lr.left + lr.width / 2;
          const ly = lr.top + lr.height / 2;
          const d = Math.hypot(cx - lx, cy - ly);
          if (d < nearest) {
            nearest = d;
            nearestLabel = ln;
          }
        }
        if (nearest > 280) return null;
        score += Math.max(0, 1800 - nearest * 2.2);
        item.matchedLabel = nearestLabel ? clean(nearestLabel.innerText || nearestLabel.textContent || "") : null;
        if (isTop5LikeTarget) {
          const matchedNorm = softNorm(item.matchedLabel || "");
          if (!/top\s*5|top5|均值|行业/.test(matchedNorm)) return null;
          if (nearest > 220) return null;
        }
        if (semanticNeedsBenchmark) {
          const matchedCtx =
            `${item.matchedLabel || ""} ${item.node?.closest?.("div,li,td,th,tr,section,article,[class*='card'],[class*='container']")?.innerText || ""}`;
          if (benchmarkSemanticHit(matchedCtx) < 2) return null;
        }
        if (expectedStage) {
          const matchedSoft = softNorm(item.matchedLabel || "");
          const stages = Array.from(matchedSoft.matchAll(/a([1-5])/g)).map((m) => `a${m[1]}`);
          if (stages.length && !stages.includes(expectedStage)) return null;
        }
      } else {
        const joined = norm(`${item.text} ${item.node.getAttribute?.("title") || ""} ${item.node.getAttribute?.("aria-label") || ""}`);
        if (isTop5LikeTarget && !/top\s*5|top5|均值|行业/.test(softNorm(joined))) return null;
        if (joined.includes(target) || softNorm(joined).includes(targetSoft)) score += 600;
      }
      if (item.scoreHint >= 200) score += 600;
      else if (item.scoreHint < 180) score -= 320;
      if (item.source === "anchor_hint") score += 1200;
      return { ...item, score };
    })
    .filter(Boolean)
    .sort((a, b) => b.score - a.score);

  if (!scored.length) {
    out.error = "hover_target_not_found";
    return out;
  }

  const best = scored[0];
  try {
    best.node.scrollIntoView({ block: "center", inline: "center" });
    const r = best.node.getBoundingClientRect();
    const x = Math.round(r.left + r.width / 2);
    const y = Math.round(r.top + r.height / 2);
    best.node.dispatchEvent(new MouseEvent("mouseenter", { bubbles: true }));
    best.node.dispatchEvent(new MouseEvent("mouseover", { bubbles: true }));
    best.node.dispatchEvent(new MouseEvent("mousemove", { bubbles: true, clientX: x, clientY: y }));
    window.__opt_last_interaction = {
      kind: "hover",
      text: clean(best.text || ""),
      x,
      y,
      ts: Date.now(),
      url: `${location.pathname}${location.search}`,
    };
    out.hovered = true;
    out.selector = best.source || null;
    out.text = clean(best.text || "");
    out.matched_label = best.matchedLabel || null;
    out.x = x;
    out.y = y;
    return out;
  } catch (e) {
    out.error = String(e);
    return out;
  }
}
"""
    raw = await page.evaluate(
        js,
        {
            "label": label,
            "semantic_text": semantic_text or label,
            "selectors": selectors,
            "allow_reverse_contains": allow_reverse_contains,
            "avoid_global_nav": avoid_global_nav,
            "max_text_len": max_text_len,
            "block_text_tokens": block_text_tokens or [],
            "strict_selectors_only": strict_selectors_only,
        },
    )
    parsed = _json_loads_if_possible(raw)
    if not isinstance(parsed, dict):
        return {"hovered": False, "selector": None, "text": None, "matched_label": None, "error": f"invalid_js_result:{raw}"}

    if parsed.get("hovered") and parsed.get("x") is not None and parsed.get("y") is not None:
        try:
            mouse = await page.mouse
            await mouse.move(x=int(parsed["x"]), y=int(parsed["y"]), steps=6)
            await asyncio.sleep(max(0.0, min(float(wait_seconds), 2.2)))
        except Exception:
            pass
    return parsed


async def _click_nearby_by_label(
    browser_session: BrowserSession,
    target_label: str,
    *,
    avoid_global_nav: bool = True,
    block_text_tokens: list[str] | None = None,
) -> dict:
    page = await browser_session.must_get_current_page()
    js = r"""
(input) => {
  const label = (input?.target_label || "").trim();
  const avoidGlobalNav = input?.avoid_global_nav !== false;
  const blockTextTokens = Array.isArray(input?.block_text_tokens) ? input.block_text_tokens : [];
  const out = { clicked: false, selector: null, text: null, error: null };
  if (!label) {
    out.error = "empty_target_label";
    return out;
  }

  const clean = (v) => String(v ?? "").replace(/\s+/g, " ").trim();
  const norm = (v) =>
    clean(v)
      .toLowerCase()
      .replace(/[\.。]/g, "")
      .replace(/[\s_\-—–·:：\/\\|（）()\[\]【】]+/g, "");
  const isVisible = (el) => {
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  const isLikelyGlobalNav = (node) => {
    if (!node) return false;
    if (node.closest("header,aside,nav,[role='banner']")) return true;
    let cur = node;
    for (let depth = 0; depth < 5 && cur; depth++) {
      const cls = clean(cur.className || "").toLowerCase();
      if (cls && /(header|topbar|navbar|toolbar|message|notice|user|avatar)/.test(cls)) {
        const r = cur.getBoundingClientRect();
        if (r && r.top < 260) return true;
      }
      cur = cur.parentElement;
    }
    return false;
  };

  const blockedNorm = blockTextTokens.map((v) => norm(v)).filter(Boolean);
  const labelNorm = norm(label);
  const last = window.__opt_last_interaction || null;
  const canUseLast =
    !!last &&
    Date.now() - Number(last.ts || 0) <= 20000 &&
    String(last.url || "") === `${location.pathname}${location.search}` &&
    Number.isFinite(Number(last.x)) &&
    Number.isFinite(Number(last.y));
  const distanceToLast = (node) => {
    try {
      const r = node.getBoundingClientRect();
      const cx = r.left + r.width / 2;
      const cy = r.top + r.height / 2;
      const dx = cx - Number(last.x || 0);
      const dy = cy - Number(last.y || 0);
      return Math.sqrt(dx * dx + dy * dy);
    } catch (_) {
      return 1e9;
    }
  };
  const labelNodes = Array.from(document.querySelectorAll("label,span,div,li,th,td,a,button"))
    .filter((node) => {
      if (!isVisible(node)) return false;
      const t = norm(node.innerText || node.textContent || "");
      if (!t) return false;
      if (!(t.includes(labelNorm) || labelNorm.includes(t))) return false;
      if (blockedNorm.some((tok) => t.includes(tok))) return false;
      return true;
    })
    .slice(0, 120);

  const clickableSelectors = [
    ".ant-tabs-tab",
    "[role='tab']",
    ".ant-menu-item",
    "[role='menuitem']",
    ".ant-dropdown-menu-item",
    "a",
    "button",
    "[role='button']",
    "li",
    "span",
    "div",
  ].join(",");

  const candidates = [];
  const addCandidate = (node, source, labelRect = null) => {
    if (!node || !isVisible(node)) return;
    if (avoidGlobalNav) {
      const inGlobalNav = isLikelyGlobalNav(node);
      if (inGlobalNav) return;
    }
    const rect = node.getBoundingClientRect();
    if (!rect || rect.width <= 0 || rect.height <= 0) return;
    const text = clean(node.innerText || node.textContent || "");
    if (!text) return;
    const tn = norm(text);
    if (!tn) return;
    if (blockedNorm.some((tok) => tn.includes(tok))) return;
    candidates.push({ node, source, text, rect, labelRect });
  };

  for (const ln of labelNodes) {
    const lnRect = ln.getBoundingClientRect();
    const container =
      ln.closest(".ant-tabs,.ant-tabs-nav,.ant-tabs-content,.ant-menu,.ant-card,section,article,div,li,td") ||
      ln.parentElement ||
      document.body;
    if (!container) continue;
    const local = Array.from(container.querySelectorAll(clickableSelectors));
    for (const n of local) addCandidate(n, "label_container_clickable", lnRect);

    let cur = container;
    for (let depth = 0; depth < 3 && cur; depth++) {
      const parent = cur.parentElement;
      if (!parent) break;
      const inParent = Array.from(parent.querySelectorAll(clickableSelectors));
      for (const n of inParent) addCandidate(n, "label_ancestor_clickable", lnRect);
      cur = parent;
    }
  }

  if (!candidates.length) {
    out.error = "nearby_click_target_not_found";
    return out;
  }

  const scoreOf = (item) => {
    const cls = clean(item.node.className || "").toLowerCase();
    let score = 0;
    if (/ant-tabs-tab|role='tab'|role="tab"/.test(cls) || item.node.getAttribute?.("role") === "tab") score += 3200;
    else if (/ant-menu-item|menuitem/.test(cls) || item.node.getAttribute?.("role") === "menuitem") score += 2600;
    else if (item.node.matches?.("a,button,[role='button']")) score += 1600;
    else score += 700;
    const ts = norm(item.text || "");
    const exact = ts === labelNorm;
    const starts = ts.startsWith(labelNorm);
    const contains = ts.includes(labelNorm);
    if (exact) score += 8000;
    else if (starts) score += 2400;
    else if (contains) score += 1200;
    if (item.labelRect) {
      const lcx = item.labelRect.left + item.labelRect.width / 2;
      const lcy = item.labelRect.top + item.labelRect.height / 2;
      const cx = item.rect.left + item.rect.width / 2;
      const cy = item.rect.top + item.rect.height / 2;
      const dx = cx - lcx;
      const dy = cy - lcy;
      score -= Math.sqrt(dx * dx + dy * dy) * 0.9;
    }
    return score;
  };

  const best = candidates
    .map((c) => ({ ...c, score: scoreOf(c) }))
    .sort((a, b) => b.score - a.score)[0];

  const clickSameTab = (node) => {
    const clickable =
      node?.closest("a,button,[role='button'],[role='menuitem'],[role='tab'],li,.ant-btn,.ant-tabs-tab,.ant-menu-item,div,span") ||
      node;
    if (!clickable) return false;
    const anchor = clickable.tagName === "A" ? clickable : clickable.closest("a[href]");
    if (anchor) {
      try {
        const href = String(anchor.getAttribute("href") || "").trim();
        const url = new URL(href || anchor.href || "", location.href);
        if (url && url.origin && url.origin !== location.origin) return false;
      } catch (_) {}
      try {
        anchor.setAttribute("target", "_self");
        anchor.removeAttribute("rel");
      } catch (_) {}
      try {
        anchor.dispatchEvent(
          new MouseEvent("click", {
            bubbles: true,
            cancelable: true,
            button: 0,
            ctrlKey: false,
            metaKey: false,
            shiftKey: false,
            altKey: false,
          })
        );
        return true;
      } catch (_) {}
    }
    clickable.click();
    return true;
  };

  try {
    best.node.scrollIntoView({ block: "center", inline: "center" });
    const didClick = clickSameTab(best.node);
    if (!didClick) {
      out.error = "external_navigation_blocked";
      return out;
    }
    const r = best.node.getBoundingClientRect();
    window.__opt_last_interaction = {
      kind: "click_by_label",
      text: clean(best.text || ""),
      x: r.left + r.width / 2,
      y: r.top + r.height / 2,
      ts: Date.now(),
      url: `${location.pathname}${location.search}`,
    };
    out.clicked = true;
    out.selector = best.source || null;
    out.text = clean(best.text || "");
    return out;
  } catch (e) {
    out.error = String(e);
    return out;
  }
}
"""
    raw = await page.evaluate(
        js,
        {
            "target_label": target_label,
            "avoid_global_nav": avoid_global_nav,
            "block_text_tokens": block_text_tokens or [],
        },
    )
    parsed = _json_loads_if_possible(raw)
    if not isinstance(parsed, dict):
        return {"clicked": False, "selector": None, "text": None, "error": f"invalid_js_result:{raw}"}
    return parsed


async def _scroll_main_content_to_top(browser_session: BrowserSession) -> None:
    page = await browser_session.must_get_current_page()
    js = r"""
() => {
  const candidates = [
    "main",
    ".ant-layout-content",
    "[class*='content']",
    "[role='main']",
    ".ant-tabs-content",
  ];
  for (const sel of candidates) {
    const el = document.querySelector(sel);
    if (!el) continue;
    try {
      el.scrollTop = 0;
      if (typeof el.scrollTo === "function") el.scrollTo({ top: 0, behavior: "instant" });
    } catch (_) {}
  }
  try {
    window.scrollTo({ top: 0, behavior: "instant" });
  } catch (_) {}
  return true;
}
"""
    try:
        await page.evaluate(js)
    except Exception:
        return


async def _select_option_by_text(
    browser_session: BrowserSession,
    option_text: str,
    *,
    target_label: str | None = None,
    avoid_global_nav: bool = True,
    block_text_tokens: list[str] | None = None,
) -> dict:
    page = await browser_session.must_get_current_page()
    js = r"""
(input) => {
  const target = (input?.option_text || "").trim();
  const targetLabel = (input?.target_label || "").trim();
  const avoidGlobalNav = input?.avoid_global_nav !== false;
  const blockTextTokens = Array.isArray(input?.block_text_tokens) ? input.block_text_tokens : [];
  const out = { selected: false, text: null, error: null };
  if (!target) {
    out.error = "empty_option_text";
    return out;
  }

  const clean = (v) => String(v ?? "").replace(/\s+/g, " ").trim();
  const norm = (v) => clean(v).toLowerCase();
  const softNorm = (v) =>
    norm(v)
      .replace(/[\.。]/g, "")
      .replace(/[\s_\-—–·:：\/\\|（）()\[\]【】]+/g, "");

  const isVisible = (el) => {
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  const isLikelyGlobalNav = (node) => {
    if (!node) return false;
    if (node.closest("header,aside,nav,[role='banner']")) return true;
    let cur = node;
    for (let depth = 0; depth < 5 && cur; depth++) {
      const cls = clean(cur.className || "").toLowerCase();
      if (cls && /(header|topbar|navbar|toolbar|message|notice|user|avatar)/.test(cls)) {
        const r = cur.getBoundingClientRect();
        if (r && r.top < 260) return true;
      }
      cur = cur.parentElement;
    }
    return false;
  };
  const distanceToLast = (node, last) => {
    try {
      const r = node.getBoundingClientRect();
      const cx = r.left + r.width / 2;
      const cy = r.top + r.height / 2;
      const dx = cx - Number(last?.x || 0);
      const dy = cy - Number(last?.y || 0);
      return Math.sqrt(dx * dx + dy * dy);
    } catch (_) {
      return 1e9;
    }
  };
  const last = window.__opt_last_interaction || null;
  const canUseLast =
    !!last &&
    Date.now() - Number(last.ts || 0) <= 20000 &&
    String(last.url || "") === `${location.pathname}${location.search}` &&
    Number.isFinite(Number(last.x)) &&
    Number.isFinite(Number(last.y));

  const overlayOptionSelectors = [
    ".ant-select-item-option-content",
    ".ant-select-item-option",
    "[role='option']",
    "[role='menuitem']",
    "[role='treeitem']",
    ".ant-cascader-menu-item",
    ".ant-cascader-option",
    ".ant-cascader-menu-item-content",
    ".ant-dropdown-menu-item",
    "[class*='menu-item']",
    "[class*='option']",
  ];
  const primarySelectors = [
    ...overlayOptionSelectors,
    ".ant-select-selection-item",
    ".ant-radio-wrapper",
    ".ant-tabs-tab",
  ];
  const fallbackSelectors = ["a", "button", "li", "span", "div"];
  const collect = (selectors) => {
    const all = [];
    const seen = new Set();
    for (const sel of selectors) {
      try {
        for (const node of Array.from(document.querySelectorAll(sel))) {
          if (!node || seen.has(node)) continue;
          seen.add(node);
          all.push(node);
        }
      } catch (_) {}
    }
    return all;
  };
  const allOverlayRoots = Array.from(
    document.querySelectorAll(
      ".ant-select-dropdown,.ant-dropdown,.ant-popover,.ant-cascader-menus,[role='listbox'],[role='menu'],[class*='dropdown'],[class*='popover']"
    )
  ).filter(isVisible);

  const collectFromRoots = (roots, selectors) => {
    const all = [];
    const seen = new Set();
    for (const root of roots) {
      for (const sel of selectors) {
        try {
          for (const node of Array.from(root.querySelectorAll(sel))) {
            if (!node || seen.has(node)) continue;
            seen.add(node);
            all.push(node);
          }
        } catch (_) {}
      }
    }
    return all;
  };

  const targetLabelSoft = softNorm(targetLabel || "");
  const targetLabelNoColon = targetLabelSoft.replace(/[:：]/g, "");
  const targetNorm = norm(target);
  const targetSoft = softNorm(target);
  let targetNodes = targetLabel
    ? Array.from(document.querySelectorAll("label,span,div,li,th,td"))
        .map((node) => {
          if (!isVisible(node)) return null;
          const rawText = clean(node.innerText || node.textContent || "");
          const tn = norm(rawText);
          const ts = softNorm(rawText);
          if (!tn || !ts) return null;
          const tsNoColon = ts.replace(/[:：]/g, "");
          if (!(ts.includes(targetLabelSoft) || targetLabelSoft.includes(ts))) return null;
          if (targetLabelNoColon.length <= 4 && tsNoColon !== targetLabelNoColon) return null;
          if (rawText.length > Math.max(64, String(targetLabel || "").length * 4 + 16)) return null;
          const tag = String(node.tagName || "").toLowerCase();
          if ((tag === "div" || tag === "li") && rawText.length > Math.max(28, String(targetLabel || "").length * 3 + 8)) {
            const cls = clean(node.className || "").toLowerCase();
            if (!/(label|title|name|field|caption|item|selector|select|form|text)/.test(cls)) return null;
          }
          let score = 0;
          if (ts === targetLabelSoft || tn === norm(targetLabel || "")) score += 12000;
          else if (ts.startsWith(targetLabelSoft)) score += 3000;
          else score += 1200;
          if (canUseLast) {
            const d = distanceToLast(node, last);
            score += Math.max(0, 700 - d);
          }
          score -= rawText.length * 2;
          return { node, score };
        })
        .filter(Boolean)
        .sort((a, b) => b.score - a.score)
        .slice(0, 12)
        .map((x) => x.node)
    : [];
  const viewportArea = Math.max(1, window.innerWidth * window.innerHeight);
  const rootSet = new Set();
  const localRoots = [];
  for (const node of targetNodes) {
    const root = node.closest(".ant-form-item,.ant-row,.ant-col,.ant-space-item,.ant-card,section,article,div,li,td");
    if (!root || rootSet.has(root)) continue;
    const rr = root.getBoundingClientRect();
    if (!rr || rr.width <= 0 || rr.height <= 0) continue;
    const areaRatio = (rr.width * rr.height) / viewportArea;
    if (areaRatio > 0.72) continue;
    rootSet.add(root);
    localRoots.push(root);
  }

  let overlayRoots = allOverlayRoots;
  if (overlayRoots.length && canUseLast) {
    const ranked = overlayRoots
      .map((root) => ({ root, d: distanceToLast(root, last) }))
      .sort((a, b) => a.d - b.d);
    const near = ranked.filter((x) => x.d <= 700).map((x) => x.root);
    overlayRoots = near.length ? near.slice(0, 3) : [];
  }
  // If near-last filtering misses (common for cascader popups rendered far from trigger),
  // fall back to all visible overlays instead of failing hard.
  if (!overlayRoots.length && allOverlayRoots.length) {
    overlayRoots = allOverlayRoots.slice(0, 6);
  }
  if (overlayRoots.length && localRoots.length) {
    const tied = overlayRoots.filter((ov) =>
      localRoots.some((lr) => {
        if (!lr || !ov) return false;
        try {
          return lr.contains(ov) || ov.contains(lr);
        } catch (_) {
          return false;
        }
      })
    );
    if (tied.length) overlayRoots = tied;
  }

  let candidates = [];
  const scopedSelectors = primarySelectors.concat(fallbackSelectors);
  const strictOverlaySelectors = overlayOptionSelectors;
  const isShortValue = targetSoft.length <= 4;
  if (overlayRoots.length) {
    candidates = collectFromRoots(overlayRoots, strictOverlaySelectors);
    if (!candidates.length) {
      candidates = collectFromRoots(overlayRoots, scopedSelectors);
    }
  }
  if (!candidates.length && localRoots.length) {
    candidates = collectFromRoots(localRoots, strictOverlaySelectors);
    if (!candidates.length) {
      candidates = collectFromRoots(localRoots, scopedSelectors);
    }
  }
  // Target-scoped fallback: still prefer overlays, but remove strict proximity/tied limits.
  if (!candidates.length && targetLabel && allOverlayRoots.length) {
    candidates = collectFromRoots(allOverlayRoots, strictOverlaySelectors);
    if (!candidates.length) {
      candidates = collectFromRoots(allOverlayRoots, scopedSelectors);
    }
  }
  if (!candidates.length && !targetLabel) candidates = collect(primarySelectors);
  if (!candidates.length && !targetLabel) candidates = collect(fallbackSelectors);
  const blockedNorm = blockTextTokens.map((v) => norm(v)).filter(Boolean);
  const isBlockedTarget = blockedNorm.some((tok) => targetNorm.includes(tok) || tok.includes(targetNorm));
  const shortTarget = targetSoft.length <= 2;

  const matched = candidates
    .map((node) => {
      const t = clean(node.innerText || node.textContent || "");
      if (!t || !isVisible(node)) return null;
      if (t.length > Math.max(84, target.length * 6 + 24)) return null;
      const tag = String(node.tagName || "").toLowerCase();
      if (tag === "body" || tag === "html") return null;
      const tn = norm(t);
      const ts = softNorm(t);
      const exact = ts === targetSoft || tn === targetNorm;
      const starts = ts.startsWith(targetSoft);
      const contains = ts.includes(targetSoft);
      if (!(exact || starts || contains)) return null;
      if (isShortValue && !node.closest(".ant-select-dropdown,.ant-dropdown,.ant-popover,.ant-cascader-menus,[role='listbox'],[role='menu']")) {
        // Short value path items like "抖音"/"全部" must come from an opened option overlay.
        return null;
      }
      // Do not over-prune short CJK values (e.g. "投后", "全部").
      // We keep locality constraints in the shortTarget block below.
      if (!isBlockedTarget && blockedNorm.some((tok) => tn.includes(tok))) return null;

      if (avoidGlobalNav) {
        const inGlobalNav = isLikelyGlobalNav(node);
        if (inGlobalNav) return null;
      }

      const diff = Math.abs(ts.length - targetSoft.length);
      let score = -t.length - diff * 16;
      if (exact) score += 12000;
      else if (starts) score += 2600;
      else score += 1200;
      if (node.closest(".ant-select-item-option,[role='option']")) score += 600;
      if (overlayRoots.length && node.closest(".ant-select-dropdown,.ant-dropdown,.ant-popover,[role='listbox'],[role='menu']")) score += 900;
      if (localRoots.length && node.closest(".ant-form-item,.ant-row,.ant-col,.ant-space-item,.ant-card,section,article,div,li,td")) score += 350;
      if (canUseLast) {
        const r = node.getBoundingClientRect();
        const cx = r.left + r.width / 2;
        const cy = r.top + r.height / 2;
        const dx = cx - Number(last.x || 0);
        const dy = cy - Number(last.y || 0);
        const d = Math.sqrt(dx * dx + dy * dy);
        score += Math.max(0, 650 - d);
      }
      if (shortTarget) {
        // For short options like "全部", require stronger locality.
        const nearLast = canUseLast
          ? (() => {
              const r = node.getBoundingClientRect();
              const cx = r.left + r.width / 2;
              const cy = r.top + r.height / 2;
              const dx = cx - Number(last.x || 0);
              const dy = cy - Number(last.y || 0);
              return Math.sqrt(dx * dx + dy * dy) <= 320;
            })()
          : false;
        const inOverlay = !!node.closest(".ant-select-dropdown,.ant-dropdown,.ant-popover,.ant-cascader-menus,[role='listbox'],[role='menu']");
        const inLocalRoot = !!node.closest(".ant-form-item,.ant-row,.ant-col,.ant-space-item,.ant-card,section,article,div,li,td");
        if (!(inOverlay || nearLast || inLocalRoot)) return null;
      }
      return { node, text: t, score };
    })
    .filter(Boolean)
    .sort((a, b) => b.score - a.score);

  if (!matched.length) {
    out.error = "option_not_found";
    return out;
  }

  const best = matched[0];
  const clickSameTab = (node) => {
    const clickable =
      node?.closest("a,button,[role='button'],li,.ant-btn,div,span") || node;
    if (!clickable) return false;
    const anchor =
      clickable.tagName === "A" ? clickable : clickable.closest("a[href]");
    if (anchor) {
      try {
        anchor.setAttribute("target", "_self");
        anchor.removeAttribute("rel");
      } catch (_) {}
      try {
        anchor.dispatchEvent(
          new MouseEvent("click", {
            bubbles: true,
            cancelable: true,
            button: 0,
            ctrlKey: false,
            metaKey: false,
            shiftKey: false,
            altKey: false,
          })
        );
        return true;
      } catch (_) {}
    }
    clickable.click();
    return true;
  };
  try {
    best.node.scrollIntoView({ block: "center", inline: "center" });
    clickSameTab(best.node);
    const r = best.node.getBoundingClientRect();
    window.__opt_last_interaction = {
      kind: "select",
      text: clean(best.text || ""),
      x: r.left + r.width / 2,
      y: r.top + r.height / 2,
      ts: Date.now(),
      url: `${location.pathname}${location.search}`,
    };
    out.selected = true;
    out.text = best.text;
    return out;
  } catch (e) {
    out.error = String(e);
    return out;
  }
}
"""

    raw = await page.evaluate(
        js,
        {
            "option_text": option_text,
            "target_label": target_label or "",
            "avoid_global_nav": avoid_global_nav,
            "block_text_tokens": block_text_tokens or [],
        },
    )
    parsed = _json_loads_if_possible(raw)
    if not isinstance(parsed, dict):
        return {"selected": False, "text": None, "error": f"invalid_js_result:{raw}"}
    return parsed


async def _select_option_in_open_overlays(
    browser_session: BrowserSession,
    option_text: str,
    *,
    block_text_tokens: list[str] | None = None,
) -> dict:
    """
    Generic option picker for opened select/cascader overlays.
    This does not depend on a specific label and is robust for multi-level path selections.
    """
    page = await browser_session.must_get_current_page()
    js = r"""
(input) => {
  const target = (input?.option_text || "").trim();
  const blockTextTokens = Array.isArray(input?.block_text_tokens) ? input.block_text_tokens : [];
  const out = { selected: false, text: null, selector: null, error: null };
  if (!target) {
    out.error = "empty_option_text";
    return out;
  }

  const clean = (v) => String(v ?? "").replace(/\s+/g, " ").trim();
  const norm = (v) => clean(v).toLowerCase();
  const softNorm = (v) =>
    norm(v)
      .replace(/[\.。]/g, "")
      .replace(/[\s_\-—–·:：\/\\|（）()\[\]【】]+/g, "");
  const isVisible = (el) => {
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };

  const targetNorm = norm(target);
  const targetSoft = softNorm(target);
  const blockedNorm = blockTextTokens.map((v) => norm(v)).filter(Boolean);
  const shortTarget = targetSoft.length <= 2;

  const overlays = Array.from(
    document.querySelectorAll(
      ".ant-select-dropdown,.ant-dropdown,.ant-popover,.ant-cascader-menus,[role='listbox'],[role='menu'],[role='tree'],[class*='dropdown'],[class*='popover']"
    )
  ).filter(isVisible);

  if (!overlays.length) {
    out.error = "option_overlay_not_visible";
    return out;
  }

  const last = window.__opt_last_interaction || null;
  const canUseLast =
    !!last &&
    Date.now() - Number(last.ts || 0) <= 20000 &&
    String(last.url || "") === `${location.pathname}${location.search}` &&
    Number.isFinite(Number(last.x)) &&
    Number.isFinite(Number(last.y));

  const optionSelectors = [
    ".ant-select-item-option-content",
    ".ant-select-item-option",
    "[role='option']",
    ".ant-cascader-menu-item",
    ".ant-cascader-option",
    ".ant-cascader-menu-item-content",
    ".ant-dropdown-menu-item",
    "[role='menuitem']",
    "[role='treeitem']",
    "[class*='menu-item']",
    "[class*='option']",
    "li",
    "div",
    "span",
  ];

  const candidates = [];
  const seen = new Set();
  for (const root of overlays) {
    for (const sel of optionSelectors) {
      try {
        for (const node of Array.from(root.querySelectorAll(sel))) {
          if (!node || seen.has(node)) continue;
          seen.add(node);
          if (!isVisible(node)) continue;
          const raw = clean(node.innerText || node.textContent || "");
          if (!raw || raw.length > 120) continue;
          const tn = norm(raw);
          const ts = softNorm(raw);
          if (!tn || !ts) continue;
          if (blockedNorm.some((tok) => tn.includes(tok))) continue;

          const exact = ts === targetSoft || tn === targetNorm;
          const starts = ts.startsWith(targetSoft);
          const contains = ts.includes(targetSoft);
          if (!(exact || starts || contains)) continue;
          if (shortTarget && !contains) continue;

          let score = 0;
          if (exact) score += 12000;
          else if (starts) score += 2800;
          else score += 1200;
          score -= raw.length * 2;

          const cls = clean(node.className || "").toLowerCase();
          if (/option|menu-item|cascader|select/.test(cls)) score += 450;

          if (canUseLast) {
            try {
              const r = node.getBoundingClientRect();
              const cx = r.left + r.width / 2;
              const cy = r.top + r.height / 2;
              const dx = cx - Number(last.x || 0);
              const dy = cy - Number(last.y || 0);
              const d = Math.sqrt(dx * dx + dy * dy);
              score += Math.max(0, 700 - d);
            } catch (_) {}
          }
          candidates.push({ node, raw, score, selector: sel });
        }
      } catch (_) {}
    }
  }

  if (!candidates.length) {
    out.error = "option_not_found";
    return out;
  }

  candidates.sort((a, b) => b.score - a.score);
  const best = candidates[0];
  const clickSameTab = (node) => {
    const clickable =
      node?.closest("a,button,[role='option'],[role='menuitem'],[role='treeitem'],li,div,span") || node;
    if (!clickable) return false;
    const anchor = clickable.tagName === "A" ? clickable : clickable.closest("a[href]");
    if (anchor) {
      try {
        anchor.setAttribute("target", "_self");
        anchor.removeAttribute("rel");
      } catch (_) {}
      try {
        anchor.dispatchEvent(
          new MouseEvent("click", {
            bubbles: true,
            cancelable: true,
            button: 0,
            ctrlKey: false,
            metaKey: false,
            shiftKey: false,
            altKey: false,
          })
        );
        return true;
      } catch (_) {}
    }
    clickable.click();
    return true;
  };

  try {
    best.node.scrollIntoView({ block: "center", inline: "center" });
    clickSameTab(best.node);
    const r = best.node.getBoundingClientRect();
    window.__opt_last_interaction = {
      kind: "select_overlay",
      text: clean(best.raw || ""),
      x: r.left + r.width / 2,
      y: r.top + r.height / 2,
      ts: Date.now(),
      url: `${location.pathname}${location.search}`,
    };
    out.selected = true;
    out.text = best.raw;
    out.selector = best.selector;
    return out;
  } catch (e) {
    out.error = String(e);
    return out;
  }
}
"""
    raw = await page.evaluate(
        js,
        {
            "option_text": option_text,
            "block_text_tokens": block_text_tokens or [],
        },
    )
    parsed = _json_loads_if_possible(raw)
    if not isinstance(parsed, dict):
        return {"selected": False, "text": None, "selector": None, "error": f"invalid_js_result:{raw}"}
    return parsed


async def _verify_selection_applied(browser_session: BrowserSession, target: str | None, value: str) -> dict:
    page = await browser_session.must_get_current_page()
    js = r"""
(input) => {
  const target = (input?.target || "").trim();
  const value = (input?.value || "").trim();
  const out = { applied: false, reason: null, evidence: null };
  if (!value) {
    out.reason = "empty_value";
    return out;
  }

  const clean = (v) => String(v ?? "").replace(/\s+/g, " ").trim();
  const norm = (v) => clean(v).toLowerCase();
  const targetNorm = norm(target);
  const valueNorm = norm(value);

  const isVisible = (el) => {
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  const textOf = (el) => clean(el?.innerText || el?.textContent || "");

  const selectedNodes = Array.from(
    document.querySelectorAll(
      ".ant-select-selection-item,.ant-select-item-option-selected,.ant-radio-wrapper-checked,.ant-tabs-tab-active,[aria-selected='true'],[aria-checked='true']"
    )
  ).filter(isVisible);
  const selectedHit = selectedNodes.find((node) => norm(textOf(node)).includes(valueNorm));

  if (targetNorm) {
    const targetNodes = Array.from(document.querySelectorAll("label,span,div,li,th,td"))
      .filter((node) => {
        if (!isVisible(node)) return false;
        const rawText = textOf(node);
        const tn = norm(rawText);
        if (!tn || !tn.includes(targetNorm)) return false;
        if (rawText.length > Math.max(64, target.length * 4 + 16)) return false;
        const tag = String(node.tagName || "").toLowerCase();
        if ((tag === "div" || tag === "li") && rawText.length > Math.max(28, target.length * 3 + 8)) {
          const cls = clean(node.className || "").toLowerCase();
          if (!/(label|title|name|field|caption|item|selector|select|form|text)/.test(cls)) return false;
        }
        return true;
      })
      .slice(0, 80);
    for (const node of targetNodes) {
      const container =
        node.closest(".ant-form-item,.ant-row,.ant-col,.ant-space-item,.ant-select,.ant-select-selector,.ant-card,section,article,div,li,td") ||
        node.parentElement ||
        document.body;
      if (!container) continue;

      const selectedInContainer = Array.from(
        container.querySelectorAll(
          ".ant-select-selection-item,.ant-select-item-option-selected,.ant-radio-wrapper-checked,.ant-tabs-tab-active,[aria-selected='true'],[aria-checked='true']"
        )
      )
        .filter(isVisible)
        .map((el) => textOf(el));
      if (selectedInContainer.some((txt) => norm(txt).includes(valueNorm))) {
        out.applied = true;
        out.reason = "target_container_selected_marker";
        out.evidence = selectedInContainer.join(" | ").slice(0, 220);
        return out;
      }

      const controlNodes = Array.from(
        container.querySelectorAll(
          ".ant-select-selector,.ant-select-selection-item,[role='combobox'],input[role='combobox'],input[type='search'],input[type='text']"
        )
      ).filter(isVisible);
      for (const control of controlNodes) {
        const controlText = clean(
          control.getAttribute?.("value") ||
            control.getAttribute?.("aria-label") ||
            control.getAttribute?.("placeholder") ||
            textOf(control)
        );
        if (norm(controlText).includes(valueNorm)) {
          out.applied = true;
          out.reason = "target_control_value_contains";
          out.evidence = controlText.slice(0, 220);
          return out;
        }
      }

      // Deliberately avoid loose "context contains value" to prevent false positives.
    }
  }

  if (selectedHit && !targetNorm) {
    out.applied = true;
    out.reason = "selected_node_contains_value";
    out.evidence = textOf(selectedHit);
    return out;
  }

  out.reason = "value_not_observed";
  return out;
}
"""
    raw = await page.evaluate(js, {"target": target or "", "value": value})
    parsed = _json_loads_if_possible(raw)
    if not isinstance(parsed, dict):
        return {"applied": False, "reason": f"invalid_js_result:{raw}", "evidence": None}
    return parsed


async def _fill_search_input(
    browser_session: BrowserSession,
    selectors: list[str],
    value: str,
    row_selectors: list[str] | None = None,
) -> dict:
    page = await browser_session.must_get_current_page()
    js = r"""
(input) => {
  const selectors = Array.isArray(input?.selectors) ? input.selectors : [];
  const value = (input?.value || "").toString();
  const rowSelectors = Array.isArray(input?.row_selectors) ? input.row_selectors : [];
  const out = { filled: false, selector: null, placeholder: null, submitted: false, submit_methods: [], error: null };

  const isVisible = (el) => {
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };

  const countRows = (root) => {
    const scope = root || document;
    const sels = rowSelectors.length ? rowSelectors : ["table tbody tr", ".ant-table-row", ".report-row", "[role='row']"];
    let total = 0;
    for (const sel of sels) {
      try {
        total += Array.from(scope.querySelectorAll(sel)).length;
      } catch (_) {}
    }
    return total;
  };

  const tableContainerOf = (node) => {
    const candidates = [
      ".ant-table-wrapper",
      ".ant-table",
      "[class*='table']",
      "[class*='list']",
      ".ant-card",
      "section",
      "main",
      "form",
      "div",
    ];
    for (const sel of candidates) {
      try {
        const found = node.closest(sel);
        if (found) return found;
      } catch (_) {}
    }
    return null;
  };

  const globalRows = countRows(document);
  const tries = selectors.length ? selectors : ["input[placeholder*='搜索']", "input[placeholder*='报告']", "input[type='search']", "input"];
  const allInputs = [];
  for (const sel of tries) {
    try {
      for (const node of Array.from(document.querySelectorAll(sel))) {
        if (!(node instanceof HTMLInputElement || node instanceof HTMLTextAreaElement)) continue;
        if (!isVisible(node)) continue;
        const ph = (node.getAttribute("placeholder") || "").trim();
        const container = tableContainerOf(node);
        const localRows = container ? countRows(container) : 0;
        let score = 0;
        if (/请输入报告名称\/报告ID|报告ID|回车搜索/i.test(ph)) score += 3600;
        if (/搜索|报告|report|search/i.test(ph)) score += 1000;
        if (/search/i.test(node.type || "")) score += 400;
        if (/报告|report/i.test(ph)) score += 500;
        if (/报告名称|报告id|report id/i.test(ph)) score += 2200;
        const containerText = (container?.innerText || "").slice(0, 900).toLowerCase();
        if (/升级版报告|全部报告|结案报告/.test(containerText)) score += 1200;
        if (/报告列表|报告名称|报告id|计算状态|计算周期/.test(containerText)) score += 1600;
        if (globalRows > 0) {
          score += Math.min(localRows, 40) * 60;
          if (localRows === 0) score -= 2800;
        }
        const hasSearchBtn = !!(container && container.querySelector(".ant-input-search-button,button[aria-label*='搜索'],button[title*='搜索'],.anticon-search"));
        if (hasSearchBtn) score += 280;
        const rect = node.getBoundingClientRect();
        if (rect && rect.top < 160 && localRows === 0) score -= 1200;
        score -= sel.length;
        allInputs.push({ node, sel, score, ph, localRows });
      }
    } catch (_) {}
  }
  allInputs.sort((a, b) => b.score - a.score);

  for (const item of allInputs) {
    try {
      const node = item.node;
      node.scrollIntoView({ block: "center", inline: "nearest" });
      node.focus();
      const setNativeValue = (el, val) => {
        const proto = Object.getPrototypeOf(el);
        const desc = proto ? Object.getOwnPropertyDescriptor(proto, "value") : null;
        if (desc && typeof desc.set === "function") desc.set.call(el, val);
        else el.value = val;
      };
      setNativeValue(node, "");
      node.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "deleteContentBackward", data: null }));
      node.dispatchEvent(new Event("change", { bubbles: true }));
      setNativeValue(node, value);
      node.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText", data: value }));
      node.dispatchEvent(new Event("change", { bubbles: true }));
      let submitted = false;
      const nearRoot =
        node.closest(".ant-input-search,.ant-input-group,.ant-space,.ant-row,form,section,main,div") || document;
      const searchBtn = nearRoot.querySelector(".ant-input-search-button,button[aria-label*='搜索'],button[title*='搜索'],button[type='submit']");
      if (searchBtn && isVisible(searchBtn)) {
        try {
          searchBtn.click();
          out.submit_methods.push("search_button_click");
          submitted = true;
        } catch (_) {}
      }
      // Always press Enter like human operation, even if search button exists.
      try {
        const keyOpts = { key: "Enter", code: "Enter", keyCode: 13, which: 13, bubbles: true, cancelable: true };
        node.dispatchEvent(new KeyboardEvent("keydown", keyOpts));
        node.dispatchEvent(new KeyboardEvent("keypress", keyOpts));
        node.dispatchEvent(new KeyboardEvent("keyup", keyOpts));
        out.submit_methods.push("enter_key");
        submitted = true;
      } catch (_) {}
      // Some forms only react on submit.
      try {
        if (node.form) {
          if (typeof node.form.requestSubmit === "function") node.form.requestSubmit();
          else node.form.dispatchEvent(new Event("submit", { bubbles: true, cancelable: true }));
          out.submit_methods.push("form_submit");
          submitted = true;
        }
      } catch (_) {}
      if (!submitted) {
        out.error = "search_submit_not_triggered";
        return out;
      }
      out.filled = true;
      out.selector = item.sel;
      out.placeholder = item.ph || null;
      out.submitted = submitted;
      return out;
    } catch (_) {}
  }

  out.error = "search_input_not_found";
  return out;
}
"""

    raw = await page.evaluate(js, {"selectors": selectors, "value": value, "row_selectors": row_selectors or []})
    parsed = _json_loads_if_possible(raw)
    if not isinstance(parsed, dict):
        return {"filled": False, "selector": None, "error": f"invalid_js_result:{raw}"}
    return parsed


async def _click_report_by_name(
    browser_session: BrowserSession,
    report_name: str,
    row_selectors: list[str],
    name_cell_selectors: list[str],
    view_button_selectors: list[str],
    require_exact: bool = False,
) -> dict:
    page = await browser_session.must_get_current_page()
    js = r"""
(input) => {
  const reportName = (input?.report_name || "").trim();
  const rowSelectors = Array.isArray(input?.row_selectors) ? input.row_selectors : [];
  const nameCellSelectors = Array.isArray(input?.name_cell_selectors) ? input.name_cell_selectors : [];
  const viewButtonSelectors = Array.isArray(input?.view_button_selectors) ? input.view_button_selectors : [];
  const requireExact = Boolean(input?.require_exact);

  const out = { clicked: false, matched_report: null, clicked_selector: null, match_mode: null, error: null };
  if (!reportName) {
    out.error = "empty_report_name";
    return out;
  }

  const clean = (v) => String(v ?? "").replace(/\s+/g, " ").trim();
  const stripTailPunct = (v) => clean(v).replace(/[\s\.,;:!?，。；：！？、·…）)】\]】]+$/g, "").trim();
  const norm = (v) => stripTailPunct(v).toLowerCase();
  const softNorm = (v) =>
    norm(v)
      .replace(/[\.。]/g, "")
      .replace(/[\s_\-—–·:：\/\\|（）()\[\]【】]+/g, "");

  const target = norm(reportName);
  const targetSoft = softNorm(reportName);
  const targetTokens = clean(target)
    .split(/[_\-\s]+/)
    .map((x) => clean(x).toLowerCase())
    .filter((x) => x.length >= 2);
  const targetNumTokens = Array.from(target.matchAll(/\d+/g)).map((m) => (m?.[0] || "").trim()).filter(Boolean);

  const isVisible = (el) => {
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };

  const rows = [];
  const pushRows = (nodes) => {
    for (const node of nodes) {
      if (!node || rows.includes(node)) continue;
      rows.push(node);
    }
  };

  for (const sel of rowSelectors) {
    try {
      pushRows(Array.from(document.querySelectorAll(sel)));
    } catch (_) {}
  }

  if (!rows.length) {
    pushRows(Array.from(document.querySelectorAll("table tbody tr,.ant-table-row,.report-row,[role='row']")));
  }

  const textOf = (node) => clean(node?.innerText || node?.textContent || "");
  const collectVisibleRowNames = () => {
    const names = [];
    const seen = new Set();
    for (const row of rows) {
      if (!isVisible(row)) continue;
      let nameText = "";
      for (const sel of nameCellSelectors) {
        try {
          const node = row.querySelector(sel);
          if (!node) continue;
          const t = textOf(node);
          if (t) {
            nameText = t;
            break;
          }
        } catch (_) {}
      }
      if (!nameText) nameText = textOf(row);
      nameText = clean((nameText || "").split("\n")[0] || "");
      if (!nameText || nameText.length > 120) continue;
      if (seen.has(nameText)) continue;
      seen.add(nameText);
      names.push(nameText);
      if (names.length >= 10) break;
    }
    return names;
  };

  const scoreRow = (row) => {
    if (!isVisible(row)) return null;
    let nameText = "";
    for (const sel of nameCellSelectors) {
      try {
        const node = row.querySelector(sel);
        if (node) {
          const t = textOf(node);
          if (t) {
            nameText = t;
            break;
          }
        }
      } catch (_) {}
    }

    if (!nameText) nameText = textOf(row);
    const canonicalName = stripTailPunct(nameText);
    const n = norm(canonicalName);
    const ns = softNorm(canonicalName);
    if (!n || !ns) return null;

    const tokenHitCount = targetTokens.filter((tk) => n.includes(tk)).length;
    const tokenCoverage = targetTokens.length ? tokenHitCount / targetTokens.length : 0;
    const containSoft = ns.includes(targetSoft) || targetSoft.includes(ns);
    const containNormal = n.includes(target) || target.includes(n);
    const rowNumTokens = Array.from(nameText.matchAll(/\d+/g)).map((m) => (m?.[0] || "").trim()).filter(Boolean);
    const numHitCount = targetNumTokens.filter((tk) => rowNumTokens.includes(tk)).length;
    const numCoverage = targetNumTokens.length ? numHitCount / targetNumTokens.length : 1;

    if (!containSoft && !containNormal && tokenCoverage < 0.5) return null;

    const exact = n === target || ns === targetSoft;
    if (!exact && targetNumTokens.length > 0 && numCoverage < 1) return null;
    let score = (exact ? 12000 : 0) - nameText.length;
    if (containSoft) score += 2600;
    if (containNormal) score += 1600;
    score += Math.round(tokenCoverage * 2000);
    score += Math.round(numCoverage * 1800);

    return { row, nameText, score, exact };
  };

  const matched = rows.map(scoreRow).filter(Boolean).sort((a, b) => b.score - a.score);
  if (!matched.length) {
    out.error = "report_row_not_found";
    out.candidate_reports = collectVisibleRowNames();
    return out;
  }

  const exactMatched = matched.filter((x) => x.exact);
  if (requireExact && !exactMatched.length) {
    out.error = "report_row_not_found_exact";
    out.candidate_reports = collectVisibleRowNames();
    return out;
  }
  const best = exactMatched.length ? exactMatched[0] : matched[0];

  const chooseViewNode = (row) => {
    for (const sel of viewButtonSelectors) {
      try {
        const nodes = Array.from(row.querySelectorAll(sel));
        for (const node of nodes) {
          if (!isVisible(node)) continue;
          const t = textOf(node);
          if (/查看报告|查看|进入|详情|report|open/i.test(t)) return { node, selector: sel };
        }
      } catch (_) {}
    }

    const fallback = Array.from(row.querySelectorAll("button,a,[role='button']"))
      .find((node) => {
        if (!isVisible(node)) return false;
        const t = textOf(node);
        return /查看报告|查看|进入|详情|report|open/i.test(t);
      });

    if (fallback) return { node: fallback, selector: "fallback:view-button" };
    return { node: row, selector: "fallback:row" };
  };

  const clickSameTab = (node) => {
    const clickable =
      node?.closest("a,button,[role='button'],li,.ant-btn,div,span") || node;
    if (!clickable) return false;
    const anchor =
      clickable.tagName === "A" ? clickable : clickable.closest("a[href]");
    if (anchor) {
      try {
        anchor.setAttribute("target", "_self");
        anchor.removeAttribute("rel");
      } catch (_) {}
      try {
        anchor.dispatchEvent(
          new MouseEvent("click", {
            bubbles: true,
            cancelable: true,
            button: 0,
            ctrlKey: false,
            metaKey: false,
            shiftKey: false,
            altKey: false,
          })
        );
        return true;
      } catch (_) {}
    }
    clickable.click();
    return true;
  };

  const pick = chooseViewNode(best.row);
  try {
    pick.node.scrollIntoView({ block: "center", inline: "center" });
    clickSameTab(pick.node);
    out.clicked = true;
    out.matched_report = best.nameText;
    out.clicked_selector = pick.selector;
    out.match_mode = best.exact ? "exact" : "fuzzy";
    return out;
  } catch (e) {
    out.error = String(e);
    return out;
  }
}
"""

    raw = await page.evaluate(
        js,
        {
            "report_name": report_name,
            "row_selectors": row_selectors,
            "name_cell_selectors": name_cell_selectors,
            "view_button_selectors": view_button_selectors,
            "require_exact": require_exact,
        },
    )
    parsed = _json_loads_if_possible(raw)
    if not isinstance(parsed, dict):
        return {"clicked": False, "matched_report": None, "clicked_selector": None, "error": f"invalid_js_result:{raw}"}
    return parsed


async def _wait_report_rows_after_search(
    browser_session: BrowserSession,
    report_name: str,
    row_selectors: list[str],
    name_cell_selectors: list[str],
    timeout_seconds: int,
) -> dict:
    page = await browser_session.must_get_current_page()
    js = r"""
(input) => {
  const reportName = (input?.report_name || "").trim();
  const rowSelectors = Array.isArray(input?.row_selectors) ? input.row_selectors : [];
  const nameCellSelectors = Array.isArray(input?.name_cell_selectors) ? input.name_cell_selectors : [];

  const clean = (v) => String(v ?? "").replace(/\s+/g, " ").trim();
  const stripTailPunct = (v) => clean(v).replace(/[\s\.,;:!?，。；：！？、·…）)】\]】]+$/g, "").trim();
  const norm = (v) => stripTailPunct(v).toLowerCase();
  const softNorm = (v) =>
    norm(v)
      .replace(/[\.。]/g, "")
      .replace(/[\s_\-—–·:：\/\\|（）()\[\]【】]+/g, "");

  const target = norm(reportName);
  const targetSoft = softNorm(reportName);
  const targetTokens = clean(target)
    .split(/[_\-\s]+/)
    .map((x) => clean(x).toLowerCase())
    .filter((x) => x.length >= 2);
  const targetNumTokens = Array.from(target.matchAll(/\d+/g)).map((m) => (m?.[0] || "").trim()).filter(Boolean);

  const isVisible = (el) => {
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  const textOf = (node) => clean(node?.innerText || node?.textContent || "");

  const rows = [];
  const pushRows = (nodes) => {
    for (const node of nodes) {
      if (!node || rows.includes(node)) continue;
      rows.push(node);
    }
  };
  for (const sel of rowSelectors) {
    try {
      pushRows(Array.from(document.querySelectorAll(sel)));
    } catch (_) {}
  }
  if (!rows.length) {
    pushRows(Array.from(document.querySelectorAll("table tbody tr,.ant-table-row,.report-row,[role='row']")));
  }

  const names = [];
  const seen = new Set();
  let matchedExact = false;
  let matchedSoft = false;
  for (const row of rows) {
    if (!isVisible(row)) continue;
    let nameText = "";
    for (const sel of nameCellSelectors) {
      try {
        const node = row.querySelector(sel);
        if (!node) continue;
        const t = textOf(node);
        if (t) {
          nameText = t;
          break;
        }
      } catch (_) {}
    }
    if (!nameText) nameText = textOf(row);
    nameText = clean((nameText || "").split("\n")[0] || "");
    if (!nameText || nameText.length > 180) continue;
    if (!seen.has(nameText)) {
      seen.add(nameText);
      names.push(nameText);
    }

    const n = norm(nameText);
    const ns = softNorm(nameText);
    const exact = n === target || ns === targetSoft;
    if (exact) matchedExact = true;
    const containSoft = ns.includes(targetSoft) || targetSoft.includes(ns);
    const tokenHitCount = targetTokens.filter((tk) => n.includes(tk)).length;
    const tokenCoverage = targetTokens.length ? tokenHitCount / targetTokens.length : 0;
    const rowNumTokens = Array.from(nameText.matchAll(/\d+/g)).map((m) => (m?.[0] || "").trim()).filter(Boolean);
    const numHitCount = targetNumTokens.filter((tk) => rowNumTokens.includes(tk)).length;
    const numCoverage = targetNumTokens.length ? numHitCount / targetNumTokens.length : 1;
    if ((containSoft || tokenCoverage >= 0.5) && numCoverage >= 1) matchedSoft = true;
  }

  return {
    row_count: names.length,
    candidate_reports: names.slice(0, 12),
    matched_exact: matchedExact,
    matched_soft: matchedSoft,
  };
}
"""

    deadline = time.time() + max(1, min(timeout_seconds, 30))
    last_row_count = -1
    stable_cycles = 0
    last_snapshot: dict = {"row_count": 0, "candidate_reports": [], "matched_exact": False, "matched_soft": False}

    while time.time() < deadline:
        raw = await page.evaluate(
            js,
            {
                "report_name": report_name,
                "row_selectors": row_selectors,
                "name_cell_selectors": name_cell_selectors,
            },
        )
        snapshot = _json_loads_if_possible(raw)
        if not isinstance(snapshot, dict):
            snapshot = {"row_count": 0, "candidate_reports": [], "matched_exact": False, "matched_soft": False}
        last_snapshot = snapshot
        row_count = int(snapshot.get("row_count") or 0)
        if snapshot.get("matched_exact") or snapshot.get("matched_soft"):
            break
        if row_count > 0 and row_count == last_row_count:
            stable_cycles += 1
        else:
            stable_cycles = 0
        last_row_count = row_count
        if row_count > 0 and stable_cycles >= 2:
            break
        await asyncio.sleep(0.35)

    return last_snapshot


def _search_snapshot_related(report_name: str, snapshot: dict | None) -> bool:
    if not report_name or not isinstance(snapshot, dict):
        return False
    matched_exact = bool(snapshot.get("matched_exact"))
    matched_soft = bool(snapshot.get("matched_soft"))
    if matched_exact or matched_soft:
        return True
    candidates = snapshot.get("candidate_reports") or []
    if not isinstance(candidates, list):
        return False
    target = re.sub(r"[\s\.,;:!?，。；：！？、·…）)】\]]+$", "", report_name).strip().lower()
    if not target:
        return False
    target_soft = re.sub(r"[\.。]", "", target)
    target_soft = re.sub(r"[\s_\-—–·:：/\\|（）()\[\]【】]+", "", target_soft)
    target_nums = re.findall(r"\d+", target)
    for item in candidates:
        text = str(item or "").strip().lower()
        if not text:
            continue
        text_soft = re.sub(r"[\.。]", "", text)
        text_soft = re.sub(r"[\s_\-—–·:：/\\|（）()\[\]【】]+", "", text_soft)
        if target_soft and (target_soft in text_soft or text_soft in target_soft):
            if not target_nums:
                return True
            row_nums = re.findall(r"\d+", text)
            if all(num in row_nums for num in target_nums):
                return True
    return False


async def _is_report_list_page(browser_session: BrowserSession, profile) -> dict:
    page = await browser_session.must_get_current_page()
    js = r"""
(input) => {
  const rowSelectors = Array.isArray(input?.row_selectors) ? input.row_selectors : [];
  const clean = (v) => String(v ?? "").replace(/\s+/g, " ").trim();
  const body = clean((document.body?.innerText || "").slice(0, 4000));
  const bodyLower = body.toLowerCase();
  let rowCount = 0;
  for (const sel of rowSelectors) {
    try {
      rowCount += Array.from(document.querySelectorAll(sel)).length;
    } catch (_) {}
  }
  const hasListTokens = /报告列表|升级版报告|报告名称|报告id|查看报告|详情|新建报告/.test(bodyLower);
  const hasSearchInput = !!document.querySelector("input[placeholder*='报告'],input[placeholder*='搜索'],input[type='search']");
  const hasReportSearchHint = hasSearchInput && /报告|report/.test(bodyLower);
  const inList = rowCount > 0 || hasListTokens || hasReportSearchHint;
  return {
    in_list: inList,
    row_count: rowCount,
    has_list_tokens: hasListTokens,
    has_search_input: hasSearchInput,
    has_report_search_hint: hasReportSearchHint
  };
}
"""
    raw = await page.evaluate(js, {"row_selectors": getattr(profile, "report_row_selectors", [])})
    parsed = _json_loads_if_possible(raw)
    if not isinstance(parsed, dict):
        return {
            "in_list": False,
            "row_count": 0,
            "has_list_tokens": False,
            "has_search_input": False,
            "has_report_search_hint": False,
        }
    return parsed


async def _is_target_report_already_open(browser_session: BrowserSession, report_name: str) -> dict:
    hint = await _get_opened_report_hint(browser_session)
    target_norm = _normalize_report_name_for_compare(report_name)
    hint_norm = _normalize_report_name_for_compare(hint)
    matched = bool(target_norm and hint_norm and (target_norm in hint_norm or hint_norm in target_norm))
    return {"matched": matched, "opened_hint": hint}


async def _resubmit_report_search(browser_session: BrowserSession, value: str) -> dict:
    page = await browser_session.must_get_current_page()
    js = r"""
(input) => {
  const value = (input?.value || "").toString();
  const out = { submitted: false, selector: null, placeholder: null, submit_methods: [], error: null };
  if (!value) {
    out.error = "empty_value";
    return out;
  }

  const clean = (v) => String(v ?? "").replace(/\s+/g, " ").trim();
  const isVisible = (el) => {
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };

  const inputs = Array.from(document.querySelectorAll("input[placeholder*='报告'],input[placeholder*='搜索'],input[type='search'],input"))
    .filter((node) => node instanceof HTMLInputElement && isVisible(node));

  const scored = inputs
    .map((node) => {
      const ph = clean(node.getAttribute("placeholder") || "");
      const container =
        node.closest(".ant-input-search,.ant-input-group,.ant-table-wrapper,.ant-table,.report-list,section,main,div,form") || node.parentElement;
      const text = clean((container?.innerText || "").slice(0, 900)).toLowerCase();
      let score = 0;
      if (/报告名称|报告id|请输入报告名称/.test(ph)) score += 2200;
      if (/报告|search|搜索/.test(ph)) score += 1200;
      if (/报告列表|升级版报告|报告名称|报告id|计算状态|计算周期/.test(text)) score += 1800;
      const hasBtn = !!(container && container.querySelector(".ant-input-search-button,button[aria-label*='搜索'],button[title*='搜索'],button[type='submit']"));
      if (hasBtn) score += 300;
      return { node, ph, container, score };
    })
    .sort((a, b) => b.score - a.score);

  if (!scored.length) {
    out.error = "search_input_not_found";
    return out;
  }

  const setNativeValue = (el, val) => {
    const proto = Object.getPrototypeOf(el);
    const desc = proto ? Object.getOwnPropertyDescriptor(proto, "value") : null;
    if (desc && typeof desc.set === "function") desc.set.call(el, val);
    else el.value = val;
  };

  const best = scored[0];
  try {
    best.node.scrollIntoView({ block: "center", inline: "nearest" });
    best.node.focus();
    setNativeValue(best.node, "");
    best.node.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "deleteContentBackward", data: null }));
    setNativeValue(best.node, value);
    best.node.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText", data: value }));
    best.node.dispatchEvent(new Event("change", { bubbles: true }));

    let submitted = false;
    const btn = best.container?.querySelector(".ant-input-search-button,button[aria-label*='搜索'],button[title*='搜索'],button[type='submit']");
    if (btn && isVisible(btn)) {
      try {
        btn.click();
        out.submit_methods.push("search_button_click");
        submitted = true;
      } catch (_) {}
    }
    try {
      const keyOpts = { key: "Enter", code: "Enter", keyCode: 13, which: 13, bubbles: true, cancelable: true };
      best.node.dispatchEvent(new KeyboardEvent("keydown", keyOpts));
      best.node.dispatchEvent(new KeyboardEvent("keypress", keyOpts));
      best.node.dispatchEvent(new KeyboardEvent("keyup", keyOpts));
      out.submit_methods.push("enter_key");
      submitted = true;
    } catch (_) {}
    try {
      if (best.node.form) {
        if (typeof best.node.form.requestSubmit === "function") best.node.form.requestSubmit();
        else best.node.form.dispatchEvent(new Event("submit", { bubbles: true, cancelable: true }));
        out.submit_methods.push("form_submit");
        submitted = true;
      }
    } catch (_) {}
    if (!submitted) {
      out.error = "search_submit_not_triggered";
      return out;
    }

    out.submitted = submitted;
    out.selector = "fallback:resubmit_search_input";
    out.placeholder = best.ph || null;
    return out;
  } catch (e) {
    out.error = String(e);
    return out;
  }
}
"""
    raw = await page.evaluate(js, {"value": value})
    parsed = _json_loads_if_possible(raw)
    if not isinstance(parsed, dict):
        return {"submitted": False, "selector": None, "placeholder": None, "error": f"invalid_js_result:{raw}"}
    return parsed


async def _open_report_once(
    browser_session: BrowserSession,
    profile,
    report_name: str,
    wait_seconds: int,
) -> dict:
    await _ensure_single_tab_guard(browser_session)
    normalized_report_name = re.sub(r"[\s\.,;:!?，。；：！？、·…）)】\]]+$", "", (report_name or "").strip())
    report_name_for_match = normalized_report_name or report_name

    already_open = await _is_target_report_already_open(browser_session, report_name_for_match)
    if already_open.get("matched"):
        return {
            "search": {"filled": False, "selector": None, "error": "already_opened_target_report"},
            "click": {
                "clicked": True,
                "matched_report": already_open.get("opened_hint"),
                "clicked_selector": None,
                "match_mode": "already_opened",
                "error": None,
            },
            "ok": True,
        }

    list_state = await _is_report_list_page(browser_session, profile)
    if not list_state.get("in_list"):
        # Self-heal: try returning to report list before hard failing.
        recover_nav = await _navigate_to_report_list_best_effort(browser_session, profile, rounds=2)
        list_state_after = await _is_report_list_page(browser_session, profile)
        if recover_nav.get("ok") or list_state_after.get("in_list"):
            list_state = list_state_after
        else:
            # If list detection failed but page still contains direct report controls,
            # try a direct text fallback before returning hard failure.
            direct_try = await _open_report_by_name_text_direct(
                browser_session=browser_session,
                report_name=report_name_for_match,
            )
            if direct_try.get("clicked"):
                await asyncio.sleep(0.5)
                verify = await _validate_opened_report_and_recover(browser_session, report_name_for_match)
                if verify.get("matched"):
                    return {
                        "search": {
                            "filled": False,
                            "selector": None,
                            "error": "not_report_list_page_but_direct_opened",
                            "page_hint": already_open.get("opened_hint"),
                            "list_state": list_state,
                            "recover_nav": recover_nav,
                            "list_state_after_recover": list_state_after,
                        },
                        "click": direct_try,
                        "ok": True,
                    }
            return {
                "search": {
                    "filled": False,
                    "selector": None,
                    "error": "not_report_list_page",
                    "page_hint": already_open.get("opened_hint"),
                    "list_state": list_state,
                    "recover_nav": recover_nav,
                    "list_state_after_recover": list_state_after,
                },
                "click": {
                    "clicked": False,
                    "matched_report": None,
                    "clicked_selector": None,
                    "match_mode": None,
                    "error": "not_report_list_page",
                },
                "ok": False,
            }

    if not list_state.get("in_list"):
        direct_try = await _open_report_by_name_text_direct(
            browser_session=browser_session,
            report_name=report_name_for_match,
        )
        if direct_try.get("clicked"):
            await asyncio.sleep(0.5)
            verify = await _validate_opened_report_and_recover(browser_session, report_name_for_match)
            if verify.get("matched"):
                return {
                    "search": {
                        "filled": False,
                        "selector": None,
                        "error": "not_report_list_page_but_direct_opened",
                        "page_hint": already_open.get("opened_hint"),
                        "list_state": list_state,
                    },
                    "click": direct_try,
                    "ok": True,
                }
        return {
            "search": {
                "filled": False,
                "selector": None,
                "error": "not_report_list_page",
                "page_hint": already_open.get("opened_hint"),
                "list_state": list_state,
            },
            "click": {
                "clicked": False,
                "matched_report": None,
                "clicked_selector": None,
                "match_mode": None,
                "error": "not_report_list_page",
            },
            "ok": False,
        }

    search_result = await _fill_search_input(
        browser_session=browser_session,
        selectors=profile.report_search_selectors,
        value=report_name_for_match,
        row_selectors=profile.report_row_selectors,
    )
    search_wait = await _wait_report_rows_after_search(
        browser_session=browser_session,
        report_name=report_name_for_match,
        row_selectors=profile.report_row_selectors,
        name_cell_selectors=profile.report_name_cell_selectors,
        timeout_seconds=max(2, min(wait_seconds + 2, 20)),
    )
    search_result["row_wait"] = search_wait
    if search_result.get("filled") and not _search_snapshot_related(report_name_for_match, search_wait):
        retry_submit = await _resubmit_report_search(browser_session, report_name_for_match)
        retry_wait = await _wait_report_rows_after_search(
            browser_session=browser_session,
            report_name=report_name_for_match,
            row_selectors=profile.report_row_selectors,
            name_cell_selectors=profile.report_name_cell_selectors,
            timeout_seconds=max(2, min(wait_seconds + 2, 20)),
        )
        search_result["retry_submit"] = retry_submit
        search_result["retry_row_wait"] = retry_wait
    wait_budget = max(1, min(wait_seconds, 30))
    deadline = time.time() + wait_budget
    click_result = {
        "clicked": False,
        "matched_report": None,
        "clicked_selector": None,
        "match_mode": None,
        "error": "report_row_not_found_exact",
    }
    while time.time() < deadline:
        click_result = await _click_report_by_name(
            browser_session=browser_session,
            report_name=report_name_for_match,
            row_selectors=profile.report_row_selectors,
            name_cell_selectors=profile.report_name_cell_selectors,
            view_button_selectors=profile.report_view_button_selectors,
            require_exact=True,
        )
        if click_result.get("clicked"):
            break
        await asyncio.sleep(0.45)

    # Disable fuzzy row match to avoid opening wrong reports with similar prefixes/same date.

    if not click_result.get("clicked"):
        click_result = await _open_report_by_text_fallback(
            browser_session=browser_session,
            report_name=report_name_for_match,
            view_button_selectors=profile.report_view_button_selectors,
            require_exact=True,
        )

    if not click_result.get("clicked"):
        click_result = await _open_report_by_text_fallback(
            browser_session=browser_session,
            report_name=report_name_for_match,
            view_button_selectors=profile.report_view_button_selectors,
            require_exact=False,
        )

    if not click_result.get("clicked"):
        click_result = await _open_report_by_name_text_direct(
            browser_session=browser_session,
            report_name=report_name_for_match,
        )

    if click_result.get("clicked"):
        await asyncio.sleep(0.6)
        verify = await _validate_opened_report_and_recover(browser_session, report_name_for_match)
        if not verify.get("matched"):
            click_result = {
                "clicked": False,
                "matched_report": click_result.get("matched_report"),
                "clicked_selector": click_result.get("clicked_selector"),
                "match_mode": click_result.get("match_mode"),
                "error": "wrong_report_opened",
                "opened_hint": verify.get("opened_hint"),
                "back_ok": verify.get("back_ok"),
            }

    return {
        "search": search_result,
        "click": click_result,
        "ok": bool(click_result.get("clicked")),
    }


async def _manual_open_report_like(
    browser_session: BrowserSession,
    report_name: str,
    row_selectors: list[str],
    name_cell_selectors: list[str],
    view_button_selectors: list[str],
) -> dict:
    page = await browser_session.must_get_current_page()
    js = r"""
(input) => {
  const reportName = (input?.report_name || "").trim();
  const rowSelectors = Array.isArray(input?.row_selectors) ? input.row_selectors : [];
  const nameCellSelectors = Array.isArray(input?.name_cell_selectors) ? input.name_cell_selectors : [];
  const viewButtonSelectors = Array.isArray(input?.view_button_selectors) ? input.view_button_selectors : [];
  const out = { clicked: false, matched_report: null, clicked_selector: null, match_mode: "fallback_manual_like", error: null };
  if (!reportName) {
    out.error = "empty_report_name";
    return out;
  }

  const clean = (v) => String(v ?? "").replace(/\s+/g, " ").trim();
  const stripTailPunct = (v) => clean(v).replace(/[\s\.,;:!?，。；：！？、·…）)】\]】]+$/g, "").trim();
  const norm = (v) => stripTailPunct(v).toLowerCase();
  const softNorm = (v) =>
    norm(v)
      .replace(/[\.。]/g, "")
      .replace(/[\s_\-—–·:：\/\\|（）()\[\]【】]+/g, "");
  const target = norm(reportName);
  const targetSoft = softNorm(reportName);
  const targetTokens = clean(target)
    .split(/[_\-\s]+/)
    .map((x) => clean(x).toLowerCase())
    .filter((x) => x.length >= 2);
  const targetNumTokens = Array.from(target.matchAll(/\d+/g)).map((m) => (m?.[0] || "").trim()).filter(Boolean);

  const isVisible = (el) => {
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  const textOf = (node) => clean(node?.innerText || node?.textContent || "");

  const rows = [];
  const pushRows = (nodes) => {
    for (const node of nodes) {
      if (!node || rows.includes(node)) continue;
      rows.push(node);
    }
  };
  for (const sel of rowSelectors) {
    try {
      pushRows(Array.from(document.querySelectorAll(sel)));
    } catch (_) {}
  }
  if (!rows.length) {
    pushRows(Array.from(document.querySelectorAll("table tbody tr,.ant-table-row,.report-row,[role='row']")));
  }

  const collectVisibleRowNames = () => {
    const names = [];
    const seen = new Set();
    for (const row of rows) {
      if (!isVisible(row)) continue;
      let nameText = "";
      for (const sel of nameCellSelectors) {
        try {
          const node = row.querySelector(sel);
          if (!node) continue;
          const t = textOf(node);
          if (t) {
            nameText = t;
            break;
          }
        } catch (_) {}
      }
      if (!nameText) nameText = textOf(row);
      nameText = clean((nameText || "").split("\n")[0] || "");
      if (!nameText || nameText.length > 120) continue;
      if (seen.has(nameText)) continue;
      seen.add(nameText);
      names.push(nameText);
      if (names.length >= 10) break;
    }
    return names;
  };

  const scoreRow = (row) => {
    if (!isVisible(row)) return null;
    let nameText = "";
    for (const sel of nameCellSelectors) {
      try {
        const node = row.querySelector(sel);
        if (node) {
          const t = textOf(node);
          if (t) {
            nameText = t;
            break;
          }
        }
      } catch (_) {}
    }
    if (!nameText) nameText = textOf(row);
    const n = norm(nameText);
    const ns = softNorm(nameText);
    if (!n || !ns) return null;

    const tokenHitCount = targetTokens.filter((tk) => n.includes(tk)).length;
    const tokenCoverage = targetTokens.length ? tokenHitCount / targetTokens.length : 0;
    const containSoft = ns.includes(targetSoft) || targetSoft.includes(ns);
    const containNormal = n.includes(target) || target.includes(n);
    const rowNumTokens = Array.from(nameText.matchAll(/\d+/g)).map((m) => (m?.[0] || "").trim()).filter(Boolean);
    const numHitCount = targetNumTokens.filter((tk) => rowNumTokens.includes(tk)).length;
    const numCoverage = targetNumTokens.length ? numHitCount / targetNumTokens.length : 1;
    const exact = n === target || ns === targetSoft;

    if (!containSoft && !containNormal && tokenCoverage < 0.34) return null;
    if (targetNumTokens.length > 0 && numCoverage < 1) return null;

    let score = (exact ? 8000 : 0) - nameText.length;
    if (containSoft) score += 2600;
    if (containNormal) score += 1600;
    score += Math.round(tokenCoverage * 2000);
    score += Math.round(numCoverage * 1800);
    return { row, nameText, score };
  };

  const ranked = rows.map(scoreRow).filter(Boolean).sort((a, b) => b.score - a.score);
  if (!ranked.length) {
    out.error = "manual_row_not_found";
    out.candidate_reports = collectVisibleRowNames();
    return out;
  }

  const best = ranked[0];
  const pickClickNode = (row) => {
    for (const sel of viewButtonSelectors) {
      try {
        const nodes = Array.from(row.querySelectorAll(sel));
        for (const node of nodes) {
          if (!isVisible(node)) continue;
          const t = textOf(node);
          if (/查看报告|查看|进入|详情|report|open/i.test(t)) return { node, selector: sel };
        }
      } catch (_) {}
    }
    const hinted = row.querySelector("[class*='view'],[data-test*='view'],[data-testid*='view'],[class*='detail'],[data-test*='详情']");
    if (hinted && isVisible(hinted)) return { node: hinted, selector: "fallback:hinted-view-node" };
    const clickable = Array.from(row.querySelectorAll("a,button,[role='button'],.ant-btn,.link"))
      .find((node) => isVisible(node));
    if (clickable) return { node: clickable, selector: "fallback:first-clickable-in-row" };
    return { node: row, selector: "fallback:row" };
  };

  const picked = pickClickNode(best.row);
  const clickSameTab = (node) => {
    const clickable =
      node?.closest("a,button,[role='button'],li,.ant-btn,div,span") || node;
    if (!clickable) return false;
    const anchor =
      clickable.tagName === "A" ? clickable : clickable.closest("a[href]");
    if (anchor) {
      try {
        anchor.setAttribute("target", "_self");
        anchor.removeAttribute("rel");
      } catch (_) {}
      try {
        anchor.dispatchEvent(
          new MouseEvent("click", {
            bubbles: true,
            cancelable: true,
            button: 0,
            ctrlKey: false,
            metaKey: false,
            shiftKey: false,
            altKey: false,
          })
        );
        return true;
      } catch (_) {}
    }
    clickable.click();
    return true;
  };
  try {
    picked.node.scrollIntoView({ block: "center", inline: "center" });
    clickSameTab(picked.node);
    out.clicked = true;
    out.matched_report = best.nameText;
    out.clicked_selector = picked.selector;
    return out;
  } catch (e) {
    out.error = String(e);
    out.matched_report = best.nameText;
    out.clicked_selector = picked.selector;
    return out;
  }
}
"""
    raw = await page.evaluate(
        js,
        {
            "report_name": report_name,
            "row_selectors": row_selectors,
            "name_cell_selectors": name_cell_selectors,
            "view_button_selectors": view_button_selectors,
        },
    )
    parsed = _json_loads_if_possible(raw)
    if not isinstance(parsed, dict):
        return {"clicked": False, "matched_report": None, "clicked_selector": None, "error": f"invalid_js_result:{raw}"}
    return parsed


async def _open_report_by_text_fallback(
    browser_session: BrowserSession,
    report_name: str,
    view_button_selectors: list[str],
    require_exact: bool = True,
) -> dict:
    page = await browser_session.must_get_current_page()
    js = r"""
(input) => {
  const reportName = (input?.report_name || "").trim();
  const viewButtonSelectors = Array.isArray(input?.view_button_selectors) ? input.view_button_selectors : [];
  const requireExact = Boolean(input?.require_exact);
  const out = { clicked: false, matched_report: null, clicked_selector: null, match_mode: null, error: null };
  if (!reportName) {
    out.error = "empty_report_name";
    return out;
  }

  const clean = (v) => String(v ?? "").replace(/\s+/g, " ").trim();
  const stripTailPunct = (v) => clean(v).replace(/[\s\.,;:!?，。；：！？、·…）)】\]】]+$/g, "").trim();
  const norm = (v) => stripTailPunct(v).toLowerCase();
  const softNorm = (v) =>
    norm(v)
      .replace(/[\.。]/g, "")
      .replace(/[\s_\-—–·:：\/\\|（）()\[\]【】]+/g, "");

  const target = norm(reportName);
  const targetSoft = softNorm(reportName);
  const targetTokens = clean(target)
    .split(/[_\-\s]+/)
    .map((x) => clean(x).toLowerCase())
    .filter((x) => x.length >= 2);
  const targetNumTokens = Array.from(target.matchAll(/\d+/g)).map((m) => (m?.[0] || "").trim()).filter(Boolean);

  const isVisible = (el) => {
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  const textOf = (node) => clean(node?.innerText || node?.textContent || "");
  const normalizedLine = (text) => clean((text || "").split("\n")[0] || "");

  const candidates = [];
  const seen = new Set();
  const reportRoots = Array.from(
    document.querySelectorAll(".ant-table-wrapper,.ant-table,.report-list,[class*='report-list'],[class*='table'],[role='table']")
  )
    .filter((node) => {
      if (!isVisible(node)) return false;
      const t = clean((node.innerText || node.textContent || "").slice(0, 1400));
      if (!t || t.length < 20) return false;
      return /报告|report|升级版报告|报告名称|报告id|报告列表/i.test(t);
    })
    .slice(0, 24);
  const scopes = reportRoots.length ? reportRoots : [document];
  const probeSelectors = [
    "[data-test*='report-name']",
    "[data-testid*='report-name']",
    ".report-name",
    "table tbody tr td:first-child",
    ".ant-table-row td:first-child",
    ".ant-table-cell",
    "[title]",
  ];
  for (const scope of scopes) {
    for (const sel of probeSelectors) {
      try {
        const nodes = Array.from(scope.querySelectorAll(sel));
        for (const node of nodes) {
          if (!isVisible(node)) continue;
          let t = clean(node.getAttribute?.("title") || "");
          if (!t) t = textOf(node);
          t = normalizedLine(t);
          if (!t || t.length < 3 || t.length > 180) continue;
          if (seen.has(`${sel}::${t}`)) continue;
          seen.add(`${sel}::${t}`);
          candidates.push({ node, text: t, sel });
        }
      } catch (_) {}
    }
  }

  const scored = candidates
    .map((item) => {
      const n = norm(item.text);
      const ns = softNorm(item.text);
      if (!n || !ns) return null;
      const exact = n === target || ns === targetSoft;
      const containSoft = ns.includes(targetSoft) || targetSoft.includes(ns);
      const containNormal = n.includes(target) || target.includes(n);
      const tokenHitCount = targetTokens.filter((tk) => n.includes(tk)).length;
      const tokenCoverage = targetTokens.length ? tokenHitCount / targetTokens.length : 0;
      const rowNumTokens = Array.from(item.text.matchAll(/\d+/g)).map((m) => (m?.[0] || "").trim()).filter(Boolean);
      const numHitCount = targetNumTokens.filter((tk) => rowNumTokens.includes(tk)).length;
      const numCoverage = targetNumTokens.length ? numHitCount / targetNumTokens.length : 1;

      if (!containSoft && !containNormal && tokenCoverage < 0.5) return null;
      if (targetNumTokens.length > 0 && numCoverage < 1) return null;
      if (requireExact && !exact) return null;

      let score = (exact ? 12000 : 0) - item.text.length;
      if (containSoft) score += 2600;
      if (containNormal) score += 1600;
      score += Math.round(tokenCoverage * 2200);
      score += Math.round(numCoverage * 2600);
      if (/report|报告/.test(item.sel)) score += 400;
      return { ...item, exact, score };
    })
    .filter(Boolean)
    .sort((a, b) => b.score - a.score);

  if (!scored.length) {
    out.error = requireExact ? "report_text_not_found_exact" : "report_text_not_found";
    out.candidate_reports = candidates.slice(0, 12).map((x) => x.text);
    return out;
  }

  const best = scored[0];
  const row = best.node.closest("tr,.ant-table-row,[role='row'],.report-row");
  const clickableInRow = () => {
    if (!row) return null;
    for (const sel of viewButtonSelectors) {
      try {
        const nodes = Array.from(row.querySelectorAll(sel));
        for (const node of nodes) {
          if (!isVisible(node)) continue;
          const t = textOf(node);
          if (/查看报告|查看|进入|详情|report|open/i.test(t)) return { node, selector: sel };
        }
      } catch (_) {}
    }
    const fallback = Array.from(row.querySelectorAll("a,button,[role='button']"))
      .find((node) => {
        if (!isVisible(node)) return false;
        const t = textOf(node);
        return /查看报告|查看|进入|详情|report|open/i.test(t);
      });
    if (fallback) return { node: fallback, selector: "fallback:text-row-view-button" };
    return null;
  };

  const pick = clickableInRow() || { node: best.node, selector: "fallback:text-node" };
  const clickSameTab = (node) => {
    const clickable =
      node?.closest("a,button,[role='button'],li,.ant-btn,div,span") || node;
    if (!clickable) return false;
    const anchor =
      clickable.tagName === "A" ? clickable : clickable.closest("a[href]");
    if (anchor) {
      try {
        anchor.setAttribute("target", "_self");
        anchor.removeAttribute("rel");
      } catch (_) {}
      try {
        anchor.dispatchEvent(
          new MouseEvent("click", {
            bubbles: true,
            cancelable: true,
            button: 0,
            ctrlKey: false,
            metaKey: false,
            shiftKey: false,
            altKey: false,
          })
        );
        return true;
      } catch (_) {}
    }
    clickable.click();
    return true;
  };
  try {
    pick.node.scrollIntoView({ block: "center", inline: "center" });
    clickSameTab(pick.node);
    out.clicked = true;
    out.matched_report = best.text;
    out.clicked_selector = pick.selector;
    out.match_mode = best.exact ? "text_exact" : "text_fuzzy";
    return out;
  } catch (e) {
    out.error = String(e);
    out.matched_report = best.text;
    out.clicked_selector = pick.selector;
    out.match_mode = best.exact ? "text_exact" : "text_fuzzy";
    return out;
  }
}
"""

    raw = await page.evaluate(
        js,
        {
            "report_name": report_name,
            "view_button_selectors": view_button_selectors,
            "require_exact": require_exact,
        },
    )
    parsed = _json_loads_if_possible(raw)
    if not isinstance(parsed, dict):
        return {"clicked": False, "matched_report": None, "clicked_selector": None, "error": f"invalid_js_result:{raw}"}
    return parsed


def _report_name_variants(report_name: str) -> list[str]:
    raw = (report_name or "").strip()
    if not raw:
        return []
    base = re.sub(r"[\s\.,;:!?，。；：！？、·…）)】\]]+$", "", raw).strip()
    variants = [
        raw,
        base,
        f"{base}.",
        f"{base}。",
        base.replace(".", ""),
    ]
    dedup: list[str] = []
    seen = set()
    for v in variants:
        vv = (v or "").strip()
        if not vv or vv in seen:
            continue
        seen.add(vv)
        dedup.append(vv)
    return dedup


async def _open_report_by_name_text_direct(browser_session: BrowserSession, report_name: str) -> dict:
    page = await browser_session.must_get_current_page()
    js = r"""
(input) => {
  const reportName = (input?.report_name || "").trim();
  const variants = Array.isArray(input?.variants) ? input.variants : [];
  const out = { clicked: false, matched_report: null, clicked_selector: null, match_mode: "name_text_direct", error: null };
  if (!reportName) {
    out.error = "empty_report_name";
    return out;
  }
  const clean = (v) => String(v ?? "").replace(/\s+/g, " ").trim();
  const stripTailPunct = (v) => clean(v).replace(/[\s\.,;:!?，。；：！？、·…）)】\]】]+$/g, "").trim();
  const norm = (v) => stripTailPunct(v).toLowerCase();
  const softNorm = (v) =>
    norm(v)
      .replace(/[\.。]/g, "")
      .replace(/[\s_\-—–·:：\/\\|（）()\[\]【】]+/g, "");

  const target = norm(reportName);
  const targetSoft = softNorm(reportName);
  const targetTokens = target
    .split(/[_\-\s]+/)
    .map((x) => clean(x).toLowerCase())
    .filter((x) => x.length >= 2);
  const targetNumTokens = Array.from(target.matchAll(/\d+/g)).map((m) => (m?.[0] || "").trim()).filter(Boolean);

  const isVisible = (el) => {
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  const textOf = (node) => clean(node?.innerText || node?.textContent || "");

  const reportRoots = Array.from(
    document.querySelectorAll(".ant-table-wrapper,.ant-table,.report-list,[class*='report-list'],[class*='table'],section,main,div")
  )
    .filter((node) => {
      if (!isVisible(node)) return false;
      const t = clean((node.innerText || node.textContent || "").slice(0, 1400));
      if (!t || t.length < 20) return false;
      return /报告|report|升级版报告|报告名称|报告id|报告列表/i.test(t);
    })
    .slice(0, 24);
  const scopes = reportRoots.length ? reportRoots : [document];
  const selectors = [
    "[data-test*='report-name']",
    "[data-testid*='report-name']",
    ".report-name",
    ".ant-table-row td:nth-child(2)",
    ".ant-table-row td:first-child",
    "table tbody tr td:nth-child(2)",
    "table tbody tr td:first-child",
    ".ant-table-cell",
  ];
  const candidates = [];
  const seen = new Set();
  for (const scope of scopes) {
    for (const sel of selectors) {
      try {
        const nodes = Array.from(scope.querySelectorAll(sel));
        for (const node of nodes) {
          if (!isVisible(node)) continue;
          const title = clean(node.getAttribute?.("title") || "");
          const rawText = title || textOf(node);
          const t = clean((rawText || "").split("\n")[0] || "");
          if (!t || t.length < 3 || t.length > 180) continue;
          const key = `${sel}::${t}`;
          if (seen.has(key)) continue;
          seen.add(key);
          candidates.push({ node, sel, text: t });
        }
      } catch (_) {}
    }
  }

  const variantNorms = variants.map((v) => norm(v)).filter(Boolean);
  const variantSoftNorms = variants.map((v) => softNorm(v)).filter(Boolean);

  const scored = candidates.map((item) => {
    const n = norm(item.text);
    const ns = softNorm(item.text);
    if (!n || !ns) return null;
    const exact = n === target || ns === targetSoft || variantNorms.includes(n) || variantSoftNorms.includes(ns);
    const tokenHitCount = targetTokens.filter((tk) => n.includes(tk)).length;
    const tokenCoverage = targetTokens.length ? tokenHitCount / targetTokens.length : 0;
    const rowNumTokens = Array.from(item.text.matchAll(/\d+/g)).map((m) => (m?.[0] || "").trim()).filter(Boolean);
    const numHitCount = targetNumTokens.filter((tk) => rowNumTokens.includes(tk)).length;
    const numCoverage = targetNumTokens.length ? numHitCount / targetNumTokens.length : 1;
    if (!exact && tokenCoverage < 1) return null;
    if (targetNumTokens.length > 0 && numCoverage < 1) return null;
    let score = (exact ? 12000 : 0) + Math.round(tokenCoverage * 1800) + Math.round(numCoverage * 2400) - item.text.length;
    if (/report|报告/.test(item.sel)) score += 300;
    return { ...item, exact, score };
  }).filter(Boolean).sort((a, b) => b.score - a.score);

  if (!scored.length) {
    out.error = "report_text_direct_not_found";
    out.candidate_reports = candidates.slice(0, 10).map((x) => x.text);
    return out;
  }

  const best = scored[0];
  const clickSameTab = (node) => {
    const clickable =
      node?.closest("a,button,[role='button'],li,.ant-btn,div,span") || node;
    if (!clickable) return false;
    const anchor =
      clickable.tagName === "A" ? clickable : clickable.closest("a[href]");
    if (anchor) {
      try {
        anchor.setAttribute("target", "_self");
        anchor.removeAttribute("rel");
      } catch (_) {}
      try {
        anchor.dispatchEvent(
          new MouseEvent("click", {
            bubbles: true,
            cancelable: true,
            button: 0,
            ctrlKey: false,
            metaKey: false,
            shiftKey: false,
            altKey: false,
          })
        );
        return true;
      } catch (_) {}
    }
    clickable.click();
    return true;
  };
  try {
    best.node.scrollIntoView({ block: "center", inline: "center" });
    clickSameTab(best.node);
    out.clicked = true;
    out.matched_report = best.text;
    out.clicked_selector = best.sel;
    return out;
  } catch (e) {
    out.error = String(e);
    out.matched_report = best.text;
    out.clicked_selector = best.sel;
    return out;
  }
}
"""
    raw = await page.evaluate(js, {"report_name": report_name, "variants": _report_name_variants(report_name)})
    parsed = _json_loads_if_possible(raw)
    if isinstance(parsed, dict):
        return parsed
    return {
        "clicked": False,
        "matched_report": None,
        "clicked_selector": None,
        "match_mode": "name_text_direct",
        "error": f"invalid_js_result:{raw}",
    }


def _normalize_report_name_for_compare(value: str | None) -> str:
    text = (value or "").strip()
    if not text:
        return ""
    text = re.sub(r"[\s\.,;:!?，。；：！？、·…）)】\]]+$", "", text).strip().lower()
    text = re.sub(r"[\.。]", "", text)
    text = re.sub(r"[\s_\-—–·:：/\\|（）()\[\]【】]+", "", text)
    return text


async def _get_opened_report_hint(browser_session: BrowserSession) -> str | None:
    page = await browser_session.must_get_current_page()
    js = r"""
() => {
  const clean = (v) => String(v ?? "").replace(/\s+/g, " ").trim();
  const isVisible = (el) => {
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  const selectors = [
    "h1","h2",".ant-page-header-heading-title",".page-title",".title",
    ".ant-breadcrumb","[data-test*='report-name']","[data-testid*='report-name']"
  ];
  const candidates = [];
  for (const sel of selectors) {
    try {
      for (const node of Array.from(document.querySelectorAll(sel))) {
        if (!isVisible(node)) continue;
        const t = clean(node.innerText || node.textContent || "");
        if (!t || t.length < 3 || t.length > 180) continue;
        candidates.push(t);
      }
    } catch (_) {}
  }
  if (!candidates.length) return null;
  candidates.sort((a, b) => a.length - b.length);
  return candidates[0];
}
"""
    raw = await page.evaluate(js)
    parsed = _json_loads_if_possible(raw)
    if isinstance(parsed, str):
        return parsed.strip() or None
    if isinstance(parsed, dict) and isinstance(parsed.get("raw"), str):
        return parsed.get("raw").strip() or None
    return None


async def _validate_opened_report_and_recover(browser_session: BrowserSession, target_report_name: str) -> dict:
    hint = await _get_opened_report_hint(browser_session)
    target_norm = _normalize_report_name_for_compare(target_report_name)
    hint_norm = _normalize_report_name_for_compare(hint)
    if not hint_norm:
        # If page title/report hint is not detectable, do not force rollback.
        return {"matched": True, "opened_hint": hint, "note": "skip_validation_no_hint"}
    matched = bool(target_norm and hint_norm and (target_norm in hint_norm or hint_norm in target_norm))
    if matched:
        return {"matched": True, "opened_hint": hint}

    page = await browser_session.must_get_current_page()
    try:
        await page.go_back(timeout=30000)
        await asyncio.sleep(0.5)
        back_ok = True
    except Exception:
        back_ok = False
    return {"matched": False, "opened_hint": hint, "back_ok": back_ok}


async def _open_report_by_scroll_scan(
    browser_session: BrowserSession,
    report_name: str,
    row_selectors: list[str],
    name_cell_selectors: list[str],
    view_button_selectors: list[str],
    max_scrolls: int = 12,
) -> dict:
    """
    Scan virtualized/long report list by scrolling list container and retrying exact row match.
    """
    page = await browser_session.must_get_current_page()
    for idx in range(max(1, max_scrolls)):
        click_result = await _click_report_by_name(
            browser_session=browser_session,
            report_name=report_name,
            row_selectors=row_selectors,
            name_cell_selectors=name_cell_selectors,
            view_button_selectors=view_button_selectors,
            require_exact=True,
        )
        if click_result.get("clicked"):
            click_result["match_mode"] = "scroll_scan_exact"
            click_result["scan_step"] = idx + 1
            return click_result

        # Scroll the most likely list container (table body / virtual list) first, then page fallback.
        try:
            await page.evaluate(
                r"""
() => {
  const isScrollable = (el) => !!el && (el.scrollHeight - el.clientHeight > 20);
  const candidates = Array.from(document.querySelectorAll(
    ".ant-table-body,.ant-table-content,.ant-table-wrapper,[class*='virtual'],[class*='list'],main,section,div"
  ));
  let best = null;
  let bestScore = -1;
  for (const node of candidates) {
    if (!isScrollable(node)) continue;
    const txt = ((node.innerText || "").slice(0, 600) || "").toLowerCase();
    let score = 0;
    if (/报告|report|结案/.test(txt)) score += 1000;
    if (node.className && /table|list|virtual|body/i.test(String(node.className))) score += 800;
    if (score > bestScore) {
      best = node;
      bestScore = score;
    }
  }
  if (best) {
    best.scrollBy({ top: Math.max(320, Math.floor(best.clientHeight * 0.75)), behavior: "instant" });
    return "list_scrolled";
  }
  window.scrollBy({ top: 650, behavior: "instant" });
  return "window_scrolled";
}
""",
            )
        except Exception:
            pass
        await asyncio.sleep(0.3)

    return {
        "clicked": False,
        "matched_report": None,
        "clicked_selector": None,
        "match_mode": "scroll_scan_exact",
        "error": "report_row_not_found_after_scroll_scan",
    }


async def _collect_report_page_evidence(
    browser_session: BrowserSession,
    row_selectors: list[str],
    name_cell_selectors: list[str],
) -> dict:
    page = await browser_session.must_get_current_page()
    js = r"""
(input) => {
  const rowSelectors = Array.isArray(input?.row_selectors) ? input.row_selectors : [];
  const nameCellSelectors = Array.isArray(input?.name_cell_selectors) ? input.name_cell_selectors : [];
  const clean = (v) => String(v ?? "").replace(/\s+/g, " ").trim();
  const isVisible = (el) => {
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  const textOf = (node) => clean(node?.innerText || node?.textContent || "");

  const rows = [];
  const pushRows = (nodes) => {
    for (const node of nodes) {
      if (!node || rows.includes(node)) continue;
      rows.push(node);
    }
  };
  for (const sel of rowSelectors) {
    try {
      pushRows(Array.from(document.querySelectorAll(sel)));
    } catch (_) {}
  }
  if (!rows.length) {
    pushRows(Array.from(document.querySelectorAll("table tbody tr,.ant-table-row,.report-row,[role='row']")));
  }

  const candidateReports = [];
  const seenNames = new Set();
  for (const row of rows) {
    if (!isVisible(row)) continue;
    let nameText = "";
    for (const sel of nameCellSelectors) {
      try {
        const node = row.querySelector(sel);
        if (!node) continue;
        const t = textOf(node);
        if (t) {
          nameText = t;
          break;
        }
      } catch (_) {}
    }
    if (!nameText) nameText = textOf(row);
    nameText = clean((nameText || "").split("\n")[0] || "");
    if (!nameText || nameText.length > 120) continue;
    if (seenNames.has(nameText)) continue;
    seenNames.add(nameText);
    candidateReports.push(nameText);
    if (candidateReports.length >= 12) break;
  }

  const tabNodes = Array.from(document.querySelectorAll("[role='tab'],.ant-tabs-tab,button,a,span,li"));
  const visibleTabs = [];
  const seenTabs = new Set();
  for (const node of tabNodes) {
    if (!isVisible(node)) continue;
    const t = textOf(node);
    if (!t || t.length > 24) continue;
    if (!/报告|历史|升级|全部/.test(t)) continue;
    if (seenTabs.has(t)) continue;
    seenTabs.add(t);
    visibleTabs.push(t);
    if (visibleTabs.length >= 10) break;
  }

  return {
    page_url: String(location.href || ""),
    candidate_reports: candidateReports,
    visible_tabs: visibleTabs,
  };
}
"""
    raw = await page.evaluate(
        js,
        {
            "row_selectors": row_selectors,
            "name_cell_selectors": name_cell_selectors,
        },
    )
    parsed = _json_loads_if_possible(raw)
    if not isinstance(parsed, dict):
        return {"page_url": None, "candidate_reports": [], "visible_tabs": [], "error": f"invalid_js_result:{raw}"}
    return parsed


def _derive_collection_status(fields: list[dict], files: list[dict]) -> str:
    statuses = [item.get("status") for item in fields + files]
    if not statuses:
        return "failed"

    ok_count = sum(1 for item in statuses if item == "ok")
    if ok_count == len(statuses):
        return "success"
    if ok_count > 0:
        return "partial"
    return "failed"


async def _snapshot_page_state(browser_session: BrowserSession) -> dict:
    page = await browser_session.must_get_current_page()
    js = r"""
() => {
  const clean = (v) => String(v ?? "").replace(/\s+/g, " ").trim();
  const url = String(location.href || "");
  const title = clean(document.title || "");
  const body = clean((document.body?.innerText || "").slice(0, 4000));
  const merged = `${title} ${body}`.toLowerCase();

  const isLogin =
    /account\/login|passport|login/.test(url.toLowerCase()) ||
    /登录|手机号登录|验证码登录|扫码登录|请登录/.test(merged);

  const menuNodes = Array.from(
    document.querySelectorAll(
      "[role='menuitem'],.ant-menu-item,.ant-menu-title-content,.ant-menu-submenu-title,.ant-tabs-tab-btn,[role='tab']"
    )
  );
  const menuTexts = menuNodes
    .map((n) => clean(n.innerText || n.textContent || ""))
    .filter((t) => t && t.length <= 24)
    .slice(0, 120);
  const workbenchMenuHitCount = menuTexts.filter((t) => /营销决策|营销触点|营销概览|结案报告|投后结案|报告列表|人群分析/.test(t)).length;
  const urlLooksWorkbench = /\/report|\/post|\/marketing|\/decision|\/insight/.test(url.toLowerCase());
  const hasWorkbenchToken = /营销决策|营销触点|营销概览|结案报告|投后结案|报告列表|人群分析/.test(merged);
  const isWorkbench = !isLogin && (workbenchMenuHitCount > 0 || (urlLooksWorkbench && hasWorkbenchToken));
  const isLanding = !isLogin && !isWorkbench && /云图|oceanengine|巨量/.test(merged);

  return {
    url,
    title,
    snippet: body.slice(0, 220),
    is_login: isLogin,
    is_landing: isLanding,
    is_workbench: isWorkbench,
    workbench_menu_hit_count: workbenchMenuHitCount,
    menu_samples: menuTexts.slice(0, 8)
  };
}
"""
    raw = await page.evaluate(js)
    parsed = _json_loads_if_possible(raw)
    if not isinstance(parsed, dict):
        return {
            "url": None,
            "title": None,
            "snippet": None,
            "is_login": False,
            "is_landing": False,
            "is_workbench": False,
            "workbench_menu_hit_count": 0,
            "menu_samples": [],
        }
    return parsed


async def _try_enter_workbench(browser_session: BrowserSession) -> dict:
    trace: list[dict] = []
    selectors = ["button", "a", "[role='button']", "span", "div"]
    labels = ["进入工作台", "进入云图", "工作台", "去工作台", "立即进入", "进入"]
    for label in labels:
        clicked = await _click_text_step(browser_session=browser_session, label=label, selectors=selectors)
        ok = bool(clicked.get("clicked"))
        trace.append({"label": label, "clicked": ok, "selector": clicked.get("selector"), "error": clicked.get("error")})
        if ok:
            await asyncio.sleep(0.8)
            break
    return {"trace": trace}


def _menu_label_candidates(label: str, profile=None) -> list[str]:
    text = (label or "").strip()
    if not text:
        return []
    candidates: list[str] = [text]

    alias_map = _profile_menu_aliases(profile)
    candidates.extend(alias_map.get(text, []) or [])

    # Deduplicate while preserving order.
    dedup: list[str] = []
    seen = set()
    for item in candidates:
        v = (item or "").strip()
        if not v or v in seen:
            continue
        seen.add(v)
        dedup.append(v)
    return dedup


async def _navigate_to_report_list_best_effort(
    browser_session: BrowserSession,
    profile,
    *,
    rounds: int = 2,
) -> dict:
    """
    Best-effort recovery to reach report list page by using profile.menu_path.
    This helper is intentionally permissive and reused by both navigate_report_list
    action and open_report flow to avoid logic drift.
    """
    await _ensure_single_tab_guard(browser_session)
    trace: list[dict] = []
    strict_menu_selectors = [
        "header [role='menuitem']",
        "header .ant-menu-item",
        "header .ant-menu-title-content",
        "header .ant-menu-submenu-title",
        "nav [role='menuitem']",
        "nav .ant-menu-item",
        ".ant-layout-header [role='menuitem']",
        ".ant-layout-header .ant-menu-item",
        ".ant-layout-header .ant-menu-submenu-title",
        "aside [role='menuitem']",
        "aside .ant-menu-item",
        "aside .ant-menu-submenu-title",
        ".ant-menu-title-content",
        ".ant-dropdown-menu-item",
        ".ant-dropdown [role='menuitem']",
        ".ant-popover [role='menuitem']",
        "[role='menuitem']",
        ".ant-menu-item",
        ".ant-tabs-tab",
        "[role='tab']",
        ".ant-tabs-tab-btn",
    ]
    broad_menu_selectors = strict_menu_selectors + [
        "a",
        "button",
        "li",
        "div",
        "span",
    ]

    list_state = await _is_report_list_page(browser_session, profile)
    if list_state.get("in_list"):
        return {"ok": True, "trace": [], "list_state": list_state, "note": "already_on_report_list"}

    menu_path = _profile_menu_path(profile)
    total_steps = max(1, len(menu_path))
    nav_block_tokens = _profile_block_text_tokens(profile)
    for round_idx in range(1, max(1, rounds) + 1):
        for step_idx, label in enumerate(menu_path, start=1):
            candidates = _menu_label_candidates(label, profile)
            clicked = None
            step_ok = False
            attempt_logs: list[dict] = []

            # Keep parent menu open before trying submenu.
            if step_idx > 1:
                parent_candidates = _menu_label_candidates(menu_path[step_idx - 2], profile)
                for parent in parent_candidates:
                    parent_res = await _click_text_step(
                        browser_session=browser_session,
                        label=parent,
                        selectors=strict_menu_selectors,
                        allow_reverse_contains=False,
                        avoid_global_nav=False,
                        require_clickable=True,
                        max_text_len=120,
                        block_text_tokens=nav_block_tokens,
                    )
                    attempt_logs.append(
                        {
                            "round": round_idx,
                            "phase": "parent_preopen",
                            "candidate": parent,
                            "clicked": bool(parent_res.get("clicked")),
                            "selector": parent_res.get("selector"),
                            "error": parent_res.get("error"),
                        }
                    )
                    if parent_res.get("clicked"):
                        await asyncio.sleep(0.2)
                        break

            for attempt in range(1, 4):
                for candidate in candidates:
                    for phase_name, step_selectors in (
                        ("target_strict", strict_menu_selectors),
                        ("target_broad", broad_menu_selectors),
                    ):
                        res = await _click_text_step(
                            browser_session=browser_session,
                            label=candidate,
                            selectors=step_selectors,
                            allow_reverse_contains=False,
                            avoid_global_nav=False,
                            require_clickable=True,
                            max_text_len=120,
                            block_text_tokens=nav_block_tokens,
                        )
                        attempt_logs.append(
                            {
                                "round": round_idx,
                                "attempt": attempt,
                                "phase": phase_name,
                                "candidate": candidate,
                                "clicked": bool(res.get("clicked")),
                                "selector": res.get("selector"),
                                "error": res.get("error"),
                            }
                        )
                        if res.get("clicked"):
                            clicked = res
                            step_ok = True
                            break
                    if step_ok:
                        break
                if step_ok:
                    break
                await asyncio.sleep(0.4 if attempt < 3 else 0.2)

            await asyncio.sleep(0.5)
            list_state = await _is_report_list_page(browser_session, profile)
            on_report_list = bool(list_state.get("in_list"))
            trace.append(
                {
                    "round": round_idx,
                    "step_no": step_idx,
                    "target": label,
                    "candidates": candidates,
                    "clicked": step_ok,
                    "selector": (clicked or {}).get("selector"),
                    "error": (clicked or {}).get("error") if clicked else "menu_target_not_found",
                    "attempts": attempt_logs,
                    "report_list_after_step": on_report_list,
                }
            )
            if on_report_list:
                return {"ok": True, "trace": trace, "list_state": list_state}
            # Don't break early on a single step failure; next round may recover with refreshed DOM/menu.
            await asyncio.sleep(0.15)

        # end one round: small settle then check again.
        await asyncio.sleep(0.35)
        list_state = await _is_report_list_page(browser_session, profile)
        if list_state.get("in_list"):
            return {"ok": True, "trace": trace, "list_state": list_state}

    # Shortcut fallback: click report-list related labels directly.
    shortcut_labels: list[str] = []
    if menu_path:
        for item in menu_path:
            shortcut_labels.extend(_menu_label_candidates(item, profile))
    for tab in _profile_report_list_tabs(profile):
        shortcut_labels.extend(_menu_label_candidates(tab, profile))
    shortcut_labels.extend(_profile_report_shortcut_labels(profile))
    dedup_labels = list(dict.fromkeys([x for x in shortcut_labels if str(x or "").strip()]))

    shortcut_selectors = [
        "[role='tab']",
        ".ant-tabs-tab",
        ".ant-tabs-tab-btn",
        ".ant-menu-item",
        ".ant-menu-submenu-title",
        ".ant-dropdown-menu-item",
        ".ant-popover [role='menuitem']",
        "[role='tab']",
        "[role='menuitem']",
        "a",
        "button",
        "span",
        "div",
    ]
    shortcut_attempts: list[dict] = []
    for text in dedup_labels:
        clicked = await _click_text_step(
            browser_session=browser_session,
            label=text,
            selectors=shortcut_selectors,
            allow_reverse_contains=False,
            avoid_global_nav=False,
            require_clickable=True,
            max_text_len=120,
            block_text_tokens=nav_block_tokens,
        )
        shortcut_attempts.append(
            {
                "phase": "shortcut",
                "candidate": text,
                "clicked": bool(clicked.get("clicked")),
                "selector": clicked.get("selector"),
                "error": clicked.get("error"),
            }
        )
        if clicked.get("clicked"):
            await asyncio.sleep(0.45)
            list_state = await _is_report_list_page(browser_session, profile)
            if list_state.get("in_list"):
                trace.append(
                    {
                        "round": rounds + 1,
                        "step_no": 0,
                        "target": "shortcut_report_list",
                        "candidates": dedup_labels,
                        "clicked": True,
                        "selector": clicked.get("selector"),
                        "error": None,
                        "attempts": shortcut_attempts,
                        "report_list_after_step": True,
                    }
                )
                return {"ok": True, "trace": trace, "list_state": list_state}

    trace.append(
        {
            "round": rounds + 1,
            "step_no": 0,
            "target": "shortcut_report_list",
            "candidates": dedup_labels,
            "clicked": False,
            "selector": None,
            "error": "menu_target_not_found",
            "attempts": shortcut_attempts,
            "report_list_after_step": bool((list_state or {}).get("in_list")),
        }
    )
    return {"ok": False, "trace": trace, "list_state": list_state}


async def _ensure_single_tab_guard(browser_session: BrowserSession) -> None:
    page = await browser_session.must_get_current_page()
    js = r"""
() => {
  try {
    if (!window.__opt_single_tab_guard_installed) {
      window.__opt_single_tab_guard_installed = true;
      window.__opt_original_open = window.open;
      window.open = function(url, target, features) {
        try {
          if (url) {
            location.href = String(url);
          }
        } catch (_) {}
        return window;
      };
      const normalizeAnchor = (a) => {
        if (!a) return;
        try { a.target = "_self"; } catch (_) {}
        try { a.removeAttribute("rel"); } catch (_) {}
      };
      const normalizeAll = () => {
        for (const a of Array.from(document.querySelectorAll("a[target='_blank'],a[rel~='noopener'],a[rel~='noreferrer']"))) {
          normalizeAnchor(a);
        }
      };
      normalizeAll();
      document.addEventListener(
        "click",
        (e) => {
          const t = e && e.target && e.target.closest ? e.target.closest("a") : null;
          if (t) normalizeAnchor(t);
        },
        true
      );
      document.addEventListener(
        "mousedown",
        (e) => {
          const t = e && e.target && e.target.closest ? e.target.closest("a") : null;
          if (t) normalizeAnchor(t);
        },
        true
      );
    } else {
      for (const a of Array.from(document.querySelectorAll("a[target='_blank']"))) {
        try { a.target = "_self"; } catch (_) {}
      }
    }
    return { ok: true };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}
"""
    try:
        await page.evaluate(js)
    except Exception:
        # Non-fatal guard: do not block business flow
        return


def _context_get(path: str | None, context: dict) -> str | list[str] | None:
    if not path:
        return None
    raw = path.strip()
    if not raw:
        return None
    if not raw.startswith(("params.", "module.", "context.")):
        raw = f"context.{raw}"

    parts = raw.split(".")
    current = context
    for key in parts:
        if key in {"params", "module", "context"}:
            continue
        if not isinstance(current, dict) or key not in current:
            return None
        current = current.get(key)
    if isinstance(current, (str, list)):
        return current
    if current is None:
        return None
    return str(current)


def _resolve_step_value(step: InteractionStep, context: dict) -> str | None:
    value = _context_get(step.value_from, context)
    if isinstance(value, list):
        return None
    if isinstance(value, str) and value.strip():
        return value.strip()
    if step.value and step.value.strip():
        return step.value.strip()
    return None


def _split_select_values(raw_value: str) -> list[str]:
    text = (raw_value or "").strip()
    if not text:
        return []
    parts = [item.strip() for item in re.split(r"\s*(?:/|／|->|→)\s*", text) if item.strip()]
    if parts:
        deduped: list[str] = []
        for item in parts:
            if not deduped:
                deduped.append(item)
                continue
            if _normalize_match_text(item) == _normalize_match_text(deduped[-1]):
                continue
            deduped.append(item)
        if len(deduped) >= 2 and len({_normalize_match_text(v) for v in deduped}) == 1:
            return [deduped[0]]
        return deduped
    return [text]


def _resolve_step_select_values(step: InteractionStep, context: dict) -> list[str]:
    raw = _context_get(step.value_from, context)
    if isinstance(raw, list):
        return [str(item).strip() for item in raw if str(item).strip()]
    if isinstance(raw, str) and raw.strip():
        return _split_select_values(raw)
    if step.value and step.value.strip():
        return _split_select_values(step.value)
    return []


def _resolve_step_path(step: InteractionStep, context: dict) -> list[str]:
    values = _context_get(step.path_from, context)
    if isinstance(values, list):
        return [str(item).strip() for item in values if str(item).strip()]
    if step.path:
        return [item.strip() for item in step.path if item.strip()]
    return []


def _eval_when_expr(when: str | None, context: dict) -> bool:
    if not when:
        return True
    expr = when.strip()
    if not expr:
        return True

    m = re.fullmatch(r"([a-zA-Z0-9_\.]+)\s*(==|!=)\s*(.+)", expr)
    if m:
        left_path, op, right_raw = m.groups()
        left = _context_get(left_path, context)
        right_value = right_raw.strip().strip("'").strip('"')
        if right_raw.strip().lower() in {"null", "none"}:
            right = None
        elif right_raw.strip().lower() in {"true", "false"}:
            right = right_raw.strip().lower() == "true"
        else:
            right = right_value

        if op == "==":
            return left == right
        return left != right

    value = _context_get(expr, context)
    if isinstance(value, list):
        return len(value) > 0
    if isinstance(value, str):
        return bool(value.strip())
    return bool(value)


async def _is_text_visible(browser_session: BrowserSession, label: str) -> bool:
    page = await browser_session.must_get_current_page()
    js = r"""
(input) => {
  const label = (input || "").trim();
  if (!label) return false;
  const clean = (v) => String(v ?? "").replace(/\s+/g, " ").trim();
  const norm = (v) => clean(v).toLowerCase();
  const target = norm(label);
  const isVisible = (el) => {
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  const nodes = Array.from(document.querySelectorAll("body *"));
  for (const node of nodes) {
    if (!isVisible(node)) continue;
    const t = norm(node.innerText || node.textContent || "");
    if (!t) continue;
    if (t.includes(target) || target.includes(t)) return true;
  }
  return false;
}
"""
    raw = await page.evaluate(js, label)
    return str(raw).lower() == "true"


async def _detect_server_error_state(browser_session: BrowserSession) -> dict:
    page = await browser_session.must_get_current_page()
    js = r"""
() => {
  const title = (document.title || "").trim();
  const body = (document.body?.innerText || "").replace(/\s+/g, " ").slice(0, 2500).trim();
  const merged = `${title} ${body}`.toLowerCase();
  const patterns = [
    { code: "503", re: /(^|\D)503(\D|$)|service unavailable|服务不可用|系统繁忙|稍后重试/i },
    { code: "502", re: /(^|\D)502(\D|$)|bad gateway|网关错误/i },
    { code: "504", re: /(^|\D)504(\D|$)|gateway timeout|网关超时/i },
    { code: "RELOAD", re: /请点击重新加载|点击重新加载|重新加载页面|页面加载失败|网络异常请重试|请重试/i },
  ];
  for (const p of patterns) {
    if (p.re.test(merged)) {
      return {
        is_error: true,
        code: p.code,
        title,
        snippet: body.slice(0, 280),
      };
    }
  }
  return { is_error: false, code: null, title, snippet: body.slice(0, 280) };
}
"""
    raw = await page.evaluate(js)
    parsed = _json_loads_if_possible(raw)
    if not isinstance(parsed, dict):
        return {"is_error": False, "code": None, "title": None, "snippet": None}
    return parsed


async def _reload_current_page(browser_session: BrowserSession) -> None:
    page = await browser_session.must_get_current_page()
    try:
        await page.reload(timeout=45000)
    except Exception:
        try:
            await page.evaluate("() => { window.location.reload(); return true; }")
        except Exception:
            return
    await asyncio.sleep(0.75)


async def _recover_report_list_page(browser_session: BrowserSession) -> dict:
    selectors = ["button", "a", "[role='button']", "span", "div"]
    for label in ["重新加载", "立即重试", "重试", "刷新", "重新加载页面"]:
        clicked = await _click_text_step(
            browser_session=browser_session,
            label=label,
            selectors=selectors,
        )
        if clicked.get("clicked"):
            await asyncio.sleep(0.75)
            return {"recovered": True, "method": "click_label", "label": label, "click": clicked}

    await _reload_current_page(browser_session)
    return {"recovered": True, "method": "page_reload", "label": None, "click": None}


async def _detect_unexpected_global_page(browser_session: BrowserSession) -> dict:
    page = await browser_session.must_get_current_page()
    js = r"""
(input) => {
  const blockedTokens = Array.isArray(input?.blocked_tokens) ? input.blocked_tokens : [];
  const reportLikeTokens = Array.isArray(input?.report_like_tokens) ? input.report_like_tokens : [];
  const clean = (v) => String(v ?? "").replace(/\s+/g, " ").trim();
  const url = String(location.href || "");
  const title = clean(document.title || "");
  const body = clean((document.body?.innerText || "").slice(0, 2600));
  const merged = `${title} ${body}`.toLowerCase();

  const lowerUrl = url.toLowerCase();
  const isMessageRoute =
    /\/(message|notice|notification)(\/|$)/.test(lowerUrl) ||
    /#\/?(message|notice|notification)/.test(lowerUrl) ||
    /messagecenter|msg-?center/.test(lowerUrl);
  const hasBlockedLabel = blockedTokens.some((tok) => merged.includes(clean(tok).toLowerCase()));
  const hasReportLike = reportLikeTokens.some((tok) => merged.includes(clean(tok).toLowerCase()));

  return {
    is_unexpected: !!isMessageRoute || (!!hasBlockedLabel && !hasReportLike),
    url,
    title,
    snippet: body.slice(0, 220),
  };
}
"""
    generic_profile = get_profile_spec("generic")
    raw = await page.evaluate(
        js,
        {
            "blocked_tokens": _profile_block_text_tokens(generic_profile),
            "report_like_tokens": _profile_report_like_tokens(generic_profile),
        },
    )
    parsed = _json_loads_if_possible(raw)
    if not isinstance(parsed, dict):
        return {"is_unexpected": False, "url": None, "title": None, "snippet": None}
    return parsed


async def _guard_and_recover_unexpected_page(browser_session: BrowserSession) -> dict:
    before = await _detect_unexpected_global_page(browser_session)
    if not before.get("is_unexpected"):
        # Also treat load-error overlays as recoverable guard failures.
        page_error_before = await _detect_server_error_state(browser_session)
        if page_error_before.get("is_error"):
            recover = await _recover_report_list_page(browser_session)
            await asyncio.sleep(0.6)
            after_unexpected = await _detect_unexpected_global_page(browser_session)
            after_page_error = await _detect_server_error_state(browser_session)
            recovered = bool(recover.get("recovered")) and (not after_page_error.get("is_error"))
            return {
                "ok": recovered and (not after_unexpected.get("is_unexpected")),
                "recovered": recovered,
                "before": {"unexpected": before, "page_error": page_error_before},
                "after": {"unexpected": after_unexpected, "page_error": after_page_error, "recover": recover},
            }
        return {"ok": True, "recovered": False, "before": before, "after": before}

    page = await browser_session.must_get_current_page()
    back_ok = False
    try:
        await page.go_back(timeout=20000)
        back_ok = True
    except Exception:
        back_ok = False
    await asyncio.sleep(0.6)
    after = await _detect_unexpected_global_page(browser_session)
    recovered = bool(back_ok and not after.get("is_unexpected"))
    return {"ok": recovered, "recovered": recovered, "before": before, "after": after}


async def _fill_input_by_target(browser_session: BrowserSession, target: str, value: str) -> dict:
    page = await browser_session.must_get_current_page()
    js = r"""
(input) => {
  const target = (input?.target || "").trim();
  const value = (input?.value || "").toString();
  const out = { filled: false, selector: null, error: null };
  const clean = (v) => String(v ?? "").replace(/\s+/g, " ").trim();
  const norm = (v) => clean(v).toLowerCase();

  const isVisible = (el) => {
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };

  const labelNodes = Array.from(document.querySelectorAll("label,span,div,td,th,p"))
    .filter((node) => {
      const t = norm(node.innerText || node.textContent || "");
      return t && (t.includes(norm(target)) || norm(target).includes(t));
    });

  for (const node of labelNodes) {
    const container = node.closest("div,section,form,li,td,tr") || node.parentElement;
    if (!container) continue;
    const inputNode = container.querySelector("input,textarea");
    if (!inputNode || !isVisible(inputNode)) continue;
    inputNode.focus();
    inputNode.value = value;
    inputNode.dispatchEvent(new Event("input", { bubbles: true }));
    inputNode.dispatchEvent(new Event("change", { bubbles: true }));
    inputNode.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", code: "Enter", bubbles: true }));
    inputNode.dispatchEvent(new KeyboardEvent("keyup", { key: "Enter", code: "Enter", bubbles: true }));
    out.filled = true;
    out.selector = "near_label_input";
    return out;
  }

  out.error = "target_input_not_found";
  return out;
}
"""
    raw = await page.evaluate(js, {"target": target, "value": value})
    parsed = _json_loads_if_possible(raw)
    if not isinstance(parsed, dict):
        return {"filled": False, "selector": None, "error": f"invalid_js_result:{raw}"}
    return parsed


async def _set_select_value_by_target(browser_session: BrowserSession, target: str, value: str) -> dict:
    page = await browser_session.must_get_current_page()
    js = r"""
(input) => {
  const target = (input?.target || "").trim();
  const value = (input?.value || "").toString().trim();
  const out = { selected: false, typed: false, clicked_option: false, text: null, error: null };
  if (!target || !value) {
    out.error = "empty_target_or_value";
    return out;
  }

  const clean = (v) => String(v ?? "").replace(/\s+/g, " ").trim();
  const norm = (v) =>
    clean(v)
      .toLowerCase()
      .replace(/[\.。]/g, "")
      .replace(/[\s_\-—–·:：\/\\|（）()\[\]【】]+/g, "");
  const targetNorm = norm(target);
  const valueNorm = norm(value);

  const isVisible = (el) => {
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  const setNativeValue = (el, val) => {
    const proto = Object.getPrototypeOf(el);
    const desc = proto ? Object.getOwnPropertyDescriptor(proto, "value") : null;
    if (desc && typeof desc.set === "function") desc.set.call(el, val);
    else el.value = val;
  };

  const scoreLabelNode = (node) => {
    if (!node || !isVisible(node)) return null;
    const rawText = clean(node.innerText || node.textContent || "");
    const t = norm(rawText);
    if (!t) return null;
    if (rawText.length > Math.max(64, target.length * 4 + 16)) return null;
    const tag = String(node.tagName || "").toLowerCase();
    if (tag === "div" && rawText.length > Math.max(28, target.length * 3 + 8)) {
      const cls = clean(node.className || "").toLowerCase();
      if (!/(label|title|name|field|caption|item|selector|select|form|text)/.test(cls)) return null;
    }
    const tNoColon = t.replace(/[:：]/g, "");
    const targetNoColon = targetNorm.replace(/[:：]/g, "");
    const exact = tNoColon === targetNoColon;
    const starts = tNoColon.startsWith(targetNoColon);
    const contains = tNoColon.includes(targetNoColon);
    if (!(exact || starts || contains || targetNoColon.includes(tNoColon))) return null;
    let score = 0;
    if (exact) score += 12000;
    else if (starts) score += 3600;
    else if (contains) score += 900;
    else score += 200;
    const extra = Math.max(0, tNoColon.length - targetNoColon.length);
    score -= extra * 45;
    score -= rawText.length * 3;
    return { node, score };
  };

  const labels = Array.from(document.querySelectorAll("label,span,p,th,td,div"))
    .map((node) => scoreLabelNode(node))
    .filter(Boolean)
    .sort((a, b) => b.score - a.score)
    .slice(0, 16)
    .map((x) => x.node);

  const scoreInput = (inputNode, labelRect) => {
    const r = inputNode.getBoundingClientRect();
    const cx = r.left + r.width / 2;
    const cy = r.top + r.height / 2;
    const lcx = labelRect.left + labelRect.width / 2;
    const lcy = labelRect.top + labelRect.height / 2;
    const dx = cx - lcx;
    const dy = cy - lcy;
    let score = 0;
    const sameRow = Math.abs(dy) <= 84;
    const rightSide = dx >= -80 && dx <= 620;
    if (sameRow) score += 460;
    else score -= 380;
    if (rightSide) score += 260;
    else score -= 220;
    score -= Math.abs(dy) * 1.6;
    score -= Math.abs(dx) * 0.42;
    if (Math.abs(dy) > 220) score -= 1200;
    if (dx < -220 || dx > 920) score -= 900;
    const ph = clean(inputNode.getAttribute("placeholder") || "").toLowerCase();
    if (/请选择|搜索|请输入/.test(ph)) score += 140;
    return score;
  };

  const inputCandidates = [];
  const seenInputs = new Set();
  const viewportArea = Math.max(1, window.innerWidth * window.innerHeight);
  for (const ln of labels) {
    const lr = ln.getBoundingClientRect();
    const roots = [];
    const near = ln.closest(".ant-form-item,.ant-row,.ant-col,.ant-space-item,.ant-space-item-compact,.ant-select,.ant-cascader-picker,.ant-card,section,article,div,li,td");
    if (near) roots.push(near);
    if (ln.parentElement) roots.push(ln.parentElement);
    for (const root of roots) {
      if (!root) continue;
      if (root.getBoundingClientRect) {
        const rr = root.getBoundingClientRect();
        const ratio = (Math.max(1, rr.width) * Math.max(1, rr.height)) / viewportArea;
        if (ratio > 0.72) continue;
      }
      const nodes = Array.from(root.querySelectorAll(
        ".ant-select-selection-search input,.ant-select-selector input,[role='combobox'] input,input[placeholder*='请选择'],input[placeholder*='选择'],input[type='search'],input[type='text']"
      ));
      for (const node of nodes) {
        if (!node || !isVisible(node)) continue;
        if (seenInputs.has(node)) continue;
        seenInputs.add(node);
        inputCandidates.push({ node, score: scoreInput(node, lr) });
      }
    }
  }

  inputCandidates.sort((a, b) => b.score - a.score);
  const pickedInput = inputCandidates.length ? inputCandidates[0].node : null;
  if (!pickedInput) {
    out.error = "target_select_input_not_found";
    return out;
  }

  try {
    pickedInput.scrollIntoView({ block: "center", inline: "center" });
    const trigger =
      pickedInput.closest(".ant-select-selector,[role='combobox'],.ant-cascader-picker,.ant-cascader-picker-label,.ant-picker") ||
      pickedInput.parentElement;
    if (trigger && typeof trigger.click === "function") {
      try {
        trigger.click();
      } catch (_) {}
    }
    pickedInput.focus();
    setNativeValue(pickedInput, "");
    pickedInput.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "deleteContentBackward", data: null }));
    pickedInput.dispatchEvent(new Event("change", { bubbles: true }));
    setNativeValue(pickedInput, value);
    pickedInput.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText", data: value }));
    pickedInput.dispatchEvent(new Event("change", { bubbles: true }));
    out.typed = true;
  } catch (e) {
    out.error = String(e);
    return out;
  }

  const pr = pickedInput.getBoundingClientRect();
  const pcx = pr.left + pr.width / 2;
  const pcy = pr.top + pr.height / 2;
  const overlayRoots = Array.from(document.querySelectorAll(
    ".ant-select-dropdown,.ant-dropdown,.ant-popover,.ant-cascader-menus,[role='listbox'],[role='menu'],[class*='dropdown'],[class*='popover']"
  ))
    .filter(isVisible)
    .map((root) => {
      const rr = root.getBoundingClientRect();
      const rcx = rr.left + rr.width / 2;
      const rcy = rr.top + rr.height / 2;
      const dx = rcx - pcx;
      const dy = rcy - pcy;
      const d = Math.sqrt(dx * dx + dy * dy);
      return { root, d };
    })
    .sort((a, b) => a.d - b.d)
    .filter((x) => x.d <= 680)
    .slice(0, 3)
    .map((x) => x.root);

  const options = [];
  const optionSelectors = ".ant-select-item-option,.ant-select-item-option-content,[role='option'],.ant-dropdown-menu-item,.ant-cascader-menu-item,li,div,span";
  const pushOption = (node) => {
    if (!node || !isVisible(node)) return;
    options.push(node);
  };
  if (overlayRoots.length) {
    for (const root of overlayRoots) {
      for (const node of Array.from(root.querySelectorAll(optionSelectors))) pushOption(node);
    }
  }
  if (!options.length) {
    out.error = "target_option_overlay_not_found";
    return out;
  }
  const matched = options
    .map((node) => {
      const t = clean(node.innerText || node.textContent || "");
      const tn = norm(t);
      if (!tn) return null;
      const exact = tn === valueNorm;
      const starts = tn.startsWith(valueNorm);
      const contains = tn.includes(valueNorm);
      if (!(exact || starts || contains)) return null;
      let score = 0;
      if (exact) score += 1200;
      else if (starts) score += 500;
      else score += 220;
      const rr = node.getBoundingClientRect();
      const cx = rr.left + rr.width / 2;
      const cy = rr.top + rr.height / 2;
      const dx = cx - pcx;
      const dy = cy - pcy;
      score -= Math.sqrt(dx * dx + dy * dy) * 0.25;
      return { node, text: t, score };
    })
    .filter(Boolean)
    .sort((a, b) => b.score - a.score);

  if (matched.length) {
    try {
      const m = matched[0];
      const clickable = m.node.closest("a,button,[role='option'],[role='menuitem'],li,div,span") || m.node;
      clickable.click();
      out.clicked_option = true;
      out.text = m.text;
    } catch (_) {}
  }
  if (!matched.length) {
    out.error = "target_option_not_found";
    return out;
  }

  try {
    const keyOpts = { key: "Enter", code: "Enter", keyCode: 13, which: 13, bubbles: true, cancelable: true };
    pickedInput.dispatchEvent(new KeyboardEvent("keydown", keyOpts));
    pickedInput.dispatchEvent(new KeyboardEvent("keypress", keyOpts));
    pickedInput.dispatchEvent(new KeyboardEvent("keyup", keyOpts));
  } catch (_) {}

  out.selected = true;
  if (!out.text) out.text = value;
  return out;
}
"""
    raw = await page.evaluate(js, {"target": target, "value": value})
    parsed = _json_loads_if_possible(raw)
    if not isinstance(parsed, dict):
        return {"selected": False, "typed": False, "clicked_option": False, "text": None, "error": f"invalid_js_result:{raw}"}
    return parsed


async def _open_nearby_select_by_label(
    browser_session: BrowserSession,
    target_label: str,
    *,
    avoid_global_nav: bool = True,
    block_text_tokens: list[str] | None = None,
) -> dict:
    page = await browser_session.must_get_current_page()
    js = r"""
(input) => {
  const label = (input?.target_label || "").trim();
  const avoidGlobalNav = input?.avoid_global_nav !== false;
  const blockTextTokens = Array.isArray(input?.block_text_tokens) ? input.block_text_tokens : [];
  const out = { clicked: false, selector: null, text: null, error: null };
  if (!label) {
    out.error = "empty_target_label";
    return out;
  }

  const clean = (v) => String(v ?? "").replace(/\s+/g, " ").trim();
  const norm = (v) =>
    clean(v)
      .toLowerCase()
      .replace(/[\.。]/g, "")
      .replace(/[\s_\-—–·:：\/\\|（）()\[\]【】]+/g, "");

  const isVisible = (el) => {
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  const isLikelyGlobalNav = (node) => {
    if (!node) return false;
    if (node.closest("header,aside,nav,[role='banner']")) return true;
    let cur = node;
    for (let depth = 0; depth < 5 && cur; depth++) {
      const cls = clean(cur.className || "").toLowerCase();
      if (cls && /(header|topbar|navbar|toolbar|message|notice|user|avatar)/.test(cls)) {
        const r = cur.getBoundingClientRect();
        if (r && r.top < 260) return true;
      }
      cur = cur.parentElement;
    }
    return false;
  };

  const blockedNorm = blockTextTokens.map((v) => norm(v)).filter(Boolean);
  const labelNorm = norm(label);
  const last = window.__opt_last_interaction || null;
  const canUseLast =
    !!last &&
    Date.now() - Number(last.ts || 0) <= 20000 &&
    String(last.url || "") === `${location.pathname}${location.search}` &&
    Number.isFinite(Number(last.x)) &&
    Number.isFinite(Number(last.y));
  const distanceToLast = (node) => {
    try {
      const r = node.getBoundingClientRect();
      const cx = r.left + r.width / 2;
      const cy = r.top + r.height / 2;
      const dx = cx - Number(last.x || 0);
      const dy = cy - Number(last.y || 0);
      return Math.sqrt(dx * dx + dy * dy);
    } catch (_) {
      return 1e9;
    }
  };

  const scoreLabelNode = (node) => {
    if (!node || !isVisible(node)) return null;
    const rawText = clean(node.innerText || node.textContent || "");
    const t = norm(rawText);
    if (!t) return null;
    if (blockedNorm.some((tok) => t.includes(tok))) return null;
    if (rawText.length > Math.max(56, label.length * 4 + 16)) return null;
    const tag = String(node.tagName || "").toLowerCase();
    if ((tag === "div" || tag === "li") && rawText.length > Math.max(26, label.length * 3 + 8)) {
      const cls = clean(node.className || "").toLowerCase();
      if (!/(label|title|name|field|caption|item|selector|select|form|text)/.test(cls)) return null;
    }
    const tNoColon = t.replace(/[:：]/g, "");
    const targetNoColon = labelNorm.replace(/[:：]/g, "");
    const exact = tNoColon === targetNoColon;
    const starts = tNoColon.startsWith(targetNoColon);
    const contains = tNoColon.includes(targetNoColon);
    if (targetNoColon.length <= 4 && !exact) return null;
    if (!(exact || starts || contains || targetNoColon.includes(tNoColon))) return null;
    let score = 0;
    if (exact) score += 12000;
    else if (starts) score += 3600;
    else if (contains) score += 900;
    else score += 200;
    const extra = Math.max(0, tNoColon.length - targetNoColon.length);
    score -= extra * 45;
    score -= rawText.length * 3;
    return { node, score };
  };

  const collectLabelNodes = (selectors) =>
    Array.from(document.querySelectorAll(selectors))
      .map((node) => scoreLabelNode(node))
      .filter(Boolean)
      .sort((a, b) => b.score - a.score)
      .slice(0, 24)
      .map((x) => x.node);

  let labelNodes = collectLabelNodes("label,span,p,th,td");
  if (!labelNodes.length) labelNodes = collectLabelNodes("div,li");
  if (labelNodes.length && canUseLast) {
    const ranked = labelNodes
      .map((node) => ({ node, d: distanceToLast(node) }))
      .sort((a, b) => a.d - b.d);
    const near = ranked.filter((x) => x.d <= 680).map((x) => x.node);
    labelNodes = (near.length ? near : ranked.map((x) => x.node)).slice(0, 28);
  }

  const candidates = [];
  const isControlNode = (node) => {
    if (!node) return false;
    const cls = clean(node.className || "").toLowerCase();
    if (node.getAttribute && node.getAttribute("role") === "combobox") return true;
    const ariaPopup = (node.getAttribute && clean(node.getAttribute("aria-haspopup") || "").toLowerCase()) || "";
    const ariaExpanded = (node.getAttribute && clean(node.getAttribute("aria-expanded") || "").toLowerCase()) || "";
    if (ariaPopup && /listbox|menu|tree|dialog/.test(ariaPopup)) return true;
    if (ariaExpanded === "true" || ariaExpanded === "false") return true;
    const hasArrowIcon = !!node.querySelector(
      "[class*='arrow'],[class*='caret'],[class*='down'],.anticon-down,.icon-down,svg"
    );
    if (hasArrowIcon && /trigger|value|selection|select|cascader|dropdown|picker/.test(cls)) return true;
    if (/ant-select-selector|ant-cascader|ant-picker|combobox|selector/.test(cls)) return true;
    return false;
  };
  const pushCandidate = (node, source, labelRect = null, labelText = "") => {
    if (!node || !isVisible(node)) return;
    if (!isControlNode(node)) return;
    if (avoidGlobalNav) {
      const inGlobalNav = isLikelyGlobalNav(node);
      if (inGlobalNav) return;
    }
    const rect = node.getBoundingClientRect();
    if (!rect || rect.width <= 0 || rect.height <= 0) return;
    const text = clean(node.innerText || node.textContent || "");
    candidates.push({ node, source, text, rect, labelRect, labelText: clean(labelText || "") });
  };

  const controlSelectors = ".ant-select-selector,[role='combobox'],.ant-cascader-picker,.ant-cascader-picker-label,.ant-picker,[class*='cascader'],[class*='selector'],[class*='combobox']";

  for (const ln of labelNodes) {
    const lnRect = ln.getBoundingClientRect();
    const container =
      ln.closest(".ant-form-item,.ant-row,.ant-col,.ant-space-item,.ant-card,section,article,div,li,td") ||
      ln.parentElement ||
      document.body;
    if (!container) continue;
    const localNodes = Array.from(container.querySelectorAll(controlSelectors));
    const lnText = clean(ln.innerText || ln.textContent || "");
    for (const n of localNodes) pushCandidate(n, "label_container_control", lnRect, lnText);

    // Broader but still safe fallback: ancestor/sibling scopes with control-only filter.
    let cur = container;
    for (let depth = 0; depth < 4 && cur; depth++) {
      const parent = cur.parentElement;
      if (!parent) break;
      const inParent = Array.from(parent.querySelectorAll(controlSelectors));
      for (const n of inParent) pushCandidate(n, "label_ancestor_control", lnRect, lnText);
      for (const sib of Array.from(parent.children)) {
        if (sib === cur) continue;
        const sibNodes = Array.from(sib.querySelectorAll(controlSelectors));
        for (const n of sibNodes) pushCandidate(n, "label_sibling_control", lnRect, lnText);
      }
      cur = parent;
    }
  }

  // Last fallback: sibling/nearby nodes with dropdown trigger traits.
  if (!candidates.length) {
    for (const ln of labelNodes) {
      const lnRect = ln.getBoundingClientRect();
      const scopes = [];
      if (ln.parentElement) scopes.push(ln.parentElement);
      if (ln.closest(".ant-form-item,.ant-row,.ant-col,.ant-space-item,.ant-card,section,article,div,li,td")) {
        scopes.push(ln.closest(".ant-form-item,.ant-row,.ant-col,.ant-space-item,.ant-card,section,article,div,li,td"));
      }
      const lnText = clean(ln.innerText || ln.textContent || "");
      for (const scope of scopes.filter(Boolean)) {
        const nodes = Array.from(
          scope.querySelectorAll(
            "[role='button'],button,div,span,a,[class*='trigger'],[class*='value'],[class*='selection'],[class*='dropdown'],[class*='cascader']"
          )
        );
        for (const n of nodes) {
          const text = clean(n.innerText || n.textContent || "");
          if (/如何|帮助|了解更多|标签分析/.test(text)) continue;
          const rect = n.getBoundingClientRect();
          if (!rect || rect.width <= 0 || rect.height <= 0) continue;
          const verticalClose = Math.abs((rect.top + rect.height / 2) - (lnRect.top + lnRect.height / 2)) <= 64;
          const toRight = rect.left >= lnRect.right - 40;
          if (!(verticalClose && toRight)) continue;
          pushCandidate(n, "label_trigger_fallback", lnRect, lnText);
        }
      }
    }
  }

  if (!candidates.length) {
    out.error = "label_nearby_control_not_found";
    return out;
  }

  const scoreOf = (item) => {
    const cls = clean(item.node.className || "").toLowerCase();
    let score = 0;
    if (/ant-select-selector|combobox|ant-cascader|ant-picker/.test(cls)) score += 2600;
    else if (/role=combobox/.test(item.source || "")) score += 2000;
    else score += 400;
    if (item.source === "label_container_control") score += 700;
    else if (item.source === "label_ancestor_control") score += 450;
    else if (item.source === "label_sibling_control") score += 220;
    else if (item.source === "label_trigger_fallback") score += 180;
    const labelTextNorm = norm(item.labelText || "");
    if (labelTextNorm) {
      if (labelTextNorm === labelNorm) score += 780;
      else if (labelTextNorm.includes(labelNorm) || labelNorm.includes(labelTextNorm)) score += 300;
      else score -= 900;
    }
    if (item.text && item.text.length <= 20) score += 120;
    if (item.labelRect) {
      const lcx = item.labelRect.left + item.labelRect.width / 2;
      const lcy = item.labelRect.top + item.labelRect.height / 2;
      const cx2 = item.rect.left + item.rect.width / 2;
      const cy2 = item.rect.top + item.rect.height / 2;
      const ddx = cx2 - lcx;
      const ddy = cy2 - lcy;
      score -= Math.sqrt(ddx * ddx + ddy * ddy) * 0.9;
      const rightDx = cx2 - item.labelRect.right;
      const absDy = Math.abs(cy2 - lcy);
      if (rightDx >= -20 && rightDx <= 520) score += 260;
      else score -= Math.abs(rightDx) * 0.45;
      score -= absDy * 1.1;
    }
    const cx = item.rect.left + item.rect.width / 2;
    const cy = item.rect.top + item.rect.height / 2;
    const dx = cx - (window.innerWidth / 2);
    const dy = cy - (window.innerHeight / 2);
    score -= Math.sqrt(dx * dx + dy * dy) * 0.08;
    return score;
  };

  const best = candidates
    .map((c) => ({ ...c, score: scoreOf(c) }))
    .sort((a, b) => b.score - a.score)[0];

  try {
    best.node.scrollIntoView({ block: "center", inline: "center" });
    const clickable =
      best.node.closest("[role='combobox'],.ant-select-selector,.ant-cascader-picker,.ant-cascader-picker-label,.ant-picker") ||
      best.node;
    clickable.click();
    const r = clickable.getBoundingClientRect();
    window.__opt_last_interaction = {
      kind: "open_by_label",
      text: clean(label),
      x: r.left + r.width / 2,
      y: r.top + r.height / 2,
      ts: Date.now(),
      url: `${location.pathname}${location.search}`,
    };
    out.clicked = true;
    out.selector = best.source || null;
    out.text = clean(clickable.innerText || clickable.textContent || "") || clean(label);
    out.matched_label = clean(best.labelText || "") || null;
    return out;
  } catch (e) {
    out.error = String(e);
    return out;
  }
}
"""
    raw = await page.evaluate(
        js,
        {
            "target_label": target_label,
            "avoid_global_nav": avoid_global_nav,
            "block_text_tokens": block_text_tokens or [],
        },
    )
    parsed = _json_loads_if_possible(raw)
    if not isinstance(parsed, dict):
        return {"clicked": False, "selector": None, "text": None, "error": f"invalid_js_result:{raw}"}
    return parsed


async def _is_option_overlay_visible(browser_session: BrowserSession) -> dict:
    page = await browser_session.must_get_current_page()
    js = r"""
() => {
  const out = { visible: false, count: 0 };
  const isVisible = (el) => {
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  const nodes = Array.from(
    document.querySelectorAll(
      ".ant-select-dropdown,.ant-dropdown,.ant-popover,.ant-cascader-menus,[role='listbox'],[role='menu'],[class*='dropdown'],[class*='popover']"
    )
  ).filter(isVisible);
  out.count = nodes.length;
  out.visible = nodes.length > 0;
  return out;
}
"""
    raw = await page.evaluate(js)
    parsed = _json_loads_if_possible(raw)
    if not isinstance(parsed, dict):
        return {"visible": False, "count": 0}
    return {"visible": bool(parsed.get("visible")), "count": int(parsed.get("count", 0) or 0)}


def _normalize_date_text(value: str | None) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    text = text.replace("至", "~").replace("—", "~").replace("-", "~")
    text = text.replace("/", ".")
    text = re.sub(r"\s+", "", text)
    return text


def _parse_date_range_value(value: str | None) -> tuple[str, str] | None:
    text = str(value or "").strip()
    if not text:
        return None
    match = re.search(
        r"(\d{4}[./-]\d{1,2}[./-]\d{1,2})\s*(?:至|~|—|-|to)\s*(\d{4}[./-]\d{1,2}[./-]\d{1,2})",
        text,
        flags=re.IGNORECASE,
    )
    if not match:
        return None
    return match.group(1).strip(), match.group(2).strip()


async def _verify_date_filter_applied(browser_session: BrowserSession, target: str | None, value: str) -> dict:
    page = await browser_session.must_get_current_page()
    expected = _normalize_date_text(value)
    if not expected:
        return {"applied": False, "reason": "empty_expected", "evidence": None}
    js = r"""
(input) => {
  const target = String(input?.target || "").trim();
  const expected = String(input?.expected || "").trim();
  const out = { applied: false, reason: null, evidence: null };
  if (!expected) {
    out.reason = "empty_expected";
    return out;
  }
  const clean = (v) => String(v ?? "").replace(/\s+/g, "").trim();
  const norm = (v) => clean(v).replace(/至/g, "~").replace(/[—-]/g, "~").replace(/\//g, ".");
  const isVisible = (el) => {
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  const textOf = (el) =>
    clean(
      el?.value ||
      el?.getAttribute?.("value") ||
      el?.getAttribute?.("aria-label") ||
      el?.innerText ||
      el?.textContent ||
      ""
    );

  const targetNorm = norm(target);
  const expectedNorm = norm(expected);
  const parseRange = (txt) => {
    const t = norm(txt);
    if (!t || !t.includes("~")) return null;
    const parts = t.split("~");
    if (parts.length < 2) return null;
    const start = parts[0] || "";
    const end = parts[1] || "";
    if (!start || !end) return null;
    return { start, end };
  };
  const expectedRange = parseRange(expectedNorm);
  const matchValue = (txt) => {
    const t = norm(txt);
    if (!t) return false;
    if (t === expectedNorm || t.includes(expectedNorm) || expectedNorm.includes(t)) return true;
    if (expectedRange) {
      const got = parseRange(t);
      if (got && got.start && got.end) {
        return (
          (got.start === expectedRange.start || got.start.includes(expectedRange.start) || expectedRange.start.includes(got.start)) &&
          (got.end === expectedRange.end || got.end.includes(expectedRange.end) || expectedRange.end.includes(got.end))
        );
      }
      return t.includes(expectedRange.start) && t.includes(expectedRange.end);
    }
    return false;
  };

  const candidateSelectors = [
    ".ant-picker-input input",
    ".ant-picker input",
    "input[placeholder*='开始']",
    "input[placeholder*='结束']",
    "input[placeholder*='日期']",
    "input[placeholder*='时间']",
    "input[placeholder*='range']",
    "input",
    ".ant-select-selector",
    "[role='combobox']",
    ".ant-form-item",
    "label",
    "span",
    "div"
  ];

  const seen = new Set();
  const nodes = [];
  for (const sel of candidateSelectors) {
    const list = Array.from(document.querySelectorAll(sel));
    for (const node of list) {
      if (!node || seen.has(node)) continue;
      seen.add(node);
      if (!isVisible(node)) continue;
      const text = textOf(node);
      if (!text) continue;
      nodes.push({ node, text });
    }
  }

  const targetFiltered = [];
  for (const item of nodes) {
    const node = item.node;
    const txt = item.text;
    if (!targetNorm) {
      targetFiltered.push(item);
      continue;
    }
    const nearText = clean(
      node?.closest?.(".ant-form-item,.ant-row,.ant-col,section,article,div")?.innerText ||
      node?.parentElement?.innerText ||
      ""
    );
    const nearNorm = norm(nearText);
    if (nearNorm.includes(targetNorm) || targetNorm.includes(nearNorm)) {
      targetFiltered.push(item);
      continue;
    }
    const ownNorm = norm(txt);
    if (ownNorm.includes(targetNorm) || targetNorm.includes(ownNorm)) {
      targetFiltered.push(item);
      continue;
    }
  }

  const pool = targetFiltered.length ? targetFiltered : nodes;
  for (const item of pool) {
    const txt = item.text;
    if (matchValue(txt)) {
      out.applied = true;
      out.reason = targetFiltered.length ? "target_value_match" : "global_value_match";
      out.evidence = txt;
      return out;
    }
  }

  if (!pool.length) {
    out.reason = "no_visible_candidates";
  } else {
    out.reason = "value_mismatch";
    out.evidence = pool.slice(0, 3).map((x) => x.text).join(" | ");
  }
  return out;
}
"""
    raw = await page.evaluate(js, {"target": target or "", "expected": expected})
    parsed = _json_loads_if_possible(raw)
    if not isinstance(parsed, dict):
        return {"applied": False, "reason": f"invalid_js_result:{raw}", "evidence": None}
    return parsed


async def _set_date_filter(browser_session: BrowserSession, target: str | None, value: str) -> dict:
    result = {
        "status": "failed",
        "date_type": None,
        "date_range": None,
        "details": [],
    }
    if not value:
        result["status"] = "skipped"
        return result

    normalized = value.strip()
    open_selectors = ["label", ".ant-picker", ".ant-select-selector", "[role='combobox']", ".ant-form-item-label"]

    if target:
        click_target = await _click_text_step(
            browser_session,
            target,
            open_selectors,
            allow_reverse_contains=False,
            avoid_global_nav=False,
            require_clickable=True,
            max_text_len=48,
        )
        result["details"].append({"phase": "open_target", "ok": bool(click_target.get("clicked")), "raw": click_target})
        await asyncio.sleep(0.35)

    date_range = _parse_date_range_value(normalized)
    if date_range is None:
        selected = {"selected": False, "error": "not_attempted"}
        if target:
            targeted = await _set_select_value_by_target(browser_session, target, normalized)
            result["details"].append({"phase": "set_type_targeted", "ok": bool(targeted.get("selected")), "raw": targeted})
            if targeted.get("selected"):
                selected = targeted
        if not selected.get("selected"):
            selected = await _select_option_by_text(
                browser_session,
                normalized,
                target_label=target,
                avoid_global_nav=True,
            )
            result["details"].append({"phase": "select_type_targeted", "ok": bool(selected.get("selected")), "raw": selected})
        if not selected.get("selected"):
            selected = await _select_option_by_text(browser_session, normalized)
            result["details"].append({"phase": "select_type_global", "ok": bool(selected.get("selected")), "raw": selected})

        verify_type = await _verify_selection_applied(browser_session, target, normalized)
        result["details"].append({"phase": "verify_type", "ok": bool(verify_type.get("applied")), "raw": verify_type})
        if (not verify_type.get("applied")) and target:
            reopen = await _click_text_step(
                browser_session,
                target,
                open_selectors,
                allow_reverse_contains=False,
                avoid_global_nav=False,
                require_clickable=True,
                max_text_len=48,
            )
            result["details"].append({"phase": "reopen_target_for_type", "ok": bool(reopen.get("clicked")), "raw": reopen})
            await asyncio.sleep(0.35)
            retry = await _set_select_value_by_target(browser_session, target, normalized)
            result["details"].append({"phase": "retry_set_type_targeted", "ok": bool(retry.get("selected")), "raw": retry})
            verify_type = await _verify_selection_applied(browser_session, target, normalized)
            result["details"].append({"phase": "retry_verify_type", "ok": bool(verify_type.get("applied")), "raw": verify_type})

        if verify_type.get("applied"):
            result["status"] = "ok"
            result["date_type"] = normalized
        return result

    start, end = date_range
    result["date_range"] = f"{start}~{end}"
    verify_range = {"applied": False, "reason": "not_verified", "evidence": None}

    for attempt in range(2):
        if attempt == 1 and target:
            reopen = await _click_text_step(
                browser_session,
                target,
                open_selectors,
                allow_reverse_contains=False,
                avoid_global_nav=False,
                require_clickable=True,
                max_text_len=48,
            )
            result["details"].append({"phase": "reopen_target_for_range", "ok": bool(reopen.get("clicked")), "raw": reopen})
            await asyncio.sleep(0.35)

        fill_start = await _fill_input_by_target(browser_session, "开始", start)
        fill_end = await _fill_input_by_target(browser_session, "结束", end)
        result["details"].append({"phase": f"fill_start_{attempt+1}", "ok": bool(fill_start.get("filled")), "raw": fill_start})
        result["details"].append({"phase": f"fill_end_{attempt+1}", "ok": bool(fill_end.get("filled")), "raw": fill_end})
        await asyncio.sleep(0.25)
        verify_range = await _verify_date_filter_applied(browser_session, target, result["date_range"])
        result["details"].append({"phase": f"verify_range_{attempt+1}", "ok": bool(verify_range.get("applied")), "raw": verify_range})
        if verify_range.get("applied"):
            result["status"] = "ok"
            return result

    result["status"] = "failed"
    return result


async def _run_module_interaction_steps(
    browser_session: BrowserSession,
    module: ModuleSpec,
    profile_name: str,
    context: dict,
    task_dir: Path | None,
    max_wait_seconds: int | None,
) -> dict:
    await _ensure_single_tab_guard(browser_session)
    profile = get_profile_spec(profile_name)
    blocked_tokens = _profile_block_text_tokens(profile)
    fields: list[dict] = []
    files: list[dict] = []
    route_trace: list[dict] = []
    step_trace: list[dict] = []
    recent_context_hints: list[str] = []
    required_failures: list[dict] = []

    def push_hint(value: str | None) -> None:
        text = (value or "").strip()
        if not text:
            return
        recent_context_hints.append(text)
        # Keep order, dedup, and cap size
        deduped = list(dict.fromkeys([item for item in recent_context_hints if item.strip()]))
        recent_context_hints[:] = deduped[-10:]

    module_anchor_candidates: list[str] = []
    for token in [*module.route_path]:
        t = str(token or "").strip()
        if t:
            module_anchor_candidates.append(t)
    for step in module.interaction_steps:
        if step.op in {"click", "hover", "ensure_visible", "select", "select_path", "input", "set_date_range"}:
            for token in [step.target, step.value, *(step.path or [])]:
                t = str(token or "").strip()
                if not t or t.startswith(("params.", "module.", "context.")):
                    continue
                module_anchor_candidates.append(t)
    module_anchor_candidates = list(dict.fromkeys(module_anchor_candidates))

    for index, step in enumerate(module.interaction_steps, 1):
        if not _eval_when_expr(step.when, context):
            step_trace.append(
                {
                    "step_no": index,
                    "op": step.op,
                    "status": "skipped",
                    "reason": "when_false",
                }
            )
            continue

        retries = max(step.retries, 1)
        last_error = None
        step_ok = False
        step_payload: dict = {}

        for _attempt in range(retries):
            try:
                if step.op == "wait":
                    seconds = step.seconds if step.seconds is not None else 0.6
                    await asyncio.sleep(max(0.0, float(seconds)))
                    step_payload = {"waited_seconds": seconds}
                    step_ok = True
                elif step.op == "ensure_visible":
                    if not step.target:
                        step_ok = True
                        step_payload = {"visible": True, "note": "empty_target"}
                    else:
                        visible = await _is_text_visible(browser_session, step.target)
                        if not visible:
                            clicked = await _click_text_step(
                                browser_session,
                                step.target,
                                module.route_selectors,
                                allow_reverse_contains=False,
                                avoid_global_nav=False,
                                require_clickable=True,
                                max_text_len=64,
                                block_text_tokens=blocked_tokens,
                            )
                            await asyncio.sleep(0.25)
                            visible = await _is_text_visible(browser_session, step.target)
                            route_trace.append(
                                {
                                    "step_no": len(route_trace) + 1,
                                    "target": step.target,
                                    "clicked": bool(clicked.get("clicked")),
                                    "selector": clicked.get("selector"),
                                    "error": clicked.get("error"),
                                }
                            )
                            step_payload = {"visible": visible, "click_attempt": clicked}
                            if clicked.get("clicked"):
                                guard = await _guard_and_recover_unexpected_page(browser_session)
                                step_payload["unexpected_page_guard"] = guard
                                if not guard.get("ok"):
                                    visible = False
                        else:
                            step_payload = {"visible": True}
                        step_ok = bool(visible)
                        if step_ok:
                            push_hint(step.target)
                elif step.op == "click":
                    if not step.target:
                        raise ValueError("click_target_empty")
                    clicked = await _click_text_step(
                        browser_session,
                        step.target,
                        module.route_selectors,
                        allow_reverse_contains=False,
                        avoid_global_nav=False,
                        require_clickable=True,
                        max_text_len=64,
                        block_text_tokens=blocked_tokens,
                    )
                    route_trace.append(
                        {
                            "step_no": len(route_trace) + 1,
                            "target": step.target,
                            "clicked": bool(clicked.get("clicked")),
                            "selector": clicked.get("selector"),
                            "error": clicked.get("error"),
                        }
                    )
                    step_ok = bool(clicked.get("clicked"))
                    step_payload = clicked
                    if step_ok:
                        guard = await _guard_and_recover_unexpected_page(browser_session)
                        step_payload = {**step_payload, "unexpected_page_guard": guard}
                        if not guard.get("ok"):
                            step_ok = False
                    if step_ok:
                        push_hint(step.target)
                        push_hint(clicked.get("text"))
                        await asyncio.sleep(0.25)
                elif step.op == "hover":
                    hover_field_keys = [str(k).strip() for k in (step.field_keys or []) if str(k).strip()]
                    if hover_field_keys:
                        specs = _resolve_field_specs(profile_name, hover_field_keys)
                        if not specs:
                            if step.required:
                                raise ValueError("hover_field_specs_empty")
                            step_ok = True
                            step_payload = {"status": "skipped", "reason": "empty_field_specs"}
                        else:
                            hover_targets: list[str] = []
                            if step.target:
                                hover_targets.append(step.target)
                            for item in specs:
                                for token in [
                                    item.field_key,
                                    item.label,
                                    *(item.alt_labels or []),
                                    _benchmark_anchor_from_spec(item),
                                ]:
                                    text = str(token or "").strip()
                                    if not text:
                                        continue
                                    if text in hover_targets:
                                        continue
                                    hover_targets.append(text)
                            hover_targets = hover_targets[:10]

                            hover_attempts: list[dict] = []
                            hovered_success: dict | None = None
                            for target in hover_targets:
                                hovered = await _hover_text_step(
                                    browser_session=browser_session,
                                    label=target,
                                    semantic_text=target,
                                    selectors=list(DEFAULT_CONTENT_HOVER_SELECTORS),
                                    allow_reverse_contains=False,
                                    avoid_global_nav=True,
                                    max_text_len=72,
                                    block_text_tokens=blocked_tokens,
                                    strict_selectors_only=True,
                                    wait_seconds=0.65,
                                )
                                if not hovered.get("hovered"):
                                    hovered = await _hover_text_step(
                                        browser_session=browser_session,
                                        label=target,
                                        semantic_text=target,
                                        selectors=list(DEFAULT_CONTENT_HOVER_SELECTORS),
                                        allow_reverse_contains=True,
                                        avoid_global_nav=True,
                                        max_text_len=96,
                                        block_text_tokens=blocked_tokens,
                                        strict_selectors_only=False,
                                        wait_seconds=0.9,
                                    )
                                hover_attempts.append(
                                    {
                                        "target": target,
                                        "hovered": bool(hovered.get("hovered")),
                                        "selector": hovered.get("selector"),
                                        "text": hovered.get("text"),
                                        "error": hovered.get("error"),
                                    }
                                )
                                route_trace.append(
                                    {
                                        "step_no": len(route_trace) + 1,
                                        "target": target,
                                        "hovered": bool(hovered.get("hovered")),
                                        "selector": hovered.get("selector"),
                                        "error": hovered.get("error"),
                                    }
                                )
                                if hovered.get("hovered"):
                                    hovered_success = hovered
                                    push_hint(target)
                                    push_hint(hovered.get("text"))
                                    await asyncio.sleep(0.2)
                                    break

                            extracted = await _extract_fields_by_specs(browser_session, specs)
                            retry_specs = [
                                spec
                                for spec, item in zip(specs, extracted, strict=False)
                                if item.get("status") in {"not_found", "parse_failed", "hover_failed", "error"}
                            ]
                            merged = extracted
                            recovered_count = 0
                            if retry_specs:
                                await asyncio.sleep(0.25)
                                retried = await _extract_fields_by_specs(browser_session, retry_specs)
                                retry_map = {item.get("field_key"): item for item in retried if item.get("field_key")}
                                new_merged: list[dict] = []
                                for item in extracted:
                                    key = item.get("field_key")
                                    retry_item = retry_map.get(key)
                                    if not retry_item:
                                        new_merged.append(item)
                                        continue
                                    item_ok = item.get("status") == "ok" and item.get("value") is not None
                                    retry_ok = retry_item.get("status") == "ok" and retry_item.get("value") is not None
                                    if (not item_ok and retry_ok) or (
                                        not item_ok and item.get("source") == "none" and retry_item.get("source") != "none"
                                    ):
                                        new_merged.append(retry_item)
                                        recovered_count += 1
                                    else:
                                        new_merged.append(item)
                                merged = new_merged
                            fields.extend(merged)
                            ok_count = sum(1 for item in merged if item.get("status") == "ok")
                            step_ok = ok_count > 0 or not step.required
                            step_payload = {
                                "field_keys": hover_field_keys,
                                "hover_targets": hover_targets,
                                "hover_attempts": hover_attempts,
                                "hovered": bool(hovered_success and hovered_success.get("hovered")),
                                "count": len(merged),
                                "ok_count": ok_count,
                                "retried": len(retry_specs),
                                "recovered": recovered_count,
                            }
                            if ok_count > 0:
                                for item in merged:
                                    if item.get("status") == "ok":
                                        push_hint(item.get("field_key"))
                    elif not step.target:
                        if step.required:
                            raise ValueError("hover_target_empty")
                        step_ok = True
                        step_payload = {"status": "skipped", "reason": "empty_target"}
                    else:
                        hovered = await _hover_text_step(
                            browser_session=browser_session,
                            label=step.target,
                            semantic_text=step.target,
                            selectors=list(DEFAULT_CONTENT_HOVER_SELECTORS),
                            allow_reverse_contains=False,
                            avoid_global_nav=True,
                            max_text_len=72,
                            block_text_tokens=blocked_tokens,
                            strict_selectors_only=True,
                            wait_seconds=0.65,
                        )
                        if not hovered.get("hovered"):
                            hovered = await _hover_text_step(
                                browser_session=browser_session,
                                label=step.target,
                                semantic_text=step.target,
                                selectors=list(DEFAULT_CONTENT_HOVER_SELECTORS),
                                allow_reverse_contains=True,
                                avoid_global_nav=True,
                                max_text_len=96,
                                block_text_tokens=blocked_tokens,
                                strict_selectors_only=False,
                                wait_seconds=0.85,
                            )
                        route_trace.append(
                            {
                                "step_no": len(route_trace) + 1,
                                "target": step.target,
                                "hovered": bool(hovered.get("hovered")),
                                "selector": hovered.get("selector"),
                                "error": hovered.get("error"),
                            }
                        )
                        step_ok = bool(hovered.get("hovered"))
                        step_payload = hovered
                        if step_ok:
                            push_hint(step.target)
                            push_hint(hovered.get("text"))
                            await asyncio.sleep(0.25)
                elif step.op == "select":
                    target_click = None
                    if step.target:
                        target_click = await _click_text_step(
                            browser_session,
                            step.target,
                            module.route_selectors,
                            allow_reverse_contains=False,
                            avoid_global_nav=False,
                            require_clickable=True,
                            max_text_len=64,
                            block_text_tokens=blocked_tokens,
                        )
                        if not target_click.get("clicked"):
                            target_click = await _open_nearby_select_by_label(
                                browser_session=browser_session,
                                target_label=step.target,
                                avoid_global_nav=True,
                                block_text_tokens=blocked_tokens,
                            )
                        await asyncio.sleep(0.2)
                    select_values = _resolve_step_select_values(step, context)
                    if not select_values:
                        if step.required:
                            raise ValueError("select_value_empty")
                        step_ok = True
                        step_payload = {"status": "skipped", "reason": "empty_value"}
                    elif len(select_values) == 1:
                        value = select_values[0]
                        selected = await _select_option_by_text(
                            browser_session,
                            value,
                            target_label=step.target,
                            avoid_global_nav=False,
                            block_text_tokens=blocked_tokens,
                        )
                        verify = {"applied": False, "reason": "not_selected", "evidence": None}
                        guard = None
                        if selected.get("selected"):
                            await asyncio.sleep(0.15)
                            verify = await _verify_selection_applied(browser_session, step.target, value)
                        step_ok = bool(selected.get("selected")) and (bool(verify.get("applied")) or not step.required)
                        if step_ok:
                            guard = await _guard_and_recover_unexpected_page(browser_session)
                            if not guard.get("ok"):
                                step_ok = False
                        step_payload = {
                            "target_click": target_click,
                            "selected": selected,
                            "value": value,
                            "verify": verify,
                            "unexpected_page_guard": guard,
                        }
                        if step_ok:
                            push_hint(step.target)
                            push_hint(value)
                            push_hint(selected.get("text"))
                            await asyncio.sleep(0.25)
                    else:
                        chain = []
                        for item in select_values:
                            sel = await _select_option_by_text(
                                browser_session,
                                item,
                                target_label=step.target,
                                avoid_global_nav=False,
                                block_text_tokens=blocked_tokens,
                            )
                            chain.append({"value": item, "selected": bool(sel.get("selected")), "raw": sel})
                            await asyncio.sleep(0.2)
                        verify = {"applied": False, "reason": "empty_path", "evidence": None}
                        if chain:
                            verify = await _verify_selection_applied(browser_session, step.target, select_values[-1])
                        step_ok = all(item["selected"] for item in chain) and (
                            bool(verify.get("applied")) or not step.required
                        )
                        guard = None
                        if step_ok:
                            guard = await _guard_and_recover_unexpected_page(browser_session)
                            if not guard.get("ok"):
                                step_ok = False
                        step_payload = {
                            "target_click": target_click,
                            "path": chain,
                            "value": " / ".join(select_values),
                            "verify": verify,
                            "unexpected_page_guard": guard,
                        }
                        if step_ok:
                            push_hint(step.target)
                            for item in select_values:
                                push_hint(item)
                            for item in chain:
                                raw_text = ((item.get("raw") or {}).get("text") if isinstance(item.get("raw"), dict) else None)
                                push_hint(raw_text)
                            await asyncio.sleep(0.25)
                elif step.op == "select_path":
                    path_values = _resolve_step_path(step, context)
                    if not path_values:
                        path_values = _resolve_step_select_values(step, context)
                    target_click = None
                    if step.target:
                        target_click = await _click_text_step(
                            browser_session,
                            step.target,
                            module.route_selectors,
                            allow_reverse_contains=False,
                            avoid_global_nav=False,
                            require_clickable=True,
                            max_text_len=64,
                            block_text_tokens=blocked_tokens,
                        )
                        await asyncio.sleep(0.2)
                    chain = []
                    if not path_values:
                        if step.required:
                            raise ValueError("select_path_empty")
                        step_ok = True
                        step_payload = {"status": "skipped", "reason": "empty_path"}
                    else:
                        for item in path_values:
                            sel = await _select_option_by_text(
                                browser_session,
                                item,
                                target_label=step.target,
                                avoid_global_nav=False,
                                block_text_tokens=blocked_tokens,
                            )
                            chain.append({"value": item, "selected": bool(sel.get("selected")), "raw": sel})
                            await asyncio.sleep(0.2)
                        verify = {"applied": False, "reason": "empty_path", "evidence": None}
                        if chain:
                            verify = await _verify_selection_applied(browser_session, step.target, path_values[-1])
                        step_ok = all(item["selected"] for item in chain) and (
                            bool(verify.get("applied")) or not step.required
                        )
                        guard = None
                        if step_ok:
                            guard = await _guard_and_recover_unexpected_page(browser_session)
                            if not guard.get("ok"):
                                step_ok = False
                        step_payload = {"target_click": target_click, "path": chain, "verify": verify, "unexpected_page_guard": guard}
                        if step_ok:
                            push_hint(step.target)
                            for item in path_values:
                                push_hint(item)
                            for item in chain:
                                raw_text = ((item.get("raw") or {}).get("text") if isinstance(item.get("raw"), dict) else None)
                                push_hint(raw_text)
                elif step.op == "input":
                    value = _resolve_step_value(step, context)
                    if not value:
                        if step.required:
                            raise ValueError("input_value_empty")
                        step_ok = True
                        step_payload = {"status": "skipped", "reason": "empty_value"}
                    else:
                        target = step.target or "输入"
                        filled = await _fill_input_by_target(browser_session, target, value)
                        step_ok = bool(filled.get("filled"))
                        step_payload = {"filled": filled, "value": value}
                        if step_ok:
                            push_hint(target)
                            push_hint(value)
                elif step.op == "set_date_range":
                    value = _resolve_step_value(step, context)
                    if not value:
                        if step.required:
                            raise ValueError("date_value_empty")
                        step_ok = True
                        step_payload = {"status": "skipped", "reason": "empty_value"}
                    else:
                        set_result = await _set_date_filter(browser_session, step.target, value)
                        step_ok = set_result.get("status") == "ok"
                        guard = None
                        if step_ok:
                            guard = await _guard_and_recover_unexpected_page(browser_session)
                            if not guard.get("ok"):
                                step_ok = False
                        step_payload = {**set_result, "unexpected_page_guard": guard}
                        if step_ok:
                            push_hint(step.target)
                            push_hint(value)
                elif step.op == "extract_fields":
                    field_keys = step.field_keys or module.field_keys
                    if not field_keys:
                        if step.required:
                            raise ValueError("field_keys_empty")
                        step_ok = True
                        step_payload = {"status": "skipped", "reason": "empty_field_keys"}
                    else:
                        specs = _resolve_field_specs(profile_name, field_keys)
                        extracted = await _extract_fields_by_specs(browser_session, specs)
                        retry_specs = [
                            spec
                            for spec, item in zip(specs, extracted, strict=False)
                            if item.get("status") in {"not_found", "parse_failed", "hover_failed", "error"}
                        ]
                        merged = extracted
                        recovered_count = 0
                        if retry_specs:
                            await asyncio.sleep(0.25)
                            retried = await _extract_fields_by_specs(browser_session, retry_specs)
                            retry_map = {item.get("field_key"): item for item in retried if item.get("field_key")}
                            new_merged: list[dict] = []
                            for item in extracted:
                                key = item.get("field_key")
                                retry_item = retry_map.get(key)
                                if not retry_item:
                                    new_merged.append(item)
                                    continue
                                item_ok = item.get("status") == "ok" and item.get("value") is not None
                                retry_ok = retry_item.get("status") == "ok" and retry_item.get("value") is not None
                                if (not item_ok and retry_ok) or (not item_ok and item.get("source") == "none" and retry_item.get("source") != "none"):
                                    new_merged.append(retry_item)
                                    recovered_count += 1
                                else:
                                    new_merged.append(item)
                            merged = new_merged
                        fields.extend(merged)
                        step_ok = any(item.get("status") == "ok" for item in merged) or not step.required
                        step_payload = {
                            "field_keys": field_keys,
                            "count": len(merged),
                            "retried": len(retry_specs),
                            "recovered": recovered_count,
                        }
                elif step.op == "download":
                    resolved_file_keys = []
                    if step.file_key:
                        resolved_file_keys.append(step.file_key)
                    if step.file_keys:
                        resolved_file_keys.extend(step.file_keys)
                    if not resolved_file_keys:
                        resolved_file_keys = module.file_keys
                    if not resolved_file_keys:
                        if step.required:
                            raise ValueError("file_keys_empty")
                        step_ok = True
                        step_payload = {"status": "skipped", "reason": "empty_file_keys"}
                    else:
                        specs = _resolve_file_specs(profile_name, resolved_file_keys)
                        download_dir = task_dir if task_dir else Path(DOWNLOADS_PATH)
                        download_dir.mkdir(parents=True, exist_ok=True)
                        generic_hint_tokens = {
                            "全部",
                            "触点",
                            "时间",
                            "日期",
                            "筛选",
                            "全部内容",
                            "下载",
                            "导出",
                            "查看",
                            "详情",
                        }
                        generic_hint_tokens = {g.strip() for g in generic_hint_tokens}
                        broad_hints = [
                            hint
                            for hint in recent_context_hints[-6:]
                            if hint and hint.strip() and hint.strip() not in generic_hint_tokens
                        ]
                        focused_hints = [
                            hint
                            for hint in recent_context_hints[-3:]
                            if hint
                            and hint.strip()
                            and hint.strip() not in generic_hint_tokens
                            and not re.fullmatch(r"-?\d+(?:[\.,]\d+)?%?", hint.strip())
                        ]
                        strong_hints = focused_hints or broad_hints
                        terminal_anchor = next(
                            (
                                hint
                                for hint in reversed(module_anchor_candidates)
                                if hint
                                and hint.strip()
                                and hint.strip() not in generic_hint_tokens
                                and not re.fullmatch(r"-?\d+(?:[\.,]\d+)?%?", hint.strip())
                            ),
                            None,
                        )
                        anchor_hints = [
                            hint
                            for hint in [terminal_anchor]
                            if hint
                            and hint.strip()
                            and hint.strip() not in generic_hint_tokens
                            and len(hint.strip()) >= 2
                            and not re.fullmatch(r"-?\d+(?:[\.,]\d+)?%?", hint.strip())
                        ]
                        context_hints = list(
                            dict.fromkeys(
                                [
                                    module.label,
                                    *module.route_path,
                                    *module_anchor_candidates,
                                    *strong_hints,
                                ]
                            )
                        )
                        min_context_hits = 3 if len(context_hints) >= 5 else (2 if len(context_hints) >= 3 else 1)
                        min_hard_context_hits = 1 if focused_hints else (2 if len(strong_hints) >= 3 else (1 if strong_hints else 0))
                        if terminal_anchor:
                            min_hard_context_hits = max(min_hard_context_hits, 1)
                        hard_context_hints = [*(strong_hints or [])]
                        if terminal_anchor:
                            hard_context_hints.append(terminal_anchor)
                        hard_context_hints = list(dict.fromkeys([h for h in hard_context_hints if str(h or "").strip()]))

                        downloaded = await _download_files_by_specs(
                            browser_session=browser_session,
                            specs=specs,
                            download_dir=download_dir,
                            max_wait_seconds=max_wait_seconds,
                            download_context={
                                "context_hints": context_hints,
                                "context_exclude_hints": [],
                                "hard_context_hints": hard_context_hints,
                                "anchor_hints": anchor_hints,
                                "min_context_hits": min_context_hits,
                                "min_hard_context_hits": min_hard_context_hits,
                                "min_anchor_hits": 1 if anchor_hints else 0,
                                "anchor_max_distance": 240 if anchor_hints else 0,
                                "last_interaction_max_distance": 420 if focused_hints else (620 if strong_hints else 760),
                            },
                        )
                        files.extend(downloaded)
                        step_ok = any(item.get("status") == "ok" for item in downloaded) or not step.required
                        step_payload = {
                            "file_keys": resolved_file_keys,
                            "count": len(downloaded),
                            "context_hints": context_hints,
                            "hard_context_hints": strong_hints,
                            "anchor_hints": anchor_hints,
                        }
                else:
                    raise ValueError(f"unsupported_op:{step.op}")

                if step_ok:
                    break
                last_error = f"step_failed:{step.op}"
            except Exception as exc:  # noqa: BLE001
                last_error = str(exc)
                await asyncio.sleep(0.2)

        step_trace.append(
            {
                "step_no": index,
                "op": step.op,
                "target": step.target,
                "success": bool(step_ok),
                "required": step.required,
                "error": None if step_ok else last_error,
                "payload": step_payload,
            }
        )
        if step.required and not step_ok:
            required_failures.append(
                {
                    "step_no": index,
                    "op": step.op,
                    "target": step.target,
                    "error": last_error,
                }
            )
            # Fail-fast for current module when a required step fails.
            # This avoids cascading drifts (e.g. accidental jump to unrelated pages).
            break

    return {
        "fields": fields,
        "files": files,
        "route_trace": route_trace,
        "step_trace": step_trace,
        "required_failures": required_failures,
    }


async def _navigate_route_path_for_module(
    *,
    browser_session: BrowserSession,
    module: ModuleSpec,
    block_text_tokens: list[str] | None = None,
) -> tuple[list[dict], list[dict]]:
    route_trace: list[dict] = []
    required_failures: list[dict] = []

    for index, step_label in enumerate(module.route_path, 1):
        step_result = await _click_text_step(
            browser_session=browser_session,
            label=step_label,
            selectors=module.route_selectors,
            allow_reverse_contains=False,
            avoid_global_nav=True,
            require_clickable=True,
            max_text_len=64,
            block_text_tokens=block_text_tokens or [],
        )
        if not step_result.get("clicked"):
            step_result = await _click_nearby_by_label(
                browser_session=browser_session,
                target_label=step_label,
                avoid_global_nav=True,
                block_text_tokens=block_text_tokens or [],
            )
        if not step_result.get("clicked"):
            await _scroll_main_content_to_top(browser_session)
            await asyncio.sleep(0.15)
            step_result = await _click_text_step(
                browser_session=browser_session,
                label=step_label,
                selectors=module.route_selectors,
                allow_reverse_contains=False,
                avoid_global_nav=True,
                require_clickable=True,
                max_text_len=64,
                block_text_tokens=block_text_tokens or [],
            )
            if not step_result.get("clicked"):
                step_result = await _click_nearby_by_label(
                    browser_session=browser_session,
                    target_label=step_label,
                    avoid_global_nav=True,
                    block_text_tokens=block_text_tokens or [],
                )
        guard = None
        clicked = bool(step_result.get("clicked"))
        if clicked:
            guard = await _guard_and_recover_unexpected_page(browser_session)
            if not guard.get("ok"):
                clicked = False
                step_result["error"] = "unexpected_global_page_after_click"
            else:
                await asyncio.sleep(0.3)

        route_trace.append(
            {
                "step_no": index,
                "target": step_label,
                "clicked": clicked,
                "selector": step_result.get("selector"),
                "error": step_result.get("error"),
                "unexpected_page_guard": guard,
            }
        )
        if not clicked:
            required_failures.append(
                {
                    "step_no": index,
                    "op": "route",
                    "target": step_label,
                    "error": step_result.get("error") or "route_click_failed",
                }
            )
            break

    return route_trace, required_failures


async def _collect_module_data_once(
    *,
    browser_session: BrowserSession,
    module: ModuleSpec,
    params: RunModuleCollectionParams,
    context: dict,
    task_dir: Path | None,
) -> dict:
    await _ensure_single_tab_guard(browser_session)
    profile = get_profile_spec(params.profile_name)
    blocked_tokens = _profile_block_text_tokens(profile)
    pre_unexpected = await _detect_unexpected_global_page(browser_session)
    if pre_unexpected.get("is_unexpected"):
        pre_guard = await _guard_and_recover_unexpected_page(browser_session)
        if not pre_guard.get("ok"):
            return {
                "touchpoint_result": None,
                "date_type_result": None,
                "fields": [],
                "files": [],
                "route_trace": [],
                "step_trace": [],
                "required_failures": [
                    {
                        "step_no": 0,
                        "op": "precheck",
                        "target": "report_detail_page",
                        "error": "unexpected_global_page_before_module_execution",
                        "guard": pre_guard,
                    }
                ],
            }

    auto_touchpoint = _infer_touchpoint_from_report_name(params.report_name)
    touchpoint = params.touchpoint or auto_touchpoint or module.default_touchpoint
    touchpoint_result = None
    date_type_result = None
    fields: list[dict] = []
    files: list[dict] = []
    route_trace: list[dict] = []
    step_trace: list[dict] = []
    required_failures: list[dict] = []

    # Always enforce DSL route_path before interaction_steps so run_module_collection
    # can execute deterministically even if agent skips navigate_module_route.
    if module.route_path:
        pre_route_trace, pre_route_failures = await _navigate_route_path_for_module(
            browser_session=browser_session,
            module=module,
            block_text_tokens=blocked_tokens,
        )
        route_trace.extend(pre_route_trace)
        if pre_route_failures:
            required_failures.extend(pre_route_failures)
            return {
                "touchpoint_result": touchpoint_result,
                "date_type_result": date_type_result,
                "fields": fields,
                "files": files,
                "route_trace": route_trace,
                "step_trace": step_trace,
                "required_failures": required_failures,
            }

    if module.interaction_steps:
        executed = await _run_module_interaction_steps(
            browser_session=browser_session,
            module=module,
            profile_name=params.profile_name,
            context=context,
            task_dir=task_dir,
            max_wait_seconds=params.max_wait_seconds,
        )
        fields = executed.get("fields", [])
        files = executed.get("files", [])
        route_trace.extend(executed.get("route_trace", []))
        step_trace = executed.get("step_trace", [])
        required_failures.extend(executed.get("required_failures", []))
        for item in step_trace:
            if item.get("op") == "select":
                target = (item.get("target") or "").strip()
                if "触点" in target and touchpoint_result is None:
                    touchpoint_result = item
                if ("日期" in target or "时间" in target) and date_type_result is None:
                    date_type_result = item
            if item.get("op") == "set_date_range" and date_type_result is None:
                date_type_result = item
    else:
        if touchpoint:
            touchpoint_result = await _select_option_by_text(
                browser_session,
                touchpoint,
                target_label="触点",
                avoid_global_nav=False,
                block_text_tokens=blocked_tokens,
            )
            await asyncio.sleep(0.3)

        if params.date_type:
            date_type_result = await _select_option_by_text(
                browser_session,
                params.date_type,
                target_label="日期",
                avoid_global_nav=False,
                block_text_tokens=blocked_tokens,
            )
            await asyncio.sleep(0.3)

        if module.acquisition_mode in {"direct", "mixed"} and module.field_keys:
            selected_field_specs = _resolve_field_specs(params.profile_name, module.field_keys)
            fields = await _extract_fields_by_specs(browser_session, selected_field_specs)

        if module.acquisition_mode in {"download", "mixed"} and module.file_keys:
            selected_file_specs = _resolve_file_specs(params.profile_name, module.file_keys)
            download_dir = task_dir if task_dir else Path(DOWNLOADS_PATH)
            download_dir.mkdir(parents=True, exist_ok=True)
            files = await _download_files_by_specs(
                browser_session=browser_session,
                specs=selected_file_specs,
                download_dir=download_dir,
                max_wait_seconds=params.max_wait_seconds,
            )

    return {
        "touchpoint_result": touchpoint_result,
        "date_type_result": date_type_result,
        "fields": fields,
        "files": files,
        "route_trace": route_trace,
        "step_trace": step_trace,
        "required_failures": required_failures,
    }


def create_opt_tools(session: TaskSession) -> Tools:
    # Modes:
    # - agent_first (default): allow generic browser actions, keep custom tools for structured extraction/download.
    # - tool_only: disable generic browser actions, force business tools.
    nav_mode = os.getenv("YUNTU_OPT_NAV_MODE", "agent_first").strip().lower()
    force_agent_dsl = os.getenv("YUNTU_OPT_FORCE_AGENT_DSL", "0").strip() == "1"
    enable_nav_tools = os.getenv("YUNTU_OPT_ENABLE_NAV_TOOLS", "0").strip() == "1"
    file_actions = [
        "write_file",
        "replace_file",
        "append_file",
        "read_file",
        "delete_file",
        "list_files",
    ]
    if nav_mode in {"tool_only", "strict"}:
        exclude_actions = [
            "search",
            "navigate",
            "go_back",
            "wait",
            "click",
            "input",
            "upload_file",
            "switch",
            "close",
            "extract",
            "scroll",
            "send_keys",
            "find_text",
            "screenshot",
            "dropdown_options",
            "select_dropdown",
            *file_actions,
        ]
    else:
        # Keep browser navigation available, but disable file-system actions
        # to avoid todo/checklist loops that can stall DSL execution.
        exclude_actions = [*file_actions]
    tools = Tools(exclude_actions=exclude_actions)
    runtime_state: dict[str, object] = {
        "extract_fail_streak": {},
        "last_extract_sig": None,
        "select_fail_streak": {},
        "last_select_sig": None,
        "download_fail_streak": {},
        "last_download_sig": None,
        "recent_hints": [],
    }

    def _append_runtime_hints(values: list[str | None]) -> None:
        hints = runtime_state.get("recent_hints")
        if not isinstance(hints, list):
            hints = []
        for value in values:
            text = str(value or "").strip()
            if not text:
                continue
            if re.fullmatch(r"-?\d+(?:[\.,]\d+)?%?", text):
                continue
            if len(text) <= 1:
                continue
            hints.append(text)
        deduped = list(dict.fromkeys([h for h in hints if str(h).strip()]))
        runtime_state["recent_hints"] = deduped[-16:]

    def _get_runtime_hints() -> list[str]:
        hints = runtime_state.get("recent_hints")
        if not isinstance(hints, list):
            return []
        return [str(h).strip() for h in hints if str(h).strip()]
    # Keep local implementation for optional fallback/reference, but do not register
    # this action in opt default flow.
    def _disabled_action(*_args, **_kwargs):
        def _decorator(func):
            return func

        return _decorator

    def _nav_action(*args, **kwargs):
        if enable_nav_tools:
            return tools.action(*args, **kwargs)
        return _disabled_action(*args, **kwargs)

    @_nav_action(
        "Open Yuntu homepage in current tab and wait for page to be ready.",
        param_model=OpenYuntuHomeParams,
    )
    async def open_yuntu_home(
        params: OpenYuntuHomeParams,
        browser_session,
    ) -> ActionResult:
        await _ensure_single_tab_guard(browser_session)
        page = await browser_session.must_get_current_page()
        try:
            await page.goto(params.url)
            await asyncio.sleep(max(0, min(params.wait_seconds, 8)))
            state = await _snapshot_page_state(browser_session)
            if state.get("is_landing") and not state.get("is_workbench"):
                enter_result = await _try_enter_workbench(browser_session)
                await asyncio.sleep(0.8)
                state = await _snapshot_page_state(browser_session)
            else:
                enter_result = {"trace": []}

            status = "ok"
            if state.get("is_login"):
                status = "login_required"
            elif state.get("is_landing") and not state.get("is_workbench"):
                status = "landing_page"

            return ActionResult(
                extracted_content=json.dumps(
                    {
                        "type": "open_yuntu_home_result",
                        "status": status,
                        "url": state.get("url"),
                        "page_state": state,
                        "enter_workbench": enter_result,
                    },
                    ensure_ascii=False,
                )
            )
        except Exception as exc:  # noqa: BLE001
            return ActionResult(
                extracted_content=json.dumps(
                    {
                        "type": "open_yuntu_home_result",
                        "status": "error",
                        "url": params.url,
                        "error": str(exc),
                    },
                    ensure_ascii=False,
                )
            )

    @_nav_action(
        "Switch current brand by opening brand selector and selecting exact brand name (with aliases fallback).",
        param_model=SwitchBrandParams,
    )
    async def switch_brand(
        params: SwitchBrandParams,
        browser_session,
    ) -> ActionResult:
        effective_profile_name = _resolve_effective_profile_name(session, params.profile_name)
        profile = get_profile_spec(effective_profile_name)
        blocked_tokens = _profile_block_text_tokens(profile)
        await _ensure_single_tab_guard(browser_session)
        page_state = await _snapshot_page_state(browser_session)
        if page_state.get("is_login"):
            return ActionResult(
                extracted_content=json.dumps(
                    {
                        "type": "switch_brand_result",
                        "profile_name": params.profile_name,
                        "status": "login_required",
                        "target_brand": params.brand_name,
                        "picked_brand": None,
                        "open_click": None,
                        "selected": None,
                        "verify": None,
                        "page_state": page_state,
                    },
                    ensure_ascii=False,
                )
            )
        if not page_state.get("is_workbench"):
            return ActionResult(
                extracted_content=json.dumps(
                    {
                        "type": "switch_brand_result",
                        "profile_name": params.profile_name,
                        "status": "not_workbench",
                        "target_brand": params.brand_name,
                        "picked_brand": None,
                        "open_click": None,
                        "selected": None,
                        "verify": None,
                        "page_state": page_state,
                    },
                    ensure_ascii=False,
                )
            )

        candidates = [params.brand_name.strip()]
        for alias in params.brand_aliases or []:
            alias_text = (alias or "").strip()
            if alias_text and alias_text not in candidates:
                candidates.append(alias_text)

        # Keep selector opening scoped to non-global selector controls to avoid
        # accidentally clicking top navigation entries such as "品牌版".
        open_click = await _click_text_step(
            browser_session=browser_session,
            label="品牌",
            selectors=[
                ".ant-select-selector",
                "[role='combobox']",
                "input[placeholder*='品牌']",
                "[class*='brand'] .ant-select-selector",
                "[class*='brand-select'] .ant-select-selector",
                ".ant-form-item .ant-select-selector",
            ],
            allow_reverse_contains=False,
            avoid_global_nav=True,
            block_text_tokens=blocked_tokens,
        )
        if not open_click.get("clicked"):
            open_click = await _click_text_step(
                browser_session=browser_session,
                label="品牌",
                selectors=[".ant-select-selector", "[role='combobox']", "label", "span", "div"],
                allow_reverse_contains=False,
                avoid_global_nav=True,
                block_text_tokens=blocked_tokens,
            )
        await asyncio.sleep(0.2)
        if not open_click.get("clicked"):
            return ActionResult(
                extracted_content=json.dumps(
                    {
                        "type": "switch_brand_result",
                        "profile_name": params.profile_name,
                        "status": "failed",
                        "target_brand": params.brand_name,
                        "picked_brand": None,
                        "open_click": open_click,
                        "selected": None,
                        "verify": None,
                        "page_state": page_state,
                    },
                    ensure_ascii=False,
                )
            )

        selected_payload = None
        verified = None
        picked_name = None
        for name in candidates:
            selected = await _select_option_by_text(
                browser_session,
                name,
                target_label="品牌",
                avoid_global_nav=True,
                block_text_tokens=blocked_tokens,
            )
            await asyncio.sleep(0.2)
            verify = await _verify_selection_applied(browser_session, "品牌", name)
            selected_payload = selected
            verified = verify
            verify_reason = str((verify or {}).get("reason") or "")
            strict_verify_ok = bool(verify.get("applied")) and verify_reason in {
                "target_container_selected_marker",
                "selected_node_contains_value",
            }
            if selected.get("selected") and strict_verify_ok:
                picked_name = name
                break

        ok = bool(picked_name)
        return ActionResult(
            extracted_content=json.dumps(
                {
                    "type": "switch_brand_result",
                    "profile_name": params.profile_name,
                    "status": "ok" if ok else "failed",
                    "target_brand": params.brand_name,
                    "picked_brand": picked_name,
                    "open_click": open_click,
                    "selected": selected_payload,
                    "verify": verified,
                    "page_state": page_state,
                },
                ensure_ascii=False,
            )
        )

    @_nav_action(
        "Navigate to report list by configured profile menu_path.",
        param_model=NavigateReportListParams,
    )
    async def navigate_report_list(
        params: NavigateReportListParams,
        browser_session,
    ) -> ActionResult:
        effective_profile_name = _resolve_effective_profile_name(session, params.profile_name)
        await _ensure_single_tab_guard(browser_session)
        page_state = await _snapshot_page_state(browser_session)
        if page_state.get("is_login"):
            return ActionResult(
                extracted_content=json.dumps(
                    {
                        "type": "navigate_report_list_result",
                        "profile_name": effective_profile_name,
                        "requested_profile_name": params.profile_name,
                        "status": "login_required",
                        "route_trace": [],
                        "page_state": page_state,
                    },
                    ensure_ascii=False,
                )
            )
        if page_state.get("is_landing") and not page_state.get("is_workbench"):
            enter_result = await _try_enter_workbench(browser_session)
            await asyncio.sleep(0.8)
            page_state = await _snapshot_page_state(browser_session)
        else:
            enter_result = {"trace": []}

        profile = get_profile_spec(effective_profile_name)
        precheck_list_state = await _is_report_list_page(browser_session, profile)
        if precheck_list_state.get("in_list"):
            return ActionResult(
                extracted_content=json.dumps(
                    {
                        "type": "navigate_report_list_result",
                        "profile_name": profile.profile_name,
                        "requested_profile_name": params.profile_name,
                        "status": "ok",
                        "route_trace": [],
                        "page_state": page_state,
                        "enter_workbench": enter_result,
                        "report_list_state": precheck_list_state,
                        "note": "already_on_report_list",
                    },
                    ensure_ascii=False,
                )
            )

        # Do not hard-stop on "not_workbench" precheck:
        # page-state heuristics can be wrong during transient rendering.
        precheck_note = None
        if not page_state.get("is_workbench"):
            precheck_note = "workbench_precheck_failed_but_navigation_attempted"

        nav_result = await _navigate_to_report_list_best_effort(browser_session, profile, rounds=2)
        status = "ok" if nav_result.get("ok") else "failed"
        return ActionResult(
            extracted_content=json.dumps(
                {
                    "type": "navigate_report_list_result",
                    "profile_name": profile.profile_name,
                    "requested_profile_name": params.profile_name,
                    "status": status,
                    "route_trace": nav_result.get("trace", []),
                    "page_state": page_state,
                    "enter_workbench": enter_result,
                    "report_list_state": nav_result.get("list_state"),
                    "precheck_note": precheck_note,
                },
                ensure_ascii=False,
            )
        )

    @tools.action(
        "Click a target text in the report content area (tab/menu/card/button). Prefer this over raw click while executing module DSL steps.",
        param_model=ClickInContentParams,
    )
    async def click_in_content(
        params: ClickInContentParams,
        browser_session,
    ) -> ActionResult:
        effective_profile_name = _resolve_effective_profile_name(session, params.profile_name)
        profile = get_profile_spec(effective_profile_name)
        blocked_tokens = _profile_block_text_tokens(profile)
        await _ensure_single_tab_guard(browser_session)
        pre_guard = await _guard_and_recover_unexpected_page(browser_session)

        selectors = [s.strip() for s in (params.selectors or []) if str(s).strip()]
        if not selectors:
            selectors = list(DEFAULT_CONTENT_CLICK_SELECTORS)

        page_before = await browser_session.must_get_current_page()
        before_url = str(getattr(page_before, "url", "") or "")

        clicked = await _click_text_step(
            browser_session=browser_session,
            label=(params.target_text or "").strip(),
            selectors=selectors,
            allow_reverse_contains=True,
            avoid_global_nav=True,
            require_clickable=True,
            block_text_tokens=blocked_tokens,
            strict_selectors_only=True,
        )
        if not clicked.get("clicked"):
            clicked = await _click_text_step(
                browser_session=browser_session,
                label=(params.target_text or "").strip(),
                selectors=selectors,
                allow_reverse_contains=True,
                avoid_global_nav=True,
                require_clickable=True,
                block_text_tokens=blocked_tokens,
                strict_selectors_only=False,
            )
        if not clicked.get("clicked"):
            clicked = await _click_nearby_by_label(
                browser_session=browser_session,
                target_label=(params.target_text or "").strip(),
                avoid_global_nav=True,
                block_text_tokens=blocked_tokens,
            )
        if not clicked.get("clicked"):
            await _scroll_main_content_to_top(browser_session)
            await asyncio.sleep(0.15)
            clicked = await _click_text_step(
                browser_session=browser_session,
                label=(params.target_text or "").strip(),
                selectors=selectors,
                allow_reverse_contains=True,
                avoid_global_nav=True,
                require_clickable=True,
                block_text_tokens=blocked_tokens,
                strict_selectors_only=True,
            )
            if not clicked.get("clicked"):
                clicked = await _click_nearby_by_label(
                    browser_session=browser_session,
                    target_label=(params.target_text or "").strip(),
                    avoid_global_nav=True,
                    block_text_tokens=blocked_tokens,
                )
        await asyncio.sleep(max(0.0, min(float(params.wait_seconds), 2.0)))

        # Guard against false-positive clicks on duplicated text nodes:
        # for tab/menu-like targets, require observable state change
        # (URL changed OR selected marker observed for target text).
        clicked_selector = str(clicked.get("selector") or "").lower()
        selector_hint_nav = any(
            tok in str(sel).lower()
            for sel in selectors
            for tok in ("role='tab'", "ant-tabs-tab", "role='menuitem'", "ant-menu-item")
        )
        target_text_norm = str(params.target_text or "").strip().lower()
        target_is_date_like = any(tok in target_text_norm for tok in ("日期", "date", "周期", "range"))
        nav_like = any(tok in clicked_selector for tok in ("role='tab'", "ant-tabs-tab", "role='menuitem'", "ant-menu-item")) or (
            selector_hint_nav and (not target_is_date_like) and (len(target_text_norm) <= 24)
        )
        nav_verify: dict | None = None
        if clicked.get("clicked") and nav_like and (params.target_text or "").strip():
            page_after = await browser_session.must_get_current_page()
            after_url = str(getattr(page_after, "url", "") or "")
            url_changed = after_url != before_url
            nav_verify = await _verify_selection_applied(
                browser_session,
                target=(params.target_text or "").strip(),
                value=(params.target_text or "").strip(),
            )
            if (not url_changed) and (not bool(nav_verify.get("applied"))):
                clicked = {
                    "clicked": False,
                    "selector": clicked.get("selector"),
                    "text": clicked.get("text"),
                    "error": f"click_not_effective:{nav_verify.get('reason') or 'unverified'}",
                }

        post_guard = await _guard_and_recover_unexpected_page(browser_session)
        if clicked.get("clicked"):
            _append_runtime_hints(
                [
                    params.target_text,
                    clicked.get("text"),
                ]
            )

        return ActionResult(
            extracted_content=json.dumps(
                {
                    "type": "click_in_content_result",
                    "profile_name": profile.profile_name,
                    "requested_profile_name": params.profile_name,
                    "target_text": params.target_text,
                    "selectors": selectors,
                    "clicked": bool(clicked.get("clicked")),
                    "selector": clicked.get("selector"),
                    "text": clicked.get("text"),
                    "error": clicked.get("error"),
                    "verify": nav_verify,
                    "page_guard": {"before": pre_guard, "after": post_guard},
                },
                ensure_ascii=False,
            )
        )

    @tools.action(
        "Hover near a target text in the report content area (usually tooltip/question icons). Use for DSL [hover] steps before hover extraction.",
        param_model=HoverInContentParams,
    )
    async def hover_in_content(
        params: HoverInContentParams,
        browser_session,
    ) -> ActionResult:
        effective_profile_name = _resolve_effective_profile_name(session, params.profile_name)
        profile = get_profile_spec(effective_profile_name)
        blocked_tokens = _profile_block_text_tokens(profile)
        await _ensure_single_tab_guard(browser_session)
        pre_guard = await _guard_and_recover_unexpected_page(browser_session)

        selectors = _sanitize_hover_selectors(params.selectors or [])
        if not selectors:
            selectors = list(DEFAULT_CONTENT_HOVER_SELECTORS)

        raw_target_text = (params.target_text or "").strip()
        hover_label = _derive_hover_anchor_text(raw_target_text)
        resolved_target_text = raw_target_text

        hovered = await _hover_text_step(
            browser_session=browser_session,
            label=raw_target_text,
            semantic_text=raw_target_text,
            selectors=selectors,
            allow_reverse_contains=False,
            avoid_global_nav=True,
            max_text_len=72,
            block_text_tokens=blocked_tokens,
            strict_selectors_only=True,
            wait_seconds=params.wait_seconds,
        )
        if not hovered.get("hovered"):
            hovered = await _hover_text_step(
                browser_session=browser_session,
                label=hover_label,
                semantic_text=raw_target_text,
                selectors=selectors,
                allow_reverse_contains=True,
                avoid_global_nav=True,
                max_text_len=96,
                block_text_tokens=blocked_tokens,
                strict_selectors_only=True,
                wait_seconds=max(0.8, float(params.wait_seconds)),
            )
            resolved_target_text = hover_label

        post_guard = await _guard_and_recover_unexpected_page(browser_session)
        if hovered.get("hovered"):
            _append_runtime_hints(
                [
                    params.target_text,
                    hovered.get("text"),
                    hovered.get("matched_label"),
                ]
            )

        return ActionResult(
            extracted_content=json.dumps(
                {
                    "type": "hover_in_content_result",
                    "profile_name": profile.profile_name,
                    "requested_profile_name": params.profile_name,
                    "target_text": params.target_text,
                    "resolved_target_text": resolved_target_text,
                    "selectors": selectors,
                    "hovered": bool(hovered.get("hovered")),
                    "selector": hovered.get("selector"),
                    "text": hovered.get("text"),
                    "matched_label": hovered.get("matched_label"),
                    "error": hovered.get("error"),
                    "page_guard": {"before": pre_guard, "after": post_guard},
                },
                ensure_ascii=False,
            )
        )

    @tools.action(
        "Select value(s) in content dropdown/cascader/radio controls. Supports single value or path-style multi-step selection.",
        param_model=SelectInContentParams,
    )
    async def select_in_content(
        params: SelectInContentParams,
        browser_session,
    ) -> ActionResult:
        effective_profile_name = _resolve_effective_profile_name(session, params.profile_name)
        profile = get_profile_spec(effective_profile_name)
        blocked_tokens = _profile_block_text_tokens(profile)
        await _ensure_single_tab_guard(browser_session)
        pre_guard = await _guard_and_recover_unexpected_page(browser_session)

        target_text = (params.target_text or "").strip() or None
        raw_value = (params.value or "").strip()
        is_date_range_value = False
        if raw_value:
            try:
                is_date_range_value = bool(
                    re.search(
                        r"\d{4}[./-]\d{1,2}[./-]\d{1,2}\s*(?:~|至|to|-)\s*\d{4}[./-]\d{1,2}[./-]\d{1,2}",
                        raw_value,
                    )
                )
            except Exception:
                is_date_range_value = False

        # For date-range selection, always prioritize value itself and ignore `path`.
        # `path` is for cascader selections (a/b/c) and can shadow date-range values.
        if is_date_range_value:
            values = [raw_value]
        else:
            values = [v.strip() for v in (params.path or []) if str(v).strip()]
            if not values and raw_value:
                values = _split_select_values(raw_value)

        if not values:
            return ActionResult(
                extracted_content=json.dumps(
                    {
                        "type": "select_in_content_result",
                        "profile_name": profile.profile_name,
                        "requested_profile_name": params.profile_name,
                        "target_text": target_text,
                        "status": "failed",
                        "error": "empty_value",
                        "target_click": None,
                        "selections": [],
                    },
                    ensure_ascii=False,
                )
            )

        select_semantic = f"{target_text or ''} {' / '.join(values)}"
        is_date_like_select = bool(
            re.search(
                r"(日期|时间|date|range|日历|calendar|自定义|按周|按月|实时|近\s*\d+\s*天)",
                select_semantic,
                flags=re.IGNORECASE,
            )
        )

        page = await browser_session.must_get_current_page()
        page_url = str(getattr(page, "url", "") or "").split("?", 1)[0]
        select_sig = (page_url, target_text or "", tuple(values))
        last_select_sig = runtime_state.get("last_select_sig")
        if last_select_sig != select_sig:
            runtime_state["last_select_sig"] = select_sig
        select_streak_map = runtime_state.get("select_fail_streak")
        if not isinstance(select_streak_map, dict):
            select_streak_map = {}
            runtime_state["select_fail_streak"] = select_streak_map
        current_select_streak = int(select_streak_map.get(select_sig, 0))
        if current_select_streak >= 2:
            return ActionResult(
                extracted_content=json.dumps(
                    {
                        "type": "select_in_content_result",
                        "profile_name": profile.profile_name,
                        "requested_profile_name": params.profile_name,
                        "target_text": target_text,
                        "values": values,
                        "status": "blocked_repeated_failure",
                        "error": "same_select_failed_multiple_times",
                        "retry_count": current_select_streak,
                        "target_click": None,
                        "selections": [],
                    },
                    ensure_ascii=False,
                )
            )

        # Date-range fast path: run date filter + strict verification first.
        if is_date_range_value and len(values) == 1:
            date_result = await _set_date_filter(
                browser_session=browser_session,
                target=target_text,
                value=values[0],
            )
            if date_result.get("status") == "ok":
                select_streak_map[select_sig] = 0
                post_guard = await _guard_and_recover_unexpected_page(browser_session)
                _append_runtime_hints([target_text, values[0], date_result.get("date_range")])
                return ActionResult(
                    extracted_content=json.dumps(
                        {
                            "type": "select_in_content_result",
                            "profile_name": profile.profile_name,
                            "requested_profile_name": params.profile_name,
                            "target_text": target_text,
                            "values": values,
                            "status": "ok",
                            "target_click": None,
                            "selections": [
                                {
                                    "step_no": 1,
                                    "value": values[0],
                                    "selected": True,
                                    "text": date_result.get("date_range") or values[0],
                                    "error": None,
                                    "verify": {"applied": True, "reason": "date_range_verified", "evidence": None},
                                    "date_fallback": date_result,
                                    "page_guard": post_guard,
                                }
                            ],
                            "page_guard": {"before": pre_guard, "after": post_guard},
                            "note": "date_range_fast_path",
                        },
                        ensure_ascii=False,
                    )
                )
            select_streak_map[select_sig] = current_select_streak + 1
            return ActionResult(
                extracted_content=json.dumps(
                    {
                        "type": "select_in_content_result",
                        "profile_name": profile.profile_name,
                        "requested_profile_name": params.profile_name,
                        "target_text": target_text,
                        "values": values,
                        "status": "failed",
                        "error": "date_range_not_applied",
                        "target_click": None,
                        "selections": [],
                        "date_result": date_result,
                        "page_guard": {"before": pre_guard},
                    },
                    ensure_ascii=False,
                )
            )

        async def _single_value_fallback(reason: str) -> ActionResult | None:
            """Generic fallback for one-value selections when target-open is ambiguous."""
            if len(values) != 1:
                return None

            fallback_value = values[0]
            overlay_selected = await _select_option_in_open_overlays(
                browser_session=browser_session,
                option_text=fallback_value,
                block_text_tokens=blocked_tokens,
            )
            if not overlay_selected.get("selected"):
                direct_selected = await _select_option_by_text(
                    browser_session=browser_session,
                    option_text=fallback_value,
                    target_label=target_text,
                    avoid_global_nav=True,
                    block_text_tokens=blocked_tokens,
                )
                if direct_selected.get("selected"):
                    overlay_selected = {
                        "selected": True,
                        "text": direct_selected.get("text") or fallback_value,
                        "selector": direct_selected.get("selector"),
                        "error": None,
                        "direct_selected": direct_selected,
                    }
            if overlay_selected.get("selected"):
                verify = (
                    await _verify_selection_applied(browser_session, target_text, fallback_value)
                    if target_text
                    else {"applied": True, "reason": "no_target_verify_bypassed", "evidence": None}
                )
                if target_text and not verify.get("applied"):
                    overlay_selected = {"selected": False, "error": "selection_not_applied_to_target"}
                else:
                    select_streak_map[select_sig] = 0
                    _append_runtime_hints([target_text, fallback_value, overlay_selected.get("text") or fallback_value])
                    return ActionResult(
                        extracted_content=json.dumps(
                            {
                                "type": "select_in_content_result",
                                "profile_name": profile.profile_name,
                                "requested_profile_name": params.profile_name,
                                "target_text": target_text,
                                "values": values,
                                "status": "ok",
                                "target_click": None,
                                "selections": [
                                    {
                                        "step_no": 1,
                                        "value": fallback_value,
                                        "selected": True,
                                        "text": overlay_selected.get("text") or fallback_value,
                                        "error": None,
                                        "verify": verify,
                                        "overlay_fallback": overlay_selected,
                                        "page_guard": {"ok": True, "recovered": False},
                                    }
                                ],
                                "page_guard": {"before": pre_guard},
                                "note": f"single_value_overlay_fallback:{reason}",
                            },
                            ensure_ascii=False,
                        )
                    )

            date_result = await _set_date_filter(browser_session, target_text, fallback_value)
            if date_result.get("status") == "ok":
                verify = (
                    await _verify_selection_applied(browser_session, target_text, fallback_value)
                    if target_text
                    else {"applied": True, "reason": "no_target_verify_bypassed", "evidence": None}
                )
                if target_text and not verify.get("applied"):
                    return None
                select_streak_map[select_sig] = 0
                _append_runtime_hints([target_text, fallback_value, str(date_result.get("date_type") or fallback_value)])
                return ActionResult(
                    extracted_content=json.dumps(
                        {
                            "type": "select_in_content_result",
                            "profile_name": profile.profile_name,
                            "requested_profile_name": params.profile_name,
                            "target_text": target_text,
                            "values": values,
                            "status": "ok",
                            "target_click": None,
                            "selections": [
                                {
                                    "step_no": 1,
                                    "value": fallback_value,
                                    "selected": True,
                                    "text": date_result.get("date_type") or fallback_value,
                                    "error": None,
                                    "verify": verify,
                                    "date_fallback": date_result,
                                    "page_guard": {"ok": True, "recovered": False},
                                }
                            ],
                            "page_guard": {"before": pre_guard},
                            "note": f"single_value_date_fallback:{reason}",
                        },
                        ensure_ascii=False,
                    )
                )
            return None

        # Generic fast-path: if the target control already has desired value, treat as success.
        if target_text and len(values) == 1:
            pre_verify = await _verify_selection_applied(browser_session, target_text, values[0])
            if pre_verify.get("applied"):
                _append_runtime_hints([target_text, *values, values[0]])
                select_streak_map[select_sig] = 0
                return ActionResult(
                    extracted_content=json.dumps(
                        {
                            "type": "select_in_content_result",
                            "profile_name": profile.profile_name,
                            "requested_profile_name": params.profile_name,
                            "target_text": target_text,
                            "values": values,
                            "status": "ok",
                            "target_click": None,
                            "selections": [
                                {
                                    "step_no": 1,
                                    "value": values[0],
                                    "selected": True,
                                    "text": values[0],
                                    "error": None,
                                    "verify": pre_verify,
                                    "page_guard": {"ok": True, "recovered": False},
                                }
                            ],
                            "page_guard": {"before": pre_guard, "after": pre_guard},
                            "note": "value_already_applied",
                        },
                        ensure_ascii=False,
                    )
                )

        # Generic direct set path for single-value select.
        # This avoids opening a wrong nearby control when there are multiple dropdowns in one area.
        if target_text and len(values) == 1:
            typed_direct = await _set_select_value_by_target(
                browser_session=browser_session,
                target=target_text,
                value=values[0],
            )
            if typed_direct.get("selected"):
                verify_direct = await _verify_selection_applied(browser_session, target_text, values[0])
                if verify_direct.get("applied"):
                    _append_runtime_hints([target_text, *values, typed_direct.get("text") or values[0]])
                    select_streak_map[select_sig] = 0
                    return ActionResult(
                        extracted_content=json.dumps(
                            {
                                "type": "select_in_content_result",
                                "profile_name": profile.profile_name,
                                "requested_profile_name": params.profile_name,
                                "target_text": target_text,
                                "values": values,
                                "status": "ok",
                                "target_click": None,
                                "selections": [
                                    {
                                        "step_no": 1,
                                        "value": values[0],
                                        "selected": True,
                                        "text": typed_direct.get("text") or values[0],
                                        "error": None,
                                        "verify": verify_direct,
                                        "typed_fallback": typed_direct,
                                        "page_guard": {"ok": True, "recovered": False},
                                    }
                                ],
                                "page_guard": {"before": pre_guard, "after": pre_guard},
                                "note": "typed_direct_set",
                            },
                            ensure_ascii=False,
                        )
                    )

        target_click = None
        if params.open_target_first and target_text:
            open_selectors = [
                "main .ant-select-selector",
                "main .ant-cascader-picker",
                "main .ant-cascader-picker-label",
                ".ant-layout-content .ant-select-selector",
                ".ant-layout-content .ant-cascader-picker",
                ".ant-layout-content .ant-cascader-picker-label",
                "[class*='content'] .ant-select-selector",
                "[class*='content'] .ant-cascader-picker",
                "[class*='content'] .ant-cascader-picker-label",
                "main [role='combobox']",
                ".ant-layout-content [role='combobox']",
            ]
            target_click = await _click_text_step(
                browser_session=browser_session,
                label=target_text,
                selectors=open_selectors,
                allow_reverse_contains=False,
                avoid_global_nav=True,
                require_clickable=True,
                max_text_len=36,
                block_text_tokens=blocked_tokens,
                strict_selectors_only=True,
                prefer_near_last_interaction=False,
            )
            if not target_click.get("clicked"):
                target_click = await _click_text_step(
                    browser_session=browser_session,
                    label=target_text,
                    selectors=list(DEFAULT_CONTENT_CLICK_SELECTORS),
                    allow_reverse_contains=False,
                    avoid_global_nav=True,
                    require_clickable=True,
                    max_text_len=48,
                    block_text_tokens=blocked_tokens,
                    strict_selectors_only=True,
                    prefer_near_last_interaction=False,
                )
            if not target_click.get("clicked"):
                target_click = await _open_nearby_select_by_label(
                    browser_session=browser_session,
                    target_label=target_text,
                    avoid_global_nav=True,
                    block_text_tokens=blocked_tokens,
                )
            if target_click.get("clicked") and target_click.get("selector") in {
                "label_container_control",
                "label_ancestor_control",
                "label_sibling_control",
                "label_trigger_fallback",
            }:
                matched_label = target_click.get("matched_label")
                if matched_label and not _is_select_label_match(target_text, matched_label):
                    target_click = {
                        "clicked": False,
                        "selector": target_click.get("selector"),
                        "text": target_click.get("text"),
                        "matched_label": matched_label,
                        "error": "target_label_mismatch",
                    }
            if target_click.get("clicked"):
                clicked_text = str(target_click.get("text") or "").strip()
                if (":" in clicked_text or "：" in clicked_text) and (not _is_select_label_match(target_text, clicked_text)):
                    target_click = {
                        "clicked": False,
                        "selector": target_click.get("selector"),
                        "text": clicked_text,
                        "matched_label": target_click.get("matched_label"),
                        "error": "target_label_mismatch",
                    }
            if target_click.get("error") == "target_label_mismatch":
                fallback = await _single_value_fallback("target_label_mismatch")
                if fallback is not None:
                    return fallback
                select_streak_map[select_sig] = current_select_streak + 1
                return ActionResult(
                    extracted_content=json.dumps(
                        {
                            "type": "select_in_content_result",
                            "profile_name": profile.profile_name,
                            "requested_profile_name": params.profile_name,
                            "target_text": target_text,
                            "values": values,
                            "status": "failed",
                            "error": "target_label_mismatch",
                            "target_click": target_click,
                            "selections": [],
                            "page_guard": {"before": pre_guard},
                        },
                        ensure_ascii=False,
                    )
                )
            await asyncio.sleep(max(0.0, min(float(params.wait_seconds), 2.0)))
            overlay_state = await _is_option_overlay_visible(browser_session)
            if not overlay_state.get("visible"):
                reopened = await _open_nearby_select_by_label(
                    browser_session=browser_session,
                    target_label=target_text,
                    avoid_global_nav=True,
                    block_text_tokens=blocked_tokens,
                )
                if reopened.get("clicked") and reopened.get("selector") in {
                    "label_container_control",
                    "label_ancestor_control",
                    "label_sibling_control",
                    "label_trigger_fallback",
                }:
                    reopened_label = reopened.get("matched_label")
                    if reopened_label and not _is_select_label_match(target_text, reopened_label):
                        reopened = {
                            "clicked": False,
                            "selector": reopened.get("selector"),
                            "text": reopened.get("text"),
                            "matched_label": reopened_label,
                            "error": "target_label_mismatch",
                        }
                if reopened.get("clicked"):
                    reopened_text = str(reopened.get("text") or "").strip()
                    if (":" in reopened_text or "：" in reopened_text) and (not _is_select_label_match(target_text, reopened_text)):
                        reopened = {
                            "clicked": False,
                            "selector": reopened.get("selector"),
                            "text": reopened_text,
                            "matched_label": reopened.get("matched_label"),
                            "error": "target_label_mismatch",
                        }
                if reopened.get("clicked"):
                    target_click = reopened
                    await asyncio.sleep(0.3)
                    overlay_state = await _is_option_overlay_visible(browser_session)
            if not target_click.get("clicked") and not overlay_state.get("visible"):
                fallback = await _single_value_fallback("target_control_not_found")
                if fallback is not None:
                    return fallback
                select_streak_map[select_sig] = current_select_streak + 1
                return ActionResult(
                    extracted_content=json.dumps(
                        {
                            "type": "select_in_content_result",
                            "profile_name": profile.profile_name,
                            "requested_profile_name": params.profile_name,
                            "target_text": target_text,
                            "values": values,
                            "status": "failed",
                            "error": "target_control_not_found",
                            "target_click": target_click,
                            "overlay_state": overlay_state,
                            "selections": [],
                            "page_guard": {"before": pre_guard},
                        },
                        ensure_ascii=False,
                    )
                )

        selections: list[dict] = []
        per_level_wait = max(0.0, min(float(params.wait_seconds), 2.0))
        if len(values) > 1:
            # Multi-level cascader paths are easier to stabilize/observe with a slightly longer gap.
            per_level_wait = max(per_level_wait, 0.45)
        for idx, value in enumerate(values):
            selected = await _select_option_by_text(
                browser_session,
                value,
                target_label=target_text if idx == 0 else None,
                avoid_global_nav=True,
                block_text_tokens=blocked_tokens,
            )
            if not selected.get("selected"):
                overlay_selected = await _select_option_in_open_overlays(
                    browser_session=browser_session,
                    option_text=value,
                    block_text_tokens=blocked_tokens,
                )
                if overlay_selected.get("selected"):
                    selected = {
                        "selected": True,
                        "text": overlay_selected.get("text") or value,
                        "error": None,
                        "overlay_selected": overlay_selected,
                    }
            if not selected.get("selected"):
                # Generic fallback: click option text strictly inside currently opened overlays.
                # This is especially important for cascader-style multi-level paths.
                overlay_click = await _click_text_step(
                    browser_session=browser_session,
                    label=value,
                    selectors=[
                        ".ant-select-dropdown .ant-select-item-option-content",
                        ".ant-select-dropdown .ant-select-item-option",
                        ".ant-dropdown .ant-dropdown-menu-item",
                        ".ant-popover .ant-dropdown-menu-item",
                        ".ant-cascader-menus .ant-cascader-menu-item",
                        ".ant-cascader-menus .ant-cascader-option",
                        ".ant-cascader-menus .ant-cascader-menu-item-content",
                        "[role='listbox'] [role='option']",
                        "[role='menu'] [role='menuitem']",
                        "[role='tree'] [role='treeitem']",
                    ],
                    allow_reverse_contains=False,
                    avoid_global_nav=True,
                    require_clickable=True,
                    max_text_len=48,
                    block_text_tokens=blocked_tokens,
                    strict_selectors_only=True,
                    prefer_near_last_interaction=False,
                )
                if overlay_click.get("clicked"):
                    selected = {
                        "selected": True,
                        "text": overlay_click.get("text") or value,
                        "error": None,
                        "overlay_fallback": overlay_click,
                    }
            if (not selected.get("selected")) and target_text and len(values) == 1 and (not is_date_like_select):
                typed_select = await _set_select_value_by_target(
                    browser_session=browser_session,
                    target=target_text,
                    value=value,
                )
                if typed_select.get("selected"):
                    selected = {
                        "selected": True,
                        "text": typed_select.get("text") or value,
                        "error": None,
                        "typed_fallback": typed_select,
                    }
            verify = {"applied": True, "reason": "skip_verify", "evidence": None}
            if (not selected.get("selected")) and target_text and len(values) == 1:
                verify_existing = await _verify_selection_applied(browser_session, target_text, value)
                if verify_existing.get("applied"):
                    selected = {"selected": True, "text": value, "error": None}
                    verify = verify_existing
            if selected.get("selected") and target_text and len(values) == 1:
                verify = await _verify_selection_applied(browser_session, target_text, value)
                if not verify.get("applied"):
                    reopened = await _open_nearby_select_by_label(
                        browser_session=browser_session,
                        target_label=target_text,
                        avoid_global_nav=True,
                        block_text_tokens=blocked_tokens,
                    )
                    if reopened.get("clicked"):
                        await asyncio.sleep(0.2)
                        selected_retry = await _select_option_by_text(
                            browser_session,
                            value,
                            target_label=target_text,
                            avoid_global_nav=True,
                            block_text_tokens=blocked_tokens,
                        )
                        if selected_retry.get("selected"):
                            selected = selected_retry
                            verify = await _verify_selection_applied(browser_session, target_text, value)
                if not verify.get("applied"):
                    selected = {
                        "selected": False,
                        "text": selected.get("text"),
                        "error": "selection_not_applied_to_target",
                    }
            post_step_guard = await _guard_and_recover_unexpected_page(browser_session)
            selections.append(
                {
                    "step_no": idx + 1,
                    "value": value,
                    "selected": bool(selected.get("selected")),
                    "text": selected.get("text"),
                    "error": selected.get("error"),
                    "verify": verify,
                    "page_guard": post_step_guard,
                }
            )
            if not selected.get("selected"):
                break
            await asyncio.sleep(per_level_wait)

        if len(values) == 1 and (not selections or not selections[0].get("selected")):
            fallback = await _single_value_fallback("post_select_failed")
            if fallback is not None:
                return fallback

        selected_count = sum(1 for item in selections if item.get("selected"))
        if selected_count == len(values):
            status = "ok"
        elif selected_count > 0:
            status = "partial"
        else:
            status = "failed"
        if status == "failed":
            select_streak_map[select_sig] = current_select_streak + 1
        else:
            select_streak_map[select_sig] = 0
        if selected_count > 0:
            selected_texts = [item.get("text") for item in selections if item.get("selected")]
            _append_runtime_hints([target_text, *values, *selected_texts])

        return ActionResult(
            extracted_content=json.dumps(
                {
                    "type": "select_in_content_result",
                    "profile_name": profile.profile_name,
                    "requested_profile_name": params.profile_name,
                    "target_text": target_text,
                    "target_click": target_click,
                    "values": values,
                    "selected_count": selected_count,
                    "status": status,
                    "selections": selections,
                    "page_guard": {"before": pre_guard},
                },
                ensure_ascii=False,
            )
        )

    @tools.action(
        "Sum search count in date range from downloaded excel. Use this when search_count must be computed from trend file."
    )
    async def sum_search_count_in_date_range(
        date_start: str,
        date_end: str,
        file_path: str | None = None,
    ) -> ActionResult:
        try:
            download_dir = Path(session.task_dir) if session.task_dir else Path(DOWNLOADS_PATH)
            if file_path:
                total = sum_search_count_from_excel(file_path, date_start, date_end)
            else:
                total = sum_search_count_from_downloads(download_dir, date_start, date_end)
            return ActionResult(
                extracted_content=json.dumps(
                    {
                        "type": "search_count_sum",
                        "date_start": date_start,
                        "date_end": date_end,
                        "total_search_count": total,
                    },
                    ensure_ascii=False,
                ),
            )
        except Exception as exc:  # noqa: BLE001
            return ActionResult(error=str(exc))

    @tools.action(
        "Extract fields by brand spec. Must be used for all required metrics. Returns JSON list with status/value/source.",
        param_model=ExtractFieldsBySpecParams,
    )
    async def extract_fields_by_spec(
        params: ExtractFieldsBySpecParams,
        browser_session,
    ) -> ActionResult:
        effective_profile_name = _resolve_effective_profile_name(session, params.profile_name)
        await _ensure_single_tab_guard(browser_session)
        page_guard = await _guard_and_recover_unexpected_page(browser_session)
        profile = get_profile_spec(effective_profile_name)
        selected_specs = _resolve_field_specs(effective_profile_name, params.field_keys)
        requested_keys = tuple(spec.field_key for spec in selected_specs)
        page = await browser_session.must_get_current_page()
        page_url = str(getattr(page, "url", "") or "").split("?", 1)[0]
        extract_sig = (page_url, requested_keys)

        last_sig = runtime_state.get("last_extract_sig")
        if last_sig != extract_sig:
            runtime_state["last_extract_sig"] = extract_sig

        streak_map = runtime_state.get("extract_fail_streak")
        if not isinstance(streak_map, dict):
            streak_map = {}
            runtime_state["extract_fail_streak"] = streak_map
        current_streak = int(streak_map.get(extract_sig, 0))
        if current_streak >= 2:
            _log_extraction_terminal(
                profile_name=profile.profile_name,
                requested_keys=list(requested_keys),
                status="blocked_repeated_failure",
                page_url=page_url,
                items=[],
                error="same_field_set_failed_multiple_times",
            )
            return ActionResult(
                extracted_content=json.dumps(
                    {
                        "profile_name": profile.profile_name,
                        "requested_profile_name": params.profile_name,
                        "type": "field_extraction_result",
                        "status": "blocked_repeated_failure",
                        "error": "same_field_set_failed_multiple_times",
                        "retry_count": current_streak,
                        "requested_field_keys": list(requested_keys),
                        "items": [],
                        "page_guard": page_guard,
                    },
                    ensure_ascii=False,
                ),
            )

        results = await _extract_fields_by_specs(browser_session, selected_specs)
        ok_count = sum(1 for item in results if item.get("status") == "ok" and item.get("value") is not None)
        if ok_count == 0:
            streak_map[extract_sig] = current_streak + 1
        else:
            streak_map[extract_sig] = 0
        _log_extraction_terminal(
            profile_name=profile.profile_name,
            requested_keys=list(requested_keys),
            status="ok" if ok_count > 0 else "failed",
            page_url=page_url,
            items=results,
            error=None,
        )

        return ActionResult(
            extracted_content=json.dumps(
                {
                    "profile_name": profile.profile_name,
                    "requested_profile_name": params.profile_name,
                    "type": "field_extraction_result",
                    "status": "ok" if ok_count > 0 else "failed",
                    "ok_count": ok_count,
                    "items": results,
                    "page_guard": page_guard,
                },
                ensure_ascii=False,
            ),
        )

    @tools.action(
        "Download files by brand spec. Must be used for all required exports. Returns JSON list with status and file path.",
        param_model=DownloadFilesBySpecParams,
    )
    async def download_files_by_spec(
        params: DownloadFilesBySpecParams,
        browser_session,
    ) -> ActionResult:
        normalized_max_wait_seconds = _cap_max_wait_seconds(params.max_wait_seconds)
        effective_profile_name = _resolve_effective_profile_name(session, params.profile_name)
        await _ensure_single_tab_guard(browser_session)
        page_guard = await _guard_and_recover_unexpected_page(browser_session)
        profile = get_profile_spec(effective_profile_name)
        requested_keys = [str(k).strip() for k in (params.file_keys or []) if str(k).strip()]
        if not requested_keys:
            _log_download_terminal(
                profile_name=profile.profile_name,
                requested_keys=[],
                status="blocked",
                items=[],
                error="empty_file_keys_not_allowed",
            )
            return ActionResult(
                extracted_content=json.dumps(
                    {
                        "profile_name": profile.profile_name,
                        "requested_profile_name": params.profile_name,
                        "type": "file_download_result",
                        "status": "blocked",
                        "error": "empty_file_keys_not_allowed",
                        "items": [],
                        "page_guard": page_guard,
                    },
                    ensure_ascii=False,
                ),
            )

        # Guard against repeated retries on the same page + file set.
        try:
            page = await browser_session.must_get_current_page()
            page_url = str(getattr(page, "url", "") or "").split("?", 1)[0]
        except Exception:
            page_url = ""
        download_sig = (page_url, tuple(requested_keys))
        last_download_sig = runtime_state.get("last_download_sig")
        if last_download_sig != download_sig:
            runtime_state["last_download_sig"] = download_sig
        download_streak_map = runtime_state.get("download_fail_streak")
        if not isinstance(download_streak_map, dict):
            download_streak_map = {}
            runtime_state["download_fail_streak"] = download_streak_map
        current_download_streak = int(download_streak_map.get(download_sig, 0))
        if current_download_streak >= 2:
            _log_download_terminal(
                profile_name=profile.profile_name,
                requested_keys=requested_keys,
                status="blocked_repeated_failure",
                items=[],
                error="same_file_set_failed_multiple_times",
            )
            return ActionResult(
                extracted_content=json.dumps(
                    {
                        "profile_name": profile.profile_name,
                        "requested_profile_name": params.profile_name,
                        "type": "file_download_result",
                        "status": "blocked_repeated_failure",
                        "error": "same_file_set_failed_multiple_times",
                        "retry_count": current_download_streak,
                        "requested_file_keys": requested_keys,
                        "items": [],
                        "page_guard": page_guard,
                    },
                    ensure_ascii=False,
                ),
            )
        selected_specs = _resolve_file_specs(effective_profile_name, requested_keys)
        if not selected_specs:
            _log_download_terminal(
                profile_name=profile.profile_name,
                requested_keys=requested_keys,
                status="blocked",
                items=[],
                error="no_matching_file_specs",
            )
            return ActionResult(
                extracted_content=json.dumps(
                    {
                        "profile_name": profile.profile_name,
                        "requested_profile_name": params.profile_name,
                        "type": "file_download_result",
                        "status": "blocked",
                        "error": "no_matching_file_specs",
                        "requested_file_keys": requested_keys,
                        "items": [],
                        "page_guard": page_guard,
                    },
                    ensure_ascii=False,
                ),
            )

        download_dir = Path(session.task_dir) if session.task_dir else Path(DOWNLOADS_PATH)
        download_dir.mkdir(parents=True, exist_ok=True)
        file_anchor_requirements = _build_file_anchor_requirements(profile)
        file_preconditions = _build_file_preconditions(profile)
        generic_hint_tokens = {
            "全部",
            "触点",
            "时间",
            "日期",
            "筛选",
            "全部内容",
            "下载",
            "导出",
            "查看",
            "详情",
        }
        runtime_hints = _get_runtime_hints()
        outputs: list[dict] = []
        debug_contexts: dict[str, dict] = {}

        for spec in selected_specs:
            req_hints = [str(v).strip() for v in file_anchor_requirements.get(spec.file_key, []) if str(v).strip()]
            precondition = file_preconditions.get(spec.file_key, {})
            terminal_anchor = next(
                (
                    hint
                    for hint in reversed(req_hints)
                    if hint
                    and hint not in generic_hint_tokens
                    and not hint.startswith(("params.", "module.", "context."))
                    and not re.fullmatch(r"-?\d+(?:[\.,]\d+)?%?", hint)
                ),
                None,
            )
            precheck = await _check_download_preconditions(
                browser_session,
                precondition,
                runtime_hints=runtime_hints,
            )
            if not precheck.get("ok", True) and isinstance(precondition, dict):
                # Fallback: trust control-level verification when UI text precheck misses SPA control values.
                missing_items = [str(v) for v in (precheck.get("missing", []) or []) if str(v)]
                still_missing: list[str] = []
                for miss in missing_items:
                    if not miss.startswith("select:"):
                        still_missing.append(miss)
                        continue
                    target_value = miss.removeprefix("select:")
                    if "=" not in target_value:
                        still_missing.append(miss)
                        continue
                    target, expected = target_value.split("=", 1)
                    target = target.strip()
                    expected = expected.strip()
                    if not target or not expected:
                        still_missing.append(miss)
                        continue
                    verify = await _verify_selection_applied(browser_session, target, expected)
                    if not verify.get("applied"):
                        still_missing.append(miss)
                if not still_missing:
                    precheck = {
                        **precheck,
                        "ok": True,
                        "missing": [],
                        "control_verify_bypass": True,
                    }
            if not precheck.get("ok", True):
                # Self-heal once: apply missing select pairs inferred from DSL precondition.
                try:
                    for pair in (precondition.get("select_pairs", []) if isinstance(precondition, dict) else []):
                        target = str((pair or {}).get("target") or "").strip()
                        value = str((pair or {}).get("value") or "").strip()
                        if not target or not value:
                            continue
                        _ = await _click_text_step(
                            browser_session=browser_session,
                            label=target,
                            selectors=CONTENT_CLICK_SELECTORS,
                            allow_reverse_contains=False,
                            avoid_global_nav=False,
                            require_clickable=True,
                            max_text_len=64,
                        )
                        await asyncio.sleep(0.2)
                        select_values = _split_select_values(value) or [value]
                        for one_value in select_values:
                            _ = await _select_option_by_text(
                                browser_session=browser_session,
                                option_text=one_value,
                                target_label=target,
                                avoid_global_nav=False,
                            )
                            await asyncio.sleep(0.2)
                except Exception:
                    pass
                precheck = await _check_download_preconditions(
                    browser_session,
                    precondition,
                    runtime_hints=_get_runtime_hints(),
                )
            if not precheck.get("ok", True):
                outputs.append(
                    {
                        "file_key": spec.file_key,
                        "status": "blocked",
                        "file_path": None,
                        "file_name": None,
                        "clicked_selector": None,
                        "error": "download_precondition_failed",
                        "missing": precheck.get("missing", []),
                    }
                )
                debug_contexts[spec.file_key] = {
                    "precondition": precondition,
                    "precheck": precheck,
                }
                continue

            context_hints = list(dict.fromkeys([*req_hints, *runtime_hints[-2:]]))
            # Keep download-context generic and robust: prefer structural anchors over volatile filter values.
            stable_req_hints = [
                h
                for h in req_hints
                if h
                and h not in generic_hint_tokens
                and "/" not in h
                and "／" not in h
                and len(str(h).strip()) >= 3
            ]
            context_hints = list(dict.fromkeys([*stable_req_hints[-6:], *runtime_hints[-2:]]))
            hard_context_hints = []
            anchor_hints = [terminal_anchor] if terminal_anchor else []
            anchor_hints = list(dict.fromkeys([h for h in anchor_hints if str(h or "").strip()]))

            context = {
                "context_hints": context_hints,
                "hard_context_hints": hard_context_hints,
                "anchor_hints": anchor_hints,
                "min_context_hits": 1 if context_hints else 0,
                "min_hard_context_hits": 0,
                "min_anchor_hits": 1 if anchor_hints else 0,
                "anchor_max_distance": 620 if anchor_hints else 0,
                "last_interaction_max_distance": 620 if terminal_anchor else 520,
            }
            one = await _download_files_by_specs(
                browser_session=browser_session,
                specs=[spec],
                download_dir=download_dir,
                max_wait_seconds=normalized_max_wait_seconds,
                download_context=context,
            )
            outputs.extend(one)
            debug_contexts[spec.file_key] = {
                "precondition": precondition,
                "precheck": precheck,
                "download_context": context,
            }

        status = "ok" if any(item.get("status") == "ok" for item in outputs) else "failed"
        if status == "ok":
            download_streak_map[download_sig] = 0
        else:
            download_streak_map[download_sig] = current_download_streak + 1
        _log_download_terminal(
            profile_name=profile.profile_name,
            requested_keys=requested_keys,
            status=status,
            items=outputs,
            error=None,
        )

        return ActionResult(
            extracted_content=json.dumps(
                {
                    "profile_name": profile.profile_name,
                    "requested_profile_name": params.profile_name,
                    "type": "file_download_result",
                    "status": status,
                    "requested_file_keys": requested_keys,
                    "items": outputs,
                    "download_context": debug_contexts,
                    "page_guard": page_guard,
                },
                ensure_ascii=False,
            ),
        )

    @_nav_action(
        "Open exact report in post-report list by report_name. It searches, matches report row, then clicks view report.",
        param_model=OpenReportByNameParams,
    )
    async def open_report_by_name(
        params: OpenReportByNameParams,
        browser_session,
    ) -> ActionResult:
        effective_profile_name = _resolve_effective_profile_name(session, params.profile_name)
        profile = get_profile_spec(effective_profile_name)
        await _ensure_single_tab_guard(browser_session)

        try:
            attempts: list[dict] = []
            evidence: dict | None = None
            recovery_trace: list[dict] = []
            ok = False
            final_search: dict = {}
            final_click: dict = {}
            matched_tab: str | None = None
            max_recover_retries = 1

            for cycle in range(max_recover_retries + 1):
                cycle_no = cycle + 1
                # Force search scope to "升级版报告" only to avoid drifting into "历史报告".
                target_tabs = _profile_report_list_tabs(profile) or ["升级版报告"]
                forced_tab = target_tabs[0] if target_tabs else "升级版报告"
                forced_tab_click = await _click_text_step(
                    browser_session=browser_session,
                    label=forced_tab,
                    selectors=["[role='tab']", ".ant-tabs-tab", ".ant-tabs-tab-btn"],
                    allow_reverse_contains=False,
                    avoid_global_nav=True,
                    require_clickable=True,
                    max_text_len=48,
                    block_text_tokens=_profile_block_text_tokens(profile),
                )
                attempts.append(
                    {
                        "cycle": cycle_no,
                        "tab": "force_scope",
                        "forced_tab": forced_tab,
                        "tab_click": forced_tab_click,
                        "ok": bool(forced_tab_click.get("clicked")),
                    }
                )
                await asyncio.sleep(0.3)

                first_try = await _open_report_once(
                    browser_session=browser_session,
                    profile=profile,
                    report_name=params.report_name,
                    wait_seconds=params.wait_seconds,
                )
                attempts.append({"cycle": cycle_no, "tab": "current", **first_try})

                ok = first_try.get("ok", False)
                final_search = first_try.get("search", {})
                final_click = first_try.get("click", {})
                matched_tab = "current"

                # Self-heal once more at action level if page drifted between steps.
                if not ok and (final_search.get("error") == "not_report_list_page"):
                    recover_nav = await _navigate_to_report_list_best_effort(browser_session, profile, rounds=2)
                    recovery_trace.append({"cycle": cycle_no, "reason": "not_report_list_page", "recover_nav": recover_nav})
                    retry_try = await _open_report_once(
                        browser_session=browser_session,
                        profile=profile,
                        report_name=params.report_name,
                        wait_seconds=params.wait_seconds,
                    )
                    attempts.append({"cycle": cycle_no, "tab": "recovered_current", **retry_try})
                    ok = retry_try.get("ok", False)
                    final_search = retry_try.get("search", {})
                    final_click = retry_try.get("click", {})
                    matched_tab = "current"
                    if not ok and (final_search.get("error") == "not_report_list_page"):
                        evidence = {
                            "reason": "not_report_list_page",
                            "search": final_search,
                            "click": final_click,
                            "recover_nav": recover_nav,
                        }
                        # Continue to next cycle to allow server-reload recovery branch.
                        if cycle < max_recover_retries:
                            server_state = await _detect_server_error_state(browser_session)
                            if server_state.get("is_error"):
                                recover = await _recover_report_list_page(browser_session)
                                recovery_trace.append(
                                    {"cycle": cycle_no, "server_state": server_state, "recover": recover}
                                )
                                await asyncio.sleep(0.5)
                                continue
                        break

                # Do not switch to other report tabs (e.g. "历史报告").
                # Keep searching only within forced "升级版报告" scope.

                if not ok:
                    manual_try = await _manual_open_report_like(
                        browser_session=browser_session,
                        report_name=params.report_name,
                        row_selectors=profile.report_row_selectors,
                        name_cell_selectors=profile.report_name_cell_selectors,
                        view_button_selectors=profile.report_view_button_selectors,
                    )
                    attempts.append(
                        {
                            "cycle": cycle_no,
                            "tab": "fallback_manual_like",
                            "search": {"filled": False, "selector": None, "error": "manual_fallback"},
                            "click": manual_try,
                            "ok": bool(manual_try.get("clicked")),
                        }
                    )
                    if manual_try.get("clicked"):
                        ok = True
                        final_click = manual_try
                        matched_tab = matched_tab if matched_tab else "current"
                    else:
                        scroll_try = await _open_report_by_scroll_scan(
                            browser_session=browser_session,
                            report_name=params.report_name,
                            row_selectors=profile.report_row_selectors,
                            name_cell_selectors=profile.report_name_cell_selectors,
                            view_button_selectors=profile.report_view_button_selectors,
                        )
                        attempts.append(
                            {
                                "cycle": cycle_no,
                                "tab": "fallback_scroll_scan",
                                "search": {"filled": False, "selector": None, "error": "scroll_scan_fallback"},
                                "click": scroll_try,
                                "ok": bool(scroll_try.get("clicked")),
                            }
                        )
                        if scroll_try.get("clicked"):
                            ok = True
                            final_click = scroll_try
                            matched_tab = matched_tab if matched_tab else "current"
                            break
                        evidence = await _collect_report_page_evidence(
                            browser_session=browser_session,
                            row_selectors=profile.report_row_selectors,
                            name_cell_selectors=profile.report_name_cell_selectors,
                        )

                if ok:
                    break

                if cycle < max_recover_retries:
                    server_state = await _detect_server_error_state(browser_session)
                    if server_state.get("is_error"):
                        recover = await _recover_report_list_page(browser_session)
                        recovery_trace.append({"cycle": cycle_no, "server_state": server_state, "recover": recover})
                        await asyncio.sleep(0.5)
                        continue
                break

            payload = {
                "profile_name": profile.profile_name,
                "requested_profile_name": params.profile_name,
                "type": "open_report_result",
                "report_name": params.report_name,
                "status": "ok" if ok else "not_found",
                "matched_tab": matched_tab if ok else None,
                "search": final_search,
                "click": final_click,
                "attempts": attempts,
                "evidence": evidence,
                "recovery_trace": recovery_trace,
            }
            return ActionResult(extracted_content=json.dumps(payload, ensure_ascii=False))
        except Exception as exc:  # noqa: BLE001
            return ActionResult(
                extracted_content=json.dumps(
                    {
                        "profile_name": profile.profile_name,
                        "requested_profile_name": params.profile_name,
                        "type": "open_report_result",
                        "report_name": params.report_name,
                        "status": "error",
                        "error": str(exc),
                    },
                    ensure_ascii=False,
                ),
            )

    @_nav_action(
        "Navigate to module route by module spec. Clicks every route step and checks ready marker selectors.",
        param_model=NavigateModuleRouteParams,
    )
    async def navigate_module_route(
        params: NavigateModuleRouteParams,
        browser_session,
    ) -> ActionResult:
        effective_profile_name = _resolve_effective_profile_name(session, params.profile_name)
        profile = get_profile_spec(effective_profile_name)
        await _ensure_single_tab_guard(browser_session)
        module = get_module_spec(effective_profile_name, params.module_key)
        if module is None:
            return ActionResult(
                extracted_content=json.dumps(
                    {
                        "profile_name": profile.profile_name,
                        "requested_profile_name": params.profile_name,
                        "type": "module_route_navigation_result",
                        "module_key": params.module_key,
                        "status": "error",
                        "error": "module_not_found",
                    },
                    ensure_ascii=False,
                ),
            )

        route_trace: list[dict] = []
        clicked_count = 0

        for index, step_label in enumerate(module.route_path, 1):
            step_result = await _click_text_step(
                browser_session=browser_session,
                label=step_label,
                selectors=module.route_selectors,
                allow_reverse_contains=False,
                avoid_global_nav=False,
                require_clickable=True,
                max_text_len=64,
                block_text_tokens=_profile_block_text_tokens(profile),
            )
            guard = None
            if step_result.get("clicked"):
                guard = await _guard_and_recover_unexpected_page(browser_session)
                if not guard.get("ok"):
                    step_result["clicked"] = False
                    step_result["error"] = "unexpected_global_page_after_click"
                else:
                    clicked_count += 1
                    await asyncio.sleep(0.35)
            route_trace.append(
                {
                    "step_no": index,
                    "target": step_label,
                    "clicked": bool(step_result.get("clicked")),
                    "selector": step_result.get("selector"),
                    "error": step_result.get("error"),
                    "unexpected_page_guard": guard,
                }
            )

        ready_ok = True
        if module.page_ready_selectors:
            page = await browser_session.must_get_current_page()
            ready_js = r"""
(selectors) => {
  if (!Array.isArray(selectors) || selectors.length === 0) return true;
  for (const sel of selectors) {
    try {
      if (document.querySelector(sel)) return true;
    } catch (_) {}
  }
  return false;
}
"""
            raw = await page.evaluate(ready_js, module.page_ready_selectors)
            ready_ok = bool(raw)

        if clicked_count == len(module.route_path) and ready_ok:
            status = "ok"
        elif clicked_count > 0:
            status = "partial"
        else:
            status = "failed"

        return ActionResult(
            extracted_content=json.dumps(
                {
                    "profile_name": profile.profile_name,
                    "requested_profile_name": params.profile_name,
                    "type": "module_route_navigation_result",
                    "module_key": module.module_key,
                    "module_label": module.label,
                    "status": status,
                    "route_trace": route_trace,
                    "ready": ready_ok,
                },
                ensure_ascii=False,
            ),
        )

    @_disabled_action(
        "Run module data collection by module spec. It supports direct extraction, file download, or mixed mode.",
        param_model=RunModuleCollectionParams,
    )
    async def run_module_collection(
        params: RunModuleCollectionParams,
        browser_session,
    ) -> ActionResult:
        if force_agent_dsl:
            return ActionResult(
                extracted_content=json.dumps(
                    {
                        "type": "module_collection_result",
                        "profile_name": params.profile_name,
                        "module_key": params.module_key,
                        "status": "disabled",
                        "error": "run_module_collection_disabled_in_agent_dsl_mode",
                    },
                    ensure_ascii=False,
                )
            )
        effective_profile_name = _resolve_effective_profile_name(session, params.profile_name)
        profile = get_profile_spec(effective_profile_name)
        await _ensure_single_tab_guard(browser_session)
        module: ModuleSpec | None = get_module_spec(effective_profile_name, params.module_key)
        if module is None:
            return ActionResult(
                extracted_content=json.dumps(
                    {
                        "profile_name": profile.profile_name,
                        "requested_profile_name": params.profile_name,
                        "type": "module_collection_result",
                        "module_key": params.module_key,
                        "status": "error",
                        "error": "module_not_found",
                    },
                    ensure_ascii=False,
                ),
            )

        normalized_max_wait_seconds = _cap_max_wait_seconds(params.max_wait_seconds)
        auto_touchpoint = _infer_touchpoint_from_report_name(params.report_name)
        touchpoint = params.touchpoint or auto_touchpoint or module.default_touchpoint
        date_range_note = params.date_range
        context = {
            "params": {
                "profile_name": effective_profile_name,
                "module_key": params.module_key,
                "report_name": params.report_name,
                "touchpoint": touchpoint,
                "auto_touchpoint": auto_touchpoint,
                "date_type": params.date_type,
                "date_range": params.date_range,
                "max_wait_seconds": normalized_max_wait_seconds,
            },
            "module": {
                "module_key": module.module_key,
                "label": module.label,
                "default_touchpoint": module.default_touchpoint,
                "touchpoint_candidates": module.touchpoint_candidates,
                "date_type_candidates": module.date_type_candidates,
                "date_range_hint": module.date_range_hint,
            },
            "context": {
                "touchpoint": touchpoint,
                "auto_touchpoint": auto_touchpoint,
                "date_type": params.date_type,
                "date_range": params.date_range,
                "report_name": params.report_name,
            },
        }

        touchpoint_result = None
        date_type_result = None

        fields: list[dict] = []
        files: list[dict] = []
        route_trace: list[dict] = []
        step_trace: list[dict] = []
        retry_trace: list[dict] = []
        required_failures: list[dict] = []
        max_server_error_retries = 1

        for attempt in range(1, max_server_error_retries + 2):
            executed = await _collect_module_data_once(
                browser_session=browser_session,
                module=module,
                params=params.model_copy(
                    update={
                        "profile_name": effective_profile_name,
                        "max_wait_seconds": normalized_max_wait_seconds,
                    }
                ),
                context=context,
                task_dir=Path(session.task_dir) if session.task_dir else None,
            )
            touchpoint_result = executed.get("touchpoint_result")
            date_type_result = executed.get("date_type_result")
            fields = executed.get("fields", [])
            files = executed.get("files", [])
            route_trace = executed.get("route_trace", [])
            step_trace = executed.get("step_trace", [])
            required_failures = executed.get("required_failures", [])

            status = _derive_collection_status(fields, files)
            if required_failures and status == "success":
                status = "partial"
            server_state = await _detect_server_error_state(browser_session)
            retry_trace.append(
                {
                    "attempt": attempt,
                    "status": status,
                    "required_failures": required_failures,
                    "server_error": server_state.get("is_error", False),
                    "server_error_code": server_state.get("code"),
                }
            )

            if status != "failed" or not server_state.get("is_error"):
                break
            if attempt > max_server_error_retries:
                break
            await _reload_current_page(browser_session)
        else:
            status = _derive_collection_status(fields, files)

        return ActionResult(
            extracted_content=json.dumps(
                {
                    "profile_name": profile.profile_name,
                    "requested_profile_name": params.profile_name,
                    "type": "module_collection_result",
                    "module_key": module.module_key,
                    "module_label": module.label,
                    "status": status,
                    "report_name": params.report_name,
                    "touchpoint": touchpoint,
                    "touchpoint_result": touchpoint_result,
                    "date_type": params.date_type,
                    "date_type_result": date_type_result,
                    "date_range": date_range_note,
                    "max_wait_seconds": normalized_max_wait_seconds,
                    "route_trace": route_trace,
                    "step_trace": step_trace,
                    "retry_trace": retry_trace,
                    "required_failures": required_failures,
                    "fields": fields,
                    "files": files,
                },
                ensure_ascii=False,
            ),
        )

    return tools

