"""
opt pipeline:
- Single-agent, tool-driven deterministic extraction flow.
"""
from __future__ import annotations

import asyncio
import json
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, urlparse

from config import (
    API_KEY,
    GEMINI_MODEL,
    PROFILE_DIR,
    STORAGE_STATE_FILE,
    TASKS_BASE_DIR,
    get_brand_english_name,
)
from intent import YuntuTask, extract_intent
from json_exporter import export_agent_output_to_json
from login import check_profile_exists, perform_manual_login
from session import TaskSession
from validator import validate_intent

from browser_use import Agent, Browser
from browser_use.llm.google import ChatGoogle
from completeness import check_completeness
from excel_processor import sum_search_count_from_excel
from schemas import (
    FieldExtractionResult,
    FileDownloadResult,
    ModuleExecutionResult,
    ReportExecutionResult,
    RouteTraceItem,
    YuntuAgentStructuredOutput,
)
from specs import build_opt_profile, resolve_profile_name
from tools import create_opt_tools

from prompt_builder import build_agent_prompt


class OptTaskPipeline:
    def __init__(self, session: TaskSession):
        self.session = session

    @staticmethod
    def _build_report_targets(task_intent: YuntuTask) -> list[str]:
        targets: list[str] = []
        if task_intent.report_name:
            targets.append(task_intent.report_name)
        for category in task_intent.brand_categories or []:
            if category.report_name_star:
                targets.append(category.report_name_star)
            if category.report_name_bid:
                targets.append(category.report_name_bid)
        return list(dict.fromkeys([r for r in targets if r]))

    @classmethod
    def _print_intent_debug(cls, task_intent: YuntuTask) -> None:
        try:
            payload = task_intent.model_dump(mode="python", exclude_none=False)
        except Exception:
            payload = {"_error": "intent.model_dump_failed"}
        report_targets = cls._build_report_targets(task_intent)
        print("[intent_parsed.json]")
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        print(f"[intent_parsed.report_targets] {report_targets}")

    def _build_tools(self, mode: str):
        old = os.environ.get("YUNTU_OPT_NAV_MODE")
        old_force = os.environ.get("YUNTU_OPT_FORCE_AGENT_DSL")
        old_enable_nav_tools = os.environ.get("YUNTU_OPT_ENABLE_NAV_TOOLS")
        try:
            os.environ["YUNTU_OPT_NAV_MODE"] = mode
            os.environ["YUNTU_OPT_FORCE_AGENT_DSL"] = "1" if mode == "agent_first" else "0"
            # Phase A uses native agent actions; navigation tools disabled by default.
            os.environ["YUNTU_OPT_ENABLE_NAV_TOOLS"] = "0"
            return create_opt_tools(self.session)
        finally:
            if old is None:
                os.environ.pop("YUNTU_OPT_NAV_MODE", None)
            else:
                os.environ["YUNTU_OPT_NAV_MODE"] = old
            if old_force is None:
                os.environ.pop("YUNTU_OPT_FORCE_AGENT_DSL", None)
            else:
                os.environ["YUNTU_OPT_FORCE_AGENT_DSL"] = old_force
            if old_enable_nav_tools is None:
                os.environ.pop("YUNTU_OPT_ENABLE_NAV_TOOLS", None)
            else:
                os.environ["YUNTU_OPT_ENABLE_NAV_TOOLS"] = old_enable_nav_tools

    @staticmethod
    def _module_supports_period(module_spec: Any) -> bool:
        steps = list(getattr(module_spec, "interaction_steps", []) or [])
        if any(str(getattr(step, "op", "") or "") == "set_date_range" for step in steps):
            return True
        for step in steps:
            for attr in ("value_from", "path_from", "when"):
                bound = str(getattr(step, attr, "") or "").lower()
                if "insight" in bound:
                    return True
        return False

    @staticmethod
    def _extract_action_names(result) -> list[str]:
        if result is None:
            return []
        try:
            names = result.action_names()
        except Exception:
            return []
        return [str(name) for name in names if name]

    @staticmethod
    def _parse_json_payload(payload: Any) -> dict[str, Any] | None:
        if isinstance(payload, dict):
            return payload
        if not isinstance(payload, str):
            return None
        text = payload.strip()
        if not text or not text.startswith("{"):
            return None
        try:
            data = json.loads(text)
        except Exception:
            return None
        return data if isinstance(data, dict) else None

    @staticmethod
    def _prefer_field_result(previous: dict[str, Any] | None, current: dict[str, Any]) -> dict[str, Any]:
        if previous is None:
            return current
        prev_ok = previous.get("status") == "ok"
        cur_ok = current.get("status") == "ok"
        if prev_ok and not cur_ok:
            return previous
        if cur_ok and not prev_ok:
            return current
        # same level: keep latest
        return current

    @staticmethod
    def _prefer_file_result(previous: dict[str, Any] | None, current: dict[str, Any]) -> dict[str, Any]:
        if previous is None:
            return current
        prev_ok = previous.get("status") == "ok"
        cur_ok = current.get("status") == "ok"
        if prev_ok and not cur_ok:
            return previous
        if cur_ok and not prev_ok:
            return current
        return current

    @classmethod
    def _parse_period_results_from_note(cls, note: str | None) -> dict[str, list[dict[str, Any]]]:
        if not note:
            return {}
        payload = cls._parse_json_payload(note)
        if payload is None:
            return {}
        raw_period_results = payload.get("period_results")
        if not isinstance(raw_period_results, dict):
            return {}
        normalized: dict[str, list[dict[str, Any]]] = {}
        for label, rows in raw_period_results.items():
            period_label = str(label).strip()
            if not period_label or not isinstance(rows, list):
                continue
            items: list[dict[str, Any]] = []
            for row in rows:
                if isinstance(row, dict):
                    items.append(row)
            if items:
                normalized[period_label] = items
        return normalized

    @staticmethod
    def _validate_period_results(raw_map: dict[str, list[dict[str, Any]]]) -> dict[str, list[FieldExtractionResult]]:
        validated: dict[str, list[FieldExtractionResult]] = {}
        for period_label, rows in raw_map.items():
            items: list[FieldExtractionResult] = []
            for row in rows:
                try:
                    items.append(FieldExtractionResult.model_validate(row))
                except Exception:
                    continue
            if items:
                validated[period_label] = items
        return validated

    @classmethod
    def _collect_tool_results(
        cls,
        extracted_contents: list[Any] | None,
    ) -> tuple[dict[str, dict[str, Any]], dict[str, dict[str, Any]]]:
        by_field_key: dict[str, dict[str, Any]] = {}
        by_file_key: dict[str, dict[str, Any]] = {}
        for entry in extracted_contents or []:
            payload = cls._parse_json_payload(entry)
            if not payload:
                continue
            payload_type = str(payload.get("type") or "")
            if payload_type == "field_extraction_result":
                for raw in payload.get("items", []) or []:
                    if not isinstance(raw, dict):
                        continue
                    key = str(raw.get("field_key") or "").strip()
                    if not key:
                        continue
                    by_field_key[key] = cls._prefer_field_result(by_field_key.get(key), raw)
            elif payload_type == "file_download_result":
                for raw in payload.get("items", []) or []:
                    if not isinstance(raw, dict):
                        continue
                    key = str(raw.get("file_key") or "").strip()
                    if not key:
                        continue
                    by_file_key[key] = cls._prefer_file_result(by_file_key.get(key), raw)
        return by_field_key, by_file_key

    @staticmethod
    def _extract_tag_content(text: str, tag: str) -> str:
        if not text:
            return ""
        pattern = rf"<{re.escape(tag)}>\s*([\s\S]*?)\s*</{re.escape(tag)}>"
        m = re.search(pattern, text, flags=re.IGNORECASE)
        if not m:
            return ""
        return str(m.group(1) or "").strip()

    @staticmethod
    def _period_label_from_url(raw_url: str | None) -> str:
        url = str(raw_url or "").strip()
        if not url:
            return ""
        try:
            parsed = urlparse(url)
            qs = parse_qs(parsed.query)
            start = (qs.get("startDate") or [""])[0]
            end = (qs.get("endDate") or [""])[0]
            if start and end:
                return f"{start} ~ {end}"
        except Exception:
            return ""
        return ""

    @staticmethod
    def _normalize_period_label(raw_label: str | None) -> str:
        text = str(raw_label or "").strip()
        if not text:
            return ""
        tokens = re.findall(r"\d{4}[./-]\d{1,2}[./-]\d{1,2}", text)
        if len(tokens) >= 2:
            def _norm(tok: str) -> str:
                y, m, d = re.split(r"[./-]", tok)
                return f"{int(y):04d}-{int(m):02d}-{int(d):02d}"
            return f"{_norm(tokens[0])} ~ {_norm(tokens[1])}"
        return text

    @staticmethod
    def _extract_period_from_text(text: str | None) -> tuple[str, str] | None:
        src = str(text or "").strip()
        if not src:
            return None
        m = re.search(
            r"(\d{4}[./-]\d{1,2}[./-]\d{1,2})\s*(?:至|~|—|-)\s*(\d{4}[./-]\d{1,2}[./-]\d{1,2})",
            src,
        )
        if not m:
            return None

        def _norm(tok: str) -> str:
            y, mo, d = re.split(r"[./-]", tok)
            return f"{int(y):04d}-{int(mo):02d}-{int(d):02d}"

        return _norm(m.group(1)), _norm(m.group(2))

    @classmethod
    def _extract_report_period_map(cls, extracted_contents: list[Any] | None) -> dict[str, tuple[str, str]]:
        out: dict[str, tuple[str, str]] = {}
        for entry in extracted_contents or []:
            payload = cls._parse_json_payload(entry)
            if not payload or str(payload.get("type") or "") != "open_report_result":
                continue
            report_name = str(payload.get("report_name") or "").strip()
            report_norm = cls._normalize_report_name_key(report_name)
            if not report_norm:
                continue

            direct_period = cls._extract_period_from_text(str(payload.get("calculation_period") or ""))
            if direct_period is not None:
                out[report_norm] = direct_period
                continue

            click_obj = payload.get("click")
            if isinstance(click_obj, dict):
                click_period = cls._extract_period_from_text(str(click_obj.get("calculation_period") or ""))
                if click_period is not None:
                    out[report_norm] = click_period
                    continue

            candidates: list[str] = []
            if isinstance(click_obj, dict):
                candidates.append(str(click_obj.get("matched_report") or ""))
                candidates.append(str(click_obj.get("matched_row_text") or ""))

            for attempt in payload.get("attempts") or []:
                if not isinstance(attempt, dict):
                    continue
                attempt_click = attempt.get("click")
                if not isinstance(attempt_click, dict):
                    continue
                attempt_period = cls._extract_period_from_text(str(attempt_click.get("calculation_period") or ""))
                if attempt_period is not None:
                    out[report_norm] = attempt_period
                    break
                candidates.append(str(attempt_click.get("matched_report") or ""))
                candidates.append(str(attempt_click.get("matched_row_text") or ""))
                for row in attempt_click.get("candidate_reports") or []:
                    candidates.append(str(row or ""))

            if report_norm in out:
                continue

            best: tuple[str, str] | None = None
            for cand in candidates:
                if not cand:
                    continue
                period = cls._extract_period_from_text(cand)
                if not period:
                    continue
                if report_name and cls._normalize_report_name_key(report_name) not in cls._normalize_report_name_key(cand):
                    if best is None:
                        best = period
                    continue
                best = period
                break
            if best is not None:
                out[report_norm] = best
        return out

    @classmethod
    def _sync_period_map_to_intent(cls, task_intent: YuntuTask | None, period_map: dict[str, tuple[str, str]] | None) -> None:
        if task_intent is None or not period_map:
            return
        try:
            normalized = {k: f"{v[0]} ~ {v[1]}" for k, v in period_map.items()}
            task_intent.report_period_map = normalized
            if task_intent.report_name:
                report_norm = cls._normalize_report_name_key(task_intent.report_name)
                if report_norm and report_norm in period_map:
                    s, e = period_map[report_norm]
                    task_intent.current_report_period = f"{s} ~ {e}"
        except Exception:
            pass

    @classmethod
    def _split_extract_query_fields(cls, query: str) -> list[str]:
        text = str(query or "").strip()
        if not text:
            return []

        found = re.findall(r"[\"'“”‘’]([^\"'“”‘’]+)[\"'“”‘’]", text)
        fields: list[str] = []
        for item in found:
            token = str(item or "").strip()
            if not token:
                continue
            if token.lower() in {"extract", "values", "value", "from", "section", "result"}:
                continue
            fields.append(token)

        if not fields:
            m = re.search(r"extract\s+the\s+values?\s+for\s+(.+?)(?:\s+from|\s*$)", text, flags=re.IGNORECASE)
            if m:
                raw = str(m.group(1) or "")
                parts = re.split(r"[,，、]| and ", raw, flags=re.IGNORECASE)
                for part in parts:
                    token = str(part or "").strip(" `*.-:;")
                    if token:
                        fields.append(token)

        dedup: list[str] = []
        seen: set[str] = set()
        for field in fields:
            norm = cls._norm_match_text(field)
            if not norm or norm in seen:
                continue
            seen.add(norm)
            dedup.append(field.strip())
        return dedup

    @staticmethod
    def _parse_extract_result_rows(result_text: str) -> list[tuple[str, str]]:
        rows: list[tuple[str, str]] = []
        if not result_text:
            return rows
        value_pattern = r"([-+]?\d[\d,]*(?:\.\d+)?\s*%?)"
        patterns = [
            re.compile(rf"^\s*[-*•]?\s*\*\*(.+?)\*\*\s*[:：]?\s*{value_pattern}\s*$"),
            re.compile(rf"^\s*[-*•]?\s*(.+?)\s*[:：]\s*{value_pattern}\s*$"),
            re.compile(rf"^\s*[-*•]?\s*\*\*(.+?)\*\*\s+{value_pattern}\s*$"),
        ]
        for raw_line in str(result_text).splitlines():
            line = re.sub(r"\s+", " ", raw_line).strip()
            if not line:
                continue
            for pattern in patterns:
                m = pattern.match(line)
                if not m:
                    continue
                label = str(m.group(1) or "").strip(" -*:`")
                value = str(m.group(2) or "").strip()
                if label and value:
                    rows.append((label, value))
                break
        return rows

    @classmethod
    def _simplify_field_label(cls, text: str | None) -> str:
        value = str(text or "").strip()
        if not value:
            return ""
        # Keep bracket content as semantic anchor, e.g. search_users_mom(搜索人数环比).
        value = re.sub(r"[\(\（]([^\)\）]*)[\)\）]", r" \1 ", value)
        value = value.replace("指标", "")
        value = re.sub(r"\b(mom|qoq|yoy)\b", " ", value, flags=re.IGNORECASE)
        value = re.sub(r"\b(core metrics?)\b", " ", value, flags=re.IGNORECASE)
        value = re.sub(r"\s+", " ", value)
        return value.strip()

    @classmethod
    def _map_native_label_to_canonical_field(
        cls,
        raw_label: str,
        canonical_keys: list[str],
        previous_metric_key: str | None = None,
    ) -> str:
        label = cls._simplify_field_label(raw_label)
        if not label:
            return ""

        norm_label = cls._norm_match_text(label)
        if not norm_label:
            return ""

        # 0) Generic ratio labels should bind to the previous metric first.
        if previous_metric_key:
            ratio_suffix = ""
            if "环比" in label:
                ratio_suffix = "环比"
            elif "同比" in label:
                ratio_suffix = "同比"
            elif "占比" in label:
                ratio_suffix = "占比"
            elif "比例" in label:
                ratio_suffix = "比例"
            if ratio_suffix:
                prev_tokens = cls._field_tokens(previous_metric_key)
                has_prev_token = any(tok in norm_label for tok in prev_tokens if tok)
                has_any_metric_token = any(
                    tok in norm_label
                    for key in canonical_keys
                    for tok in cls._field_tokens(key)
                    if tok and tok not in {"环比", "同比", "占比", "比例"}
                )
                if (not has_prev_token) and (not has_any_metric_token):
                    candidate = f"{previous_metric_key}{ratio_suffix}"
                    if candidate in canonical_keys:
                        return candidate

        canonical_norm: dict[str, str] = {key: cls._norm_match_text(cls._simplify_field_label(key)) for key in canonical_keys}

        # 1) exact / contains matching to canonical keys
        best_key = ""
        best_score = 0
        for key, norm_key in canonical_norm.items():
            if not norm_key:
                continue
            score = 0
            if norm_key == norm_label:
                score += 200
            if norm_key in norm_label or norm_label in norm_key:
                score += 120
            tokens = cls._field_tokens(key)
            overlap = [tok for tok in tokens if tok and tok in norm_label]
            score += len(overlap) * 25
            score += sum(min(len(tok), 6) for tok in overlap)
            if score > best_score:
                best_score = score
                best_key = key

        if best_key and best_score > 0:
            return best_key

        # 2) Fallback ratio mapping.
        is_ratio_label = any(token in label for token in ("环比", "同比", "占比", "比例"))
        if is_ratio_label and previous_metric_key:
            suffix = ""
            if "环比" in label:
                suffix = "环比"
            elif "同比" in label:
                suffix = "同比"
            elif "占比" in label:
                suffix = "占比"
            elif "比例" in label:
                suffix = "比例"
            if suffix:
                candidate = f"{previous_metric_key}{suffix}"
                if candidate in canonical_keys:
                    return candidate

        return ""

    @classmethod
    def _match_extract_rows_to_fields(
        cls,
        query_fields: list[str],
        result_rows: list[tuple[str, str]],
    ) -> dict[str, tuple[str, str]]:
        matches: dict[str, tuple[str, str]] = {}
        used_idx: set[int] = set()

        def score(field: str, label: str) -> int:
            fn = cls._norm_match_text(field)
            ln = cls._norm_match_text(label)
            if not fn or not ln:
                return 0
            s = 0
            if fn in ln or ln in fn:
                s += 100
            field_tokens = cls._field_tokens(field)
            overlap = [tok for tok in field_tokens if tok and tok in ln]
            s += len(overlap) * 20
            s += sum(min(len(tok), 6) for tok in overlap)
            return s

        for field in query_fields:
            best_idx = -1
            best_score = 0
            for idx, (label, _value) in enumerate(result_rows):
                if idx in used_idx:
                    continue
                s = score(field, label)
                if s > best_score:
                    best_score = s
                    best_idx = idx
            if best_idx >= 0 and best_score > 0:
                used_idx.add(best_idx)
                matches[field] = result_rows[best_idx]

        ratio_fields = [
            field
            for field in query_fields
            if field not in matches
            and any(token in str(field) for token in ("环比", "同比", "率", "占比", "比例"))
        ]
        ratio_candidates = [
            (idx, row)
            for idx, row in enumerate(result_rows)
            if idx not in used_idx and any(token in row[0] for token in ("环比", "同比", "率", "占比", "比例"))
        ]
        for field in ratio_fields:
            if not ratio_candidates:
                break
            idx, row = ratio_candidates.pop(0)
            used_idx.add(idx)
            matches[field] = row

        return matches

    @classmethod
    def _collect_native_extract_results(
        cls,
        profile,
        extracted_contents: list[Any] | None,
        default_period_label: str | None = None,
    ) -> tuple[dict[str, dict[str, Any]], dict[str, dict[str, list[dict[str, Any]]]]]:
        by_field_key: dict[str, dict[str, Any]] = {}
        by_module_period: dict[str, dict[str, dict[str, dict[str, Any]]]] = {}

        modules_by_field: dict[str, list[str]] = {}
        period_enabled_modules: set[str] = set()
        for module in profile.module_specs:
            for field_key in getattr(module, "field_keys", []) or []:
                key = str(field_key).strip()
                if not key:
                    continue
                modules_by_field.setdefault(key, []).append(module.module_key)
            steps = list(getattr(module, "interaction_steps", []) or [])
            if cls._module_supports_period(module):
                period_enabled_modules.add(module.module_key)
        canonical_field_keys = list(modules_by_field.keys())

        for entry in extracted_contents or []:
            if not isinstance(entry, str):
                continue
            text = str(entry).strip()
            if not text or "<query>" not in text.lower() or "<result>" not in text.lower():
                continue

            query_text = cls._extract_tag_content(text, "query")
            result_text = cls._extract_tag_content(text, "result")
            if not result_text:
                continue

            query_fields = cls._split_extract_query_fields(query_text)

            rows = cls._parse_extract_result_rows(result_text)
            if not rows:
                continue

            period_label = cls._period_label_from_url(cls._extract_tag_content(text, "url"))
            matched_rows: dict[str, tuple[str, str]] = {}

            # Prefer canonical mapping from result labels (works for variants like "搜索次数（指数）" -> "搜索次数")
            previous_metric_key = ""
            for label, value_text in rows:
                ratio_suffix = ""
                if "环比" in label:
                    ratio_suffix = "环比"
                elif "同比" in label:
                    ratio_suffix = "同比"
                elif "占比" in label:
                    ratio_suffix = "占比"
                elif "比例" in label:
                    ratio_suffix = "比例"

                canonical_key = cls._map_native_label_to_canonical_field(
                    label,
                    canonical_field_keys,
                    previous_metric_key=previous_metric_key or None,
                )
                if not canonical_key:
                    continue
                if ratio_suffix and canonical_key in matched_rows:
                    preferred = f"{previous_metric_key}{ratio_suffix}" if previous_metric_key else ""
                    if preferred and preferred in canonical_field_keys and preferred not in matched_rows:
                        canonical_key = preferred
                    else:
                        alt_ratio_keys = [
                            key
                            for key in canonical_field_keys
                            if key.endswith(ratio_suffix) and key not in matched_rows
                        ]
                        if alt_ratio_keys:
                            canonical_key = alt_ratio_keys[0]
                if not any(token in canonical_key for token in ("环比", "同比", "占比", "比例")):
                    previous_metric_key = canonical_key
                matched_rows[canonical_key] = (label, value_text)

            # Fallback: query-to-result matching if canonical mapping not enough.
            if query_fields:
                query_matched = cls._match_extract_rows_to_fields(query_fields, rows)
                for field_key, pair in query_matched.items():
                    canonical_key = cls._map_native_label_to_canonical_field(field_key, canonical_field_keys)
                    if canonical_key:
                        matched_rows.setdefault(canonical_key, pair)

            if not matched_rows:
                continue

            for field_key, (label, value_text) in matched_rows.items():
                row = {
                    "field_key": field_key,
                    "status": "ok",
                    "source": "auto",
                    "value": cls._parse_token_value(value_text),
                    "raw_text": f"{label} {value_text}",
                    "error": None,
                }
                by_field_key[field_key] = cls._prefer_field_result(by_field_key.get(field_key), row)
                for module_key in modules_by_field.get(field_key, []):
                    if module_key not in period_enabled_modules:
                        continue
                    module_period = by_module_period.setdefault(module_key, {})
                    period_key = period_label or cls._normalize_period_label(default_period_label) or "default"
                    period_rows = module_period.setdefault(period_key, {})
                    period_rows[field_key] = cls._prefer_field_result(period_rows.get(field_key), row)

        normalized_period_map: dict[str, dict[str, list[dict[str, Any]]]] = {}
        for module_key, period_map in by_module_period.items():
            normalized_period_map[module_key] = {}
            for period_label, rows in period_map.items():
                normalized_period_map[module_key][period_label] = list(rows.values())
        return by_field_key, normalized_period_map

    @classmethod
    def _collect_module_results(cls, extracted_contents: list[Any] | None) -> dict[str, dict[str, Any]]:
        by_module_key: dict[str, dict[str, Any]] = {}
        for entry in extracted_contents or []:
            payload = cls._parse_json_payload(entry)
            if not payload:
                continue
            if str(payload.get("type") or "") != "module_collection_result":
                continue
            module_key = str(payload.get("module_key") or "").strip()
            if not module_key:
                continue
            by_module_key[module_key] = payload
        return by_module_key

    @classmethod
    def _infer_default_period_label_from_intent(cls, task_intent: Any | None) -> str | None:
        if task_intent is None:
            return None
        data: dict[str, Any] = {}
        try:
            if hasattr(task_intent, "model_dump"):
                data = task_intent.model_dump(mode="python", exclude_none=False) or {}
            elif isinstance(task_intent, dict):
                data = task_intent
        except Exception:
            data = {}
        ranges = data.get("insight_time_ranges") or []
        if isinstance(ranges, list):
            for item in ranges:
                if not isinstance(item, dict):
                    continue
                label = cls._normalize_period_label(item.get("insight_date_range"))
                if label:
                    return label
        label = cls._normalize_period_label(data.get("kol_content_date_range"))
        if label:
            return label
        return None

    @staticmethod
    def _period_results_has_items(period_map: dict[str, Any] | None) -> bool:
        if not isinstance(period_map, dict):
            return False
        for rows in period_map.values():
            if isinstance(rows, list) and len(rows) > 0:
                return True
        return False

    @classmethod
    def _compact_period_results(cls, period_map: Any) -> dict[str, list[FieldExtractionResult]]:
        if not isinstance(period_map, dict):
            return {}
        compacted: dict[str, list[FieldExtractionResult]] = {}
        for label, rows in period_map.items():
            key = str(label or "").strip()
            if not key or not isinstance(rows, list):
                continue
            valid_rows: list[FieldExtractionResult] = []
            for row in rows:
                try:
                    valid_rows.append(FieldExtractionResult.model_validate(row))
                except Exception:
                    continue
            if valid_rows:
                compacted[key] = valid_rows
        return compacted

    @classmethod
    def _rename_default_period_label(
        cls,
        period_map: dict[str, list[FieldExtractionResult]] | None,
        default_period_label: str | None,
    ) -> dict[str, list[FieldExtractionResult]]:
        compacted = cls._compact_period_results(period_map or {})
        target = cls._normalize_period_label(default_period_label)
        if not target:
            return compacted
        default_rows = list(compacted.get("default") or [])
        if not default_rows:
            return compacted
        merged = dict(compacted)
        existed = list(merged.get(target) or [])
        merged[target] = [*existed, *default_rows]
        merged.pop("default", None)
        return merged

    @classmethod
    def _normalize_report_name_key(cls, report_name: str | None) -> str:
        text = str(report_name or "").strip().lower()
        if not text:
            return ""
        text = re.sub(r"\s+", "", text)
        text = re.sub(r"[\.。,_\-—–:：;；/\\\(\)\[\]\{\}\"'`]+", "", text)
        return text

    @staticmethod
    def _sanitize_category_name(raw: str | None) -> str | None:
        text = str(raw or "").strip()
        if not text:
            return None
        m = re.search(r"品牌分类[:：]\s*([^;；\n]+)", text)
        if m:
            text = str(m.group(1) or "").strip()
        # Remove obvious noisy prefixes.
        text = re.sub(r"^\s*品牌[:：]\s*", "", text)
        text = re.sub(r"^\s*品牌分类[:：]\s*", "", text)
        # If whole sentence-like chunk leaks in, keep the first clean token.
        text = re.split(r"[;；\n]", text)[0].strip()
        if not text:
            return None
        # Prefer concise chinese token before report/amount keywords.
        m = re.match(r"^([\u4e00-\u9fffA-Za-z0-9_-]{1,20})", text)
        if m:
            candidate = m.group(1).strip()
            bad_tokens = {"品牌", "品牌分类", "报告", "星图", "竞价"}
            if candidate and candidate not in bad_tokens:
                return candidate
        # Fallback: trim very long noisy values.
        if len(text) > 30:
            return None
        return text

    @staticmethod
    def _is_unknown_report_name(report_name: str | None) -> bool:
        name = str(report_name or "").strip().lower()
        if not name:
            return True
        return name in {"unknown_report", "unknown", "n/a", "none", "null"}

    @classmethod
    def _ensure_report_targets_in_structured_output(
        cls,
        structured_output: YuntuAgentStructuredOutput,
        report_targets: list[str] | None,
    ) -> YuntuAgentStructuredOutput:
        targets = [str(t).strip() for t in (report_targets or []) if str(t).strip()]
        if not targets:
            return structured_output

        dedup_targets: list[str] = []
        seen_targets: set[str] = set()
        for t in targets:
            norm_t = cls._normalize_report_name_key(t)
            if not norm_t or norm_t in seen_targets:
                continue
            seen_targets.add(norm_t)
            dedup_targets.append(t)
        if not dedup_targets:
            return structured_output

        reports = list(structured_output.reports or [])
        if not reports:
            reports = [
                ReportExecutionResult(
                    report_name=t,
                    report_type="unknown",
                    status="failed",
                    modules=[],
                    note="No structured module data captured for this report in this run.",
                )
                for t in dedup_targets
            ]
            return structured_output.model_copy(update={"reports": reports})

        # Common fallback: one unknown report with data while intent has multiple targets.
        if len(reports) == 1 and cls._is_unknown_report_name(reports[0].report_name):
            first = reports[0].model_copy(update={"report_name": dedup_targets[0]})
            rebuilt: list[ReportExecutionResult] = [first]
            for t in dedup_targets[1:]:
                rebuilt.append(
                    ReportExecutionResult(
                        report_name=t,
                        report_type="unknown",
                        status="failed",
                        modules=[],
                        note="No structured module data captured for this report in this run.",
                    )
                )
            return structured_output.model_copy(update={"reports": rebuilt})

        existing_norms = {
            cls._normalize_report_name_key(r.report_name)
            for r in reports
            if not cls._is_unknown_report_name(r.report_name)
        }
        rebuilt = list(reports)
        for t in dedup_targets:
            norm_t = cls._normalize_report_name_key(t)
            if not norm_t or norm_t in existing_norms:
                continue
            rebuilt.append(
                ReportExecutionResult(
                    report_name=t,
                    report_type="unknown",
                    status="failed",
                    modules=[],
                    note="No structured module data captured for this report in this run.",
                )
            )
            existing_norms.add(norm_t)
        return structured_output.model_copy(update={"reports": rebuilt})

    @classmethod
    def _resolve_brand_category_for_report(
        cls,
        task_intent: YuntuTask | None,
        report_name: str | None,
    ) -> str | None:
        if task_intent is None:
            return None
        categories = getattr(task_intent, "brand_categories", None) or []
        if not categories:
            return None
        target = cls._normalize_report_name_key(report_name)
        if not target:
            return None
        for category in categories:
            category_name = cls._sanitize_category_name(getattr(category, "category_name", None))
            candidates = [
                str(getattr(category, "report_name_star", "") or "").strip(),
                str(getattr(category, "report_name_bid", "") or "").strip(),
            ]
            for candidate in candidates:
                cand_norm = cls._normalize_report_name_key(candidate)
                if not cand_norm:
                    continue
                if cand_norm == target or cand_norm in target or target in cand_norm:
                    return category_name
        return None

    @staticmethod
    def _to_numeric(value: Any) -> float | None:
        if value is None:
            return None
        if isinstance(value, (int, float)):
            return float(value)
        text = str(value).strip()
        if not text:
            return None
        text = text.replace(",", "").replace("，", "")
        m = re.search(r"-?\d+(?:\.\d+)?", text)
        if not m:
            return None
        try:
            num = float(m.group(0))
        except Exception:
            return None
        if "%" in str(value):
            return num / 100.0
        return num

    @staticmethod
    def _safe_div(numerator: float | None, denominator: float | None) -> float | None:
        if numerator is None or denominator is None:
            return None
        if abs(denominator) < 1e-12:
            return None
        return numerator / denominator

    @classmethod
    def _resolve_actual_consumption_for_report(
        cls,
        task_intent: YuntuTask | None,
        report_name: str | None,
    ) -> float | None:
        if task_intent is None:
            return None
        target = cls._normalize_report_name_key(report_name)
        if not target:
            return None

        categories = list(getattr(task_intent, "brand_categories", None) or [])
        for category in categories:
            star_name = str(getattr(category, "report_name_star", "") or "").strip()
            bid_name = str(getattr(category, "report_name_bid", "") or "").strip()
            star_norm = cls._normalize_report_name_key(star_name)
            bid_norm = cls._normalize_report_name_key(bid_name)
            if star_norm and (star_norm == target or star_norm in target or target in star_norm):
                amt = getattr(category, "consumption_amount_star", None)
                return float(amt) if amt is not None else None
            if bid_norm and (bid_norm == target or bid_norm in target or target in bid_norm):
                amt = getattr(category, "consumption_amount_bid", None)
                return float(amt) if amt is not None else None

        report_text = str(report_name or "")
        star_amt = getattr(task_intent, "consumption_amount_star", None)
        bid_amt = getattr(task_intent, "consumption_amount_bid", None)
        if "星图" in report_text and star_amt is not None:
            return float(star_amt)
        if "竞价" in report_text and bid_amt is not None:
            return float(bid_amt)
        if star_amt is not None and bid_amt is None:
            return float(star_amt)
        if bid_amt is not None and star_amt is None:
            return float(bid_amt)
        return None

    @classmethod
    def _pick_first_module_payload(cls, row: dict[str, Any], labels: list[str]) -> dict[str, Any] | None:
        for label in labels:
            payload = row.get(label)
            if isinstance(payload, dict):
                return payload
        return None

    @classmethod
    def _apply_computed_metrics_for_report_row(
        cls,
        *,
        row: dict[str, Any],
        report_name: str | None,
        task_intent: YuntuTask | None,
        report_period_map: dict[str, tuple[str, str]] | None = None,
        output_dir: Path | None = None,
    ) -> None:
        actual_spend = cls._resolve_actual_consumption_for_report(task_intent, report_name)

        overview_key = "项目整体Overview"
        overview = row.get(overview_key)
        if not isinstance(overview, dict):
            overview = {}
        row[overview_key] = overview

        flow = cls._pick_first_module_payload(row, ["5A人群资产流转", "5A人群资产流转-触点口径"])

        exposure_count = cls._to_numeric(overview.get("曝光次数"))
        interaction_rate = cls._to_numeric(overview.get("互动率"))
        completion_rate = cls._to_numeric(overview.get("完播率"))
        interaction_count = cls._to_numeric(overview.get("互动量")) or cls._to_numeric(overview.get("互动次数"))
        conversion_amount = cls._to_numeric(overview.get("本次活动转化金额")) or cls._to_numeric(overview.get("转化金额"))

        a3_count = None
        if isinstance(flow, dict):
            a3_count = cls._to_numeric(flow.get("A3流转人数")) or cls._to_numeric(flow.get("A3流转人群"))
        if a3_count is None:
            a3_count = cls._to_numeric(overview.get("A3流转人数")) or cls._to_numeric(overview.get("A3流转人群"))

        cpm = None
        cpe = None
        cps = None
        cpa3 = None
        roi = None
        interaction_count_calc = None
        completion_count_calc = None
        if exposure_count is not None and interaction_rate is not None:
            interaction_count_calc = exposure_count * interaction_rate
        if exposure_count is not None and completion_rate is not None:
            completion_count_calc = exposure_count * completion_rate
        if interaction_count is None and interaction_count_calc is not None:
            interaction_count = interaction_count_calc

        search_count = cls._to_numeric(overview.get("回搜次数"))
        if search_count is None:
            period: tuple[str, str] | None = None
            report_norm = cls._normalize_report_name_key(report_name)
            if report_period_map and report_norm:
                period = report_period_map.get(report_norm)
            if period is not None:
                start_date, end_date = period
                row["计算周期"] = f"{start_date} ~ {end_date}"
                print(f"[计算周期]: {start_date} ~ {end_date}")
                search_file_path: str | None = None
                files = overview.get("files")
                if isinstance(files, list):
                    for fp in files:
                        candidate = str(fp or "").strip()
                        if candidate.lower().endswith((".csv", ".xlsx", ".xls")):
                            search_file_path = candidate
                            break

                if not search_file_path and output_dir and output_dir.exists():
                    candidates = sorted(
                        [
                            p
                            for p in output_dir.glob("**/*")
                            if p.is_file() and p.suffix.lower() in {".csv", ".xlsx", ".xls"}
                        ],
                        key=lambda p: p.stat().st_mtime,
                        reverse=True,
                    )
                    for p in candidates:
                        name = p.name.lower()
                        if any(token in name for token in ("搜索", "趋势", "search")):
                            search_file_path = str(p)
                            break
                    if not search_file_path and candidates:
                        search_file_path = str(candidates[0])

                if search_file_path:
                    try:
                        summed = sum_search_count_from_excel(search_file_path, start_date, end_date)
                        search_count = float(summed)
                        overview["回搜次数"] = int(summed)
                        print(
                            "[search_count_calc]",
                            {
                                "report_name": report_name,
                                "period": f"{start_date} ~ {end_date}",
                                "file_path": search_file_path,
                                "total_search_count": int(summed),
                            },
                        )
                    except Exception:
                        print(
                            "[search_count_calc_failed]",
                            {
                                "report_name": report_name,
                                "period": f"{start_date} ~ {end_date}",
                                "file_path": search_file_path,
                            },
                        )
            else:
                print(
                    "[search_count_period_missing]",
                    {"report_name": report_name, "report_norm": report_norm},
                )

        if actual_spend is not None:
            cpm = cls._safe_div(actual_spend * 1000.0, exposure_count)
            cpe = cls._safe_div(actual_spend, interaction_count)
            cps = cls._safe_div(actual_spend, search_count)
            cpa3 = cls._safe_div(actual_spend, a3_count)
            roi = cls._safe_div(conversion_amount, actual_spend)
            overview["实际消耗金额"] = round(actual_spend, 6)

        if interaction_count_calc is not None:
            overview["互动次数"] = int(round(interaction_count_calc))
            if "互动量" not in overview:
                overview["互动量"] = int(round(interaction_count_calc))
        if completion_count_calc is not None:
            overview["完播数"] = int(round(completion_count_calc))
        if cpm is not None:
            overview["CPM"] = round(cpm, 6)
        if cpe is not None:
            overview["CPE"] = round(cpe, 6)
        if cps is not None:
            overview["CPS"] = round(cps, 6)
        if cpa3 is not None:
            overview["CPA3"] = round(cpa3, 6)
        if roi is not None:
            overview["ROI"] = round(roi, 6)

    @staticmethod
    def _module_business_payload(module: ModuleExecutionResult) -> dict[str, Any]:
        payload: dict[str, Any] = {}
        field_values: dict[str, Any] = {}
        for item in module.fields:
            if item.status == "ok" and item.value is not None:
                field_values[item.field_key] = item.value
        if field_values:
            payload.update(field_values)

        period_values: dict[str, dict[str, Any]] = {}
        for period_label, rows in (module.period_results or {}).items():
            row_map: dict[str, Any] = {}
            for row in rows:
                if row.status == "ok" and row.value is not None:
                    row_map[row.field_key] = row.value
            if row_map:
                period_values[str(period_label)] = row_map
        if period_values:
            payload.update(period_values)

        file_paths: list[str] = []
        for item in module.files:
            if item.status == "ok" and item.file_path:
                file_paths.append(str(item.file_path))
        if file_paths:
            payload["files"] = file_paths
        return payload

    @classmethod
    def _build_business_summary_json(
        cls,
        *,
        structured_output: YuntuAgentStructuredOutput,
        profile,
        task_intent: YuntuTask | None,
        report_period_map: dict[str, tuple[str, str]] | None = None,
        output_dir: Path | None = None,
    ) -> dict[str, Any]:
        report_targets = cls._build_report_targets(task_intent) if task_intent is not None else []
        structured_output = cls._ensure_report_targets_in_structured_output(structured_output, report_targets)

        module_specs = list(getattr(profile, "module_specs", []) or [])
        report_specs = [m for m in module_specs if getattr(m, "requires_report", False)]
        global_specs = [m for m in module_specs if not getattr(m, "requires_report", False)]

        global_map = {m.module_key: m for m in (structured_output.global_modules or [])}

        report_rows: list[dict[str, Any]] = []
        reports = list(structured_output.reports or [])
        if not reports:
            if report_targets:
                reports = [
                    ReportExecutionResult(
                        report_name=target,
                        report_type="unknown",
                        status="failed",
                        modules=[],
                    )
                    for target in report_targets
                ]
            else:
                reports = [
                    ReportExecutionResult(
                        report_name=str(getattr(task_intent, "report_name", "") or ""),
                        report_type="unknown",
                        status="failed",
                        modules=[],
                    )
                ]

        for report in reports:
            brand_value = (
                str(getattr(structured_output, "switched_brand", "") or "").strip()
                or str(getattr(task_intent, "brand_name", "") or "").strip()
                or None
            )
            row: dict[str, Any] = {
                "品牌": brand_value,
                "品牌分类": cls._resolve_brand_category_for_report(task_intent, report.report_name),
                "报告名": report.report_name,
            }
            report_map = {m.module_key: m for m in (report.modules or [])}

            for spec in report_specs:
                module = report_map.get(spec.module_key)
                if module is None:
                    continue
                module_payload = cls._module_business_payload(module)
                if module_payload:
                    row[spec.label] = module_payload

            for spec in global_specs:
                module = global_map.get(spec.module_key)
                if module is None:
                    continue
                module_payload = cls._module_business_payload(module)
                if module_payload:
                    row[spec.label] = module_payload

            cls._apply_computed_metrics_for_report_row(
                row=row,
                report_name=report.report_name,
                task_intent=task_intent,
                report_period_map=report_period_map,
                output_dir=output_dir,
            )
            report_rows.append(row)

        return {"结案报告": report_rows}

    @staticmethod
    def _safe_filename_component(text: str | None) -> str:
        value = str(text or "").strip()
        if not value:
            return "UNKNOWN_REPORT"
        value = re.sub(r"[<>:\"/\\|?*\x00-\x1F]+", "_", value)
        value = re.sub(r"\s+", "_", value).strip("._ ")
        return value[:120] or "UNKNOWN_REPORT"

    @staticmethod
    def _merge_value_score(value: Any) -> int:
        if value in (None, "", [], {}):
            return 0
        if isinstance(value, (int, float)):
            return 1 if float(value) == 0.0 else 2
        if isinstance(value, str):
            text = value.strip()
            if not text:
                return 0
            compact = text.replace(",", "").replace(" ", "")
            if compact in {"0", "0.0", "0.00", "0%", "0.0%", "0.00%"}:
                return 1
            return 2
        return 2

    @classmethod
    def _deep_merge_dict(cls, base: dict[str, Any], extra: dict[str, Any]) -> dict[str, Any]:
        merged = dict(base)
        for key, val in (extra or {}).items():
            if key not in merged:
                merged[key] = val
                continue
            prev = merged.get(key)
            if isinstance(prev, dict) and isinstance(val, dict):
                merged[key] = cls._deep_merge_dict(prev, val)
                continue
            if isinstance(prev, list) and isinstance(val, list):
                merged[key] = list(dict.fromkeys([*prev, *val]))
                continue
            prev_score = cls._merge_value_score(prev)
            new_score = cls._merge_value_score(val)
            if prev_score == 0 or new_score > prev_score or (new_score == prev_score and new_score >= 2):
                merged[key] = val
        return merged

    @classmethod
    def _update_report_progress_map(cls, output_dir: Path, business_summary: dict[str, Any]) -> None:
        progress_path = output_dir / "report_progress_map.json"
        existing: dict[str, Any] = {}
        if progress_path.exists():
            try:
                existing = json.loads(progress_path.read_text(encoding="utf-8"))
            except Exception:
                existing = {}
        reports_state = existing.get("reports")
        if not isinstance(reports_state, dict):
            reports_state = {}

        rows = business_summary.get("结案报告")
        if not isinstance(rows, list):
            rows = []
        for row in rows:
            if not isinstance(row, dict):
                continue
            report_name = str(row.get("报告名") or "").strip()
            norm = cls._normalize_report_name_key(report_name)
            if not norm:
                continue
            prev_row = reports_state.get(norm)
            if isinstance(prev_row, dict):
                reports_state[norm] = cls._deep_merge_dict(prev_row, row)
            else:
                reports_state[norm] = row

        payload = {
            "updated_at": datetime.now().isoformat(timespec="seconds"),
            "reports": reports_state,
        }
        progress_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    @classmethod
    def _merge_business_summary_with_progress(cls, output_dir: Path, business_summary: dict[str, Any]) -> dict[str, Any]:
        progress_path = output_dir / "report_progress_map.json"
        if not progress_path.exists():
            return business_summary
        try:
            payload = json.loads(progress_path.read_text(encoding="utf-8"))
        except Exception:
            return business_summary

        reports_state = payload.get("reports")
        if not isinstance(reports_state, dict):
            return business_summary

        rows = business_summary.get("结案报告")
        if not isinstance(rows, list):
            rows = []
        row_map: dict[str, dict[str, Any]] = {}
        ordered_keys: list[str] = []
        for row in rows:
            if not isinstance(row, dict):
                continue
            norm = cls._normalize_report_name_key(row.get("报告名"))
            if not norm:
                continue
            row_map[norm] = row
            ordered_keys.append(norm)

        for norm, saved_row in reports_state.items():
            if not isinstance(saved_row, dict):
                continue
            if norm in row_map:
                row_map[norm] = cls._deep_merge_dict(saved_row, row_map[norm])
            else:
                row_map[norm] = saved_row
                ordered_keys.append(norm)

        merged_rows: list[dict[str, Any]] = []
        seen: set[str] = set()
        for norm in ordered_keys:
            if norm in seen:
                continue
            seen.add(norm)
            row = row_map.get(norm)
            if isinstance(row, dict):
                merged_rows.append(row)
        return {"结案报告": merged_rows}

    @classmethod
    def _persist_incremental_outputs(
        cls,
        *,
        output_dir: Path,
        profile,
        task_intent: YuntuTask | None,
        extracted_contents: list[Any],
        log_messages: list[str],
        current_report: str | None,
        step_no: int,
    ) -> None:
        report_targets = cls._build_report_targets(task_intent) if task_intent else []
        default_period_label = cls._infer_default_period_label_from_intent(task_intent)
        brand_name = (getattr(task_intent, "brand_name", None) if task_intent else None)
        fallback_report = current_report or (getattr(task_intent, "report_name", None) if task_intent else None)

        structured_output = cls._build_partial_structured_output(
            profile=profile,
            extracted_contents=extracted_contents,
            output_dir=output_dir,
            report_name=fallback_report,
            brand_name=brand_name,
            stop_reason=f"incremental_step_{step_no}",
            default_period_label=default_period_label,
        )
        structured_output = cls._ensure_report_targets_in_structured_output(structured_output, report_targets)
        hint_text = "\n".join(log_messages[-160:])
        structured_output = cls._apply_report_field_hints_to_structured_output(
            structured_output=structured_output,
            profile=profile,
            report_targets=report_targets,
            hint_text=hint_text,
        )
        report_period_map = cls._extract_report_period_map(extracted_contents)
        cls._sync_period_map_to_intent(task_intent, report_period_map)

        partial_structured_path = output_dir / "agent_structured_output.partial.json"
        partial_structured_path.write_text(structured_output.model_dump_json(indent=2), encoding="utf-8")

        business_summary = cls._build_business_summary_json(
            structured_output=structured_output,
            profile=profile,
            task_intent=task_intent,
            report_period_map=report_period_map,
            output_dir=output_dir,
        )
        business_summary = cls._merge_business_summary_with_progress(output_dir, business_summary)
        cls._update_report_progress_map(output_dir, business_summary)

        partial_business_path = output_dir / "business_summary.partial.json"
        partial_business_path.write_text(
            json.dumps(business_summary, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

        checkpoints_dir = output_dir / "checkpoints"
        checkpoints_dir.mkdir(parents=True, exist_ok=True)
        step_tag = f"step_{int(step_no):04d}"
        (checkpoints_dir / f"{step_tag}_structured.json").write_text(
            structured_output.model_dump_json(indent=2),
            encoding="utf-8",
        )
        (checkpoints_dir / f"{step_tag}_business_summary.json").write_text(
            json.dumps(business_summary, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

        report_progress_dir = output_dir / "report_progress"
        report_progress_dir.mkdir(parents=True, exist_ok=True)
        for row in business_summary.get("结案报告", []) or []:
            if not isinstance(row, dict):
                continue
            report_name = str(row.get("报告名") or "").strip()
            filename = cls._safe_filename_component(report_name) + ".json"
            (report_progress_dir / filename).write_text(
                json.dumps(row, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )

    @classmethod
    def _extract_report_field_hints_from_text(
        cls,
        text: str | None,
        report_targets: list[str],
        field_keys: list[str],
    ) -> dict[str, dict[str, tuple[str, str]]]:
        raw = str(text or "")
        if not raw:
            return {}
        chunks = [c.strip() for c in re.split(r"[;\n。]+", raw) if c and c.strip()]
        if not chunks:
            return {}

        target_by_norm: dict[str, str] = {}
        for target in report_targets:
            norm = cls._normalize_report_name_key(target)
            if norm:
                target_by_norm[norm] = target

        hints: dict[str, dict[str, tuple[str, str]]] = {}
        value_pattern = r"(-?\d[\d,]*(?:\.\d+)?\s*%?)"
        for chunk in chunks:
            chunk_norm = cls._normalize_report_name_key(chunk)
            if not chunk_norm:
                continue
            matched_report_norm: str | None = None
            for norm in target_by_norm:
                if norm and (norm in chunk_norm or chunk_norm in norm):
                    matched_report_norm = norm
                    break
            if not matched_report_norm:
                continue
            field_map = hints.setdefault(matched_report_norm, {})
            for field_key in field_keys:
                if field_key not in chunk:
                    continue
                m = re.search(rf"{re.escape(field_key)}\s*[:：]?\s*{value_pattern}", chunk)
                if not m:
                    continue
                token = str(m.group(1) or "").strip()
                if not token:
                    continue
                field_map[field_key] = (token, chunk)
        return hints

    @classmethod
    def _apply_report_field_hints_to_structured_output(
        cls,
        *,
        structured_output: YuntuAgentStructuredOutput,
        profile,
        report_targets: list[str],
        hint_text: str | None,
    ) -> YuntuAgentStructuredOutput:
        if not report_targets or not hint_text:
            return structured_output
        report_modules = [m for m in (getattr(profile, "module_specs", []) or []) if getattr(m, "requires_report", False)]
        if not report_modules:
            return structured_output

        report_field_keys: list[str] = []
        field_to_module_key: dict[str, str] = {}
        module_label_map: dict[str, str] = {}
        for module in report_modules:
            module_label_map[module.module_key] = module.label
            for key in getattr(module, "field_keys", []) or []:
                field_key = str(key).strip()
                if not field_key:
                    continue
                report_field_keys.append(field_key)
                field_to_module_key[field_key] = module.module_key
        report_field_keys = list(dict.fromkeys(report_field_keys))
        if not report_field_keys:
            return structured_output

        hints = cls._extract_report_field_hints_from_text(hint_text, report_targets, report_field_keys)
        if not hints:
            return structured_output
        try:
            debug_hints = {
                report_norm: {k: v[0] for k, v in field_map.items()}
                for report_norm, field_map in hints.items()
            }
            print("[report_field_hints]", json.dumps(debug_hints, ensure_ascii=False))
        except Exception:
            pass

        report_by_norm = {
            cls._normalize_report_name_key(r.report_name): r for r in (structured_output.reports or [])
        }
        rebuilt_reports: list[ReportExecutionResult] = []
        for report in structured_output.reports:
            report_norm = cls._normalize_report_name_key(report.report_name)
            report_hints = hints.get(report_norm, {})
            if not report_hints:
                rebuilt_reports.append(report)
                continue

            modules_map = {m.module_key: m for m in report.modules}
            for field_key, (token, chunk) in report_hints.items():
                module_key = field_to_module_key.get(field_key)
                if not module_key:
                    continue
                module = modules_map.get(module_key)
                if module is None:
                    module = ModuleExecutionResult(
                        module_key=module_key,
                        module_label=module_label_map.get(module_key),
                        status="partial",
                    )
                field_map = {f.field_key: f for f in module.fields}
                field_map[field_key] = FieldExtractionResult(
                    field_key=field_key,
                    status="ok",
                    source="auto",
                    value=cls._parse_token_value(token),
                    raw_text=chunk,
                    error=None,
                )
                module = module.model_copy(update={"fields": list(field_map.values())})
                modules_map[module_key] = module

            modules = list(modules_map.values())
            statuses = [m.status for m in modules]
            report_status = report.status
            if statuses and all(s == "success" for s in statuses):
                report_status = "success"
            elif any(s in {"success", "partial"} for s in statuses):
                report_status = "partial"
            rebuilt_reports.append(report.model_copy(update={"modules": modules, "status": report_status}))

        if not rebuilt_reports:
            return structured_output
        return structured_output.model_copy(update={"reports": rebuilt_reports})

    @classmethod
    def _extract_switched_brand(cls, extracted_contents: list[Any] | None, default_brand: str | None = None) -> str | None:
        picked: str | None = None
        for entry in extracted_contents or []:
            payload = cls._parse_json_payload(entry)
            if not payload:
                continue
            if str(payload.get("type") or "") != "switch_brand_result":
                continue
            status = str(payload.get("status") or "")
            if status not in {"ok", "partial", "success"}:
                continue
            picked_brand = str(payload.get("picked_brand") or "").strip()
            target_brand = str(payload.get("target_brand") or "").strip()
            picked = picked_brand or target_brand or picked
        return picked or default_brand

    @staticmethod
    def _scan_downloaded_files(output_dir: Path) -> list[Path]:
        if not output_dir.exists():
            return []
        ignored_suffixes = {
            ".json",
            ".txt",
            ".log",
            ".md",
            ".png",
            ".jpg",
            ".jpeg",
            ".gif",
            ".webp",
            ".html",
        }
        ignored_names = {"execution_logs.txt", "agent_output.json", "agent_structured_output.json"}
        files: list[Path] = []
        for path in output_dir.rglob("*"):
            if not path.is_file():
                continue
            if path.name in ignored_names:
                continue
            if any(part == "conversations" for part in path.parts):
                continue
            if path.suffix.lower() in ignored_suffixes:
                continue
            files.append(path)
        files.sort(key=lambda p: p.stat().st_mtime if p.exists() else 0.0, reverse=True)
        return files

    @staticmethod
    def _is_noisy_chunk(text: str) -> bool:
        t = str(text or "").strip()
        if not t:
            return True
        lower = t.lower()
        noisy_tokens = [
            "navigated to",
            "clicked",
            "waited for",
            "step ",
            "execution logs",
            "open_report_result",
            "click_in_content_result",
            "select_in_content_result",
            "file_download_result",
            "module_route_navigation_result",
            "open_yuntu_home_result",
            "navigate_report_list_result",
            "switch_brand_result",
            "open_report_result",
            "todo.md",
        ]
        if any(tok in lower for tok in noisy_tokens):
            return True
        if "http://" in lower or "https://" in lower:
            return True
        return False

    @staticmethod
    def _parse_token_value(token: str) -> float | int | str | None:
        text = str(token or "").strip()
        if not text:
            return None
        if "%" in text:
            return text.replace(" ", "")
        num = text.replace(",", "").strip()
        try:
            if "." in num:
                return float(num)
            return int(num)
        except Exception:
            return text

    @staticmethod
    def _norm_match_text(value: str | None) -> str:
        text = str(value or "").strip().lower()
        if not text:
            return ""
        text = re.sub(r"\s+", "", text)
        text = re.sub(r"[\u3002\u3001\uff0c\uff1a:,\._\-—–/\\|\(\)\[\]\{\}<>\"'`]", "", text)
        return text

    @classmethod
    def _field_tokens(cls, field_key: str) -> list[str]:
        text = str(field_key or "").strip()
        if not text:
            return []
        parts = re.split(r"[\s\u3002\u3001\uff0c\uff1a:,\._\-—–/\\|\(\)\[\]\{\}<>\"'`]+", text)
        tokens: list[str] = []
        for part in parts:
            p = cls._norm_match_text(part)
            if not p:
                continue
            if len(p) == 1 and not re.match(r"a[1-5]", p):
                continue
            tokens.append(p)
        out: list[str] = []
        seen: set[str] = set()
        for t in tokens:
            if t in seen:
                continue
            seen.add(t)
            out.append(t)
        return out

    @classmethod
    def _extract_field_rows_from_json_snippets(
        cls, chunks: list[str]
    ) -> dict[str, FieldExtractionResult]:
        out: dict[str, FieldExtractionResult] = {}
        pattern = re.compile(
            r'"field_key"\s*:\s*"(?P<key>[^"]+)"[\s\S]{0,400}?"value"\s*:\s*(?P<value>"[^"]*"|-?\d[\d,]*(?:\.\d+)?\s*%?)',
            flags=re.IGNORECASE,
        )
        for chunk in chunks:
            for m in pattern.finditer(chunk):
                key = str(m.group("key") or "").strip()
                raw_val = str(m.group("value") or "").strip()
                if not key or not raw_val:
                    continue
                if raw_val.startswith('"') and raw_val.endswith('"'):
                    raw_val = raw_val[1:-1]
                row = FieldExtractionResult(
                    field_key=key,
                    status="ok",
                    source="auto",
                    value=cls._parse_token_value(raw_val),
                    raw_text=f"{key} {raw_val}",
                    error=None,
                )
                prev = out.get(key)
                if prev is None or (prev.value is None and row.value is not None):
                    out[key] = row
        return out

    @classmethod
    def _recover_single_field_from_lines(
        cls,
        field_key: str,
        chunks: list[str],
    ) -> FieldExtractionResult | None:
        key = str(field_key or "").strip()
        if not key:
            return None
        tokens = cls._field_tokens(key)
        if not tokens:
            tokens = [cls._norm_match_text(key)]
        value_token_pattern = r"(-?\d[\d,]*(?:\.\d+)?\s*%?)"
        best_row: tuple[int, str, str] | None = None

        for chunk in chunks:
            if cls._is_noisy_chunk(chunk):
                continue
            for line in str(chunk).splitlines():
                txt = re.sub(r"\s+", " ", line).strip()
                if not txt or cls._is_noisy_chunk(txt):
                    continue
                line_norm = cls._norm_match_text(txt)
                if not line_norm:
                    continue
                overlap = [t for t in tokens if t and t in line_norm]
                if not overlap:
                    continue
                m = re.search(value_token_pattern, txt)
                if not m:
                    continue
                token = m.group(1).strip()
                score = 100 + len(overlap) * 20 + sum(min(len(t), 6) for t in overlap)
                if "%" in token and any(s in key for s in ("率", "占比", "比例", "环比")):
                    score += 20
                if best_row is None or score > best_row[0]:
                    best_row = (score, token, txt)

        if best_row is None:
            return None
        token, raw_line = best_row[1], best_row[2]
        return FieldExtractionResult(
            field_key=key,
            status="ok",
            source="auto",
            value=cls._parse_token_value(token),
            raw_text=raw_line,
            error=None,
        )

    @classmethod
    def _recover_single_field_from_chunks(
        cls,
        field_key: str,
        chunks: list[str],
    ) -> FieldExtractionResult | None:
        return cls._recover_single_field_from_lines(field_key, chunks)

    @classmethod
    def _recover_fields_from_raw_text(
        cls,
        profile,
        extracted_contents: list[Any] | None,
    ) -> dict[str, list[FieldExtractionResult]]:
        # Strict mode for interruption/failure recovery:
        # do NOT infer/guess fields from raw text; only keep tool-emitted structured items.
        return {}

    @classmethod
    def _build_partial_structured_output(
        cls,
        *,
        profile,
        extracted_contents: list[Any] | None,
        output_dir: Path,
        report_name: str | None,
        brand_name: str | None,
        stop_reason: str | None,
        default_period_label: str | None = None,
    ) -> YuntuAgentStructuredOutput:
        def _norm_module_status(raw_status: Any) -> str:
            text = str(raw_status or "").strip().lower()
            if text in {"success", "partial", "failed", "skipped"}:
                return text
            if text in {"ok"}:
                return "success"
            if text in {"error", "not_found", "timeout"}:
                return "failed"
            return "partial"

        tool_field_map, file_map = cls._collect_tool_results(extracted_contents)
        native_field_map, native_period_map = cls._collect_native_extract_results(
            profile,
            extracted_contents,
            default_period_label=default_period_label,
        )
        field_map: dict[str, dict[str, Any]] = dict(native_field_map)
        for key, raw in tool_field_map.items():
            field_map[key] = cls._prefer_field_result(field_map.get(key), raw)
        module_payload_map = cls._collect_module_results(extracted_contents)
        downloaded_files = cls._scan_downloaded_files(output_dir)

        module_results: dict[str, ModuleExecutionResult] = {}
        for module in profile.module_specs:
            payload = module_payload_map.get(module.module_key)
            if payload:
                raw_fields = payload.get("fields") or []
                raw_files = payload.get("files") or []
                raw_route = payload.get("route_trace") or []
                fields: list[FieldExtractionResult] = []
                files: list[FileDownloadResult] = []
                route_trace: list[RouteTraceItem] = []
                for row in raw_fields:
                    if isinstance(row, dict):
                        try:
                            fields.append(FieldExtractionResult.model_validate(row))
                        except Exception:
                            continue
                for row in raw_files:
                    if isinstance(row, dict):
                        try:
                            files.append(FileDownloadResult.model_validate(row))
                        except Exception:
                            continue
                for row in raw_route:
                    if isinstance(row, dict):
                        try:
                            route_trace.append(RouteTraceItem.model_validate(row))
                        except Exception:
                            continue
                period_results = cls._validate_period_results(payload.get("period_results") or {})
                if not period_results:
                    period_results = cls._validate_period_results(native_period_map.get(module.module_key, {}))
                if cls._module_supports_period(module):
                    period_results = cls._rename_default_period_label(period_results, default_period_label)
                else:
                    period_results = {}
                module_results[module.module_key] = ModuleExecutionResult(
                    module_key=module.module_key,
                    module_label=module.label,
                    status=_norm_module_status(payload.get("status")),
                    route_trace=route_trace,
                    fields=fields,
                    period_results=period_results,
                    files=files,
                    touchpoint=payload.get("touchpoint"),
                    note=payload.get("note"),
                )
                continue

            fields = []
            for key in getattr(module, "field_keys", []) or []:
                raw = field_map.get(str(key))
                if raw is not None:
                    try:
                        fields.append(FieldExtractionResult.model_validate(raw))
                    except Exception:
                        continue
            files = []
            for key in getattr(module, "file_keys", []) or []:
                raw = file_map.get(str(key))
                if raw is not None:
                    try:
                        files.append(FileDownloadResult.model_validate(raw))
                    except Exception:
                        continue
            period_results = cls._validate_period_results(native_period_map.get(module.module_key, {}))
            if cls._module_supports_period(module):
                period_results = cls._rename_default_period_label(period_results, default_period_label)
            else:
                period_results = {}
            if fields or files or period_results:
                status = "success" if (fields or files) and all(
                    (f.status == "ok" for f in fields)
                ) and all((f.status == "ok" for f in files)) else "partial"
            else:
                status = "skipped"
            module_results[module.module_key] = ModuleExecutionResult(
                module_key=module.module_key,
                module_label=module.label,
                status=status,
                fields=fields,
                period_results=period_results,
                files=files,
            )

        # Append filesystem-discovered downloads not captured by tool payloads.
        known_paths = {
            str((item.file_path or "")).strip()
            for module in module_results.values()
            for item in module.files
            if (item.file_path or "").strip()
        }
        recovered_file_rows: list[FileDownloadResult] = []
        for idx, path in enumerate(downloaded_files, 1):
            abs_path = str(path)
            if abs_path in known_paths:
                continue
            recovered_file_rows.append(
                FileDownloadResult(
                    file_key=f"recovered_file_{idx}",
                    status="ok",
                    file_path=abs_path,
                    file_name=path.name,
                    clicked_selector="fs_scan",
                    error=None,
                )
            )

        if recovered_file_rows:
            recovery_module_key = "__partial_recovery_downloads__"
            existing = module_results.get(recovery_module_key)
            if existing is None:
                module_results[recovery_module_key] = ModuleExecutionResult(
                    module_key=recovery_module_key,
                    module_label="Recovered Downloads",
                    status="partial",
                    files=recovered_file_rows,
                    note="Recovered from task directory scan after interruption/failure.",
                )
            else:
                module_results[recovery_module_key] = existing.model_copy(
                    update={"files": [*existing.files, *recovered_file_rows], "status": "partial"}
                )

        report_modules: list[ModuleExecutionResult] = []
        global_modules: list[ModuleExecutionResult] = []
        report_module_keys = {m.module_key for m in profile.module_specs if getattr(m, "requires_report", False)}
        for key, module_result in module_results.items():
            if key in report_module_keys:
                report_modules.append(module_result)
            else:
                global_modules.append(module_result)

        report_name_value = (report_name or "").strip() or "UNKNOWN_REPORT"
        report_status = "failed"
        if any(m.status in {"success", "partial"} for m in report_modules):
            report_status = "partial"
        if report_modules and all(m.status == "success" for m in report_modules):
            report_status = "success"

        reports: list[ReportExecutionResult] = []
        if report_modules:
            reports.append(
                ReportExecutionResult(
                    report_name=report_name_value,
                    report_type="single",
                    status=report_status,
                    modules=report_modules,
                    note="Partial output recovered after interruption/failure.",
                )
            )

        any_data = any(
            (module.fields or module.files or module.period_results)
            for module in [*report_modules, *global_modules]
        )
        task_status = "partial" if any_data else "failed"
        summary_bits = ["Partial output recovered."]
        if stop_reason:
            summary_bits.append(f"reason={stop_reason}")

        return YuntuAgentStructuredOutput(
            task_status=task_status,
            switched_brand=cls._extract_switched_brand(extracted_contents, default_brand=brand_name),
            menu_trace=[],
            reports=reports,
            global_modules=global_modules,
            summary=" ".join(summary_bits),
        )

    @staticmethod
    def _structured_output_has_data(structured_output: YuntuAgentStructuredOutput) -> bool:
        all_modules = [m for r in structured_output.reports for m in r.modules] + list(structured_output.global_modules)
        return any((m.fields or m.files or OptTaskPipeline._period_results_has_items(m.period_results)) for m in all_modules)

    @classmethod
    def _merge_structured_with_recovery(
        cls,
        primary: YuntuAgentStructuredOutput,
        recovery: YuntuAgentStructuredOutput,
    ) -> YuntuAgentStructuredOutput:
        def _merge_field_rows(
            current_rows: list[FieldExtractionResult],
            recovery_rows: list[FieldExtractionResult],
        ) -> list[FieldExtractionResult]:
            merged: dict[str, FieldExtractionResult] = {}
            for row in recovery_rows:
                key = str(getattr(row, "field_key", "") or "").strip()
                if not key:
                    continue
                merged[key] = row
            for row in current_rows:
                key = str(getattr(row, "field_key", "") or "").strip()
                if not key:
                    continue
                prev = merged.get(key)
                if prev is None:
                    merged[key] = row
                    continue
                if row.status == "ok" and prev.status != "ok":
                    merged[key] = row
                    continue
                if row.status == prev.status:
                    merged[key] = row
            return list(merged.values())

        def _merge_file_rows(
            current_rows: list[FileDownloadResult],
            recovery_rows: list[FileDownloadResult],
        ) -> list[FileDownloadResult]:
            merged: dict[str, FileDownloadResult] = {}
            for row in recovery_rows:
                key = str(getattr(row, "file_key", "") or "").strip()
                if not key:
                    continue
                merged[key] = row
            for row in current_rows:
                key = str(getattr(row, "file_key", "") or "").strip()
                if not key:
                    continue
                prev = merged.get(key)
                if prev is None:
                    merged[key] = row
                    continue
                if row.status == "ok" and prev.status != "ok":
                    merged[key] = row
                    continue
                if row.status == prev.status:
                    merged[key] = row
            return list(merged.values())

        def _merge_period_rows(
            current_period: dict[str, list[FieldExtractionResult]] | None,
            recovery_period: dict[str, list[FieldExtractionResult]] | None,
        ) -> dict[str, list[FieldExtractionResult]]:
            cur = cls._compact_period_results(current_period or {})
            rec = cls._compact_period_results(recovery_period or {})
            all_labels = sorted(set(cur.keys()) | set(rec.keys()))
            merged: dict[str, list[FieldExtractionResult]] = {}
            for label in all_labels:
                merged_rows = _merge_field_rows(cur.get(label, []), rec.get(label, []))
                if merged_rows:
                    merged[label] = merged_rows
            return merged

        def _merge_modules(primary_modules: list[ModuleExecutionResult], recovery_modules: list[ModuleExecutionResult]) -> list[ModuleExecutionResult]:
            merged: dict[str, ModuleExecutionResult] = {m.module_key: m for m in primary_modules}
            for rec in recovery_modules:
                cur = merged.get(rec.module_key)
                if cur is None:
                    merged[rec.module_key] = rec
                    continue
                merged_fields = _merge_field_rows(cur.fields, rec.fields)
                merged_files = _merge_file_rows(cur.files, rec.files)
                merged_period = _merge_period_rows(cur.period_results, rec.period_results)
                merged_status = cur.status
                if merged_status == "skipped" and (
                    merged_fields or merged_files or cls._period_results_has_items(merged_period)
                ):
                    merged_status = "partial"
                merged[rec.module_key] = cur.model_copy(
                    update={
                        "fields": merged_fields,
                        "files": merged_files,
                        "period_results": merged_period,
                        "status": merged_status,
                        "note": cur.note or rec.note,
                    }
                )
            return list(merged.values())

        primary_reports = list(primary.reports)
        if not primary_reports and recovery.reports:
            primary_reports = list(recovery.reports)
        elif primary_reports and recovery.reports:
            rec_report = recovery.reports[0]
            first = primary_reports[0]
            primary_reports[0] = first.model_copy(
                update={
                    "modules": _merge_modules(list(first.modules), list(rec_report.modules)),
                    "note": first.note or rec_report.note,
                }
            )

        merged_globals = _merge_modules(list(primary.global_modules), list(recovery.global_modules))
        merged = primary.model_copy(
            update={
                "reports": primary_reports,
                "global_modules": merged_globals,
                "switched_brand": primary.switched_brand or recovery.switched_brand,
                "summary": primary.summary or recovery.summary,
            }
        )
        if not cls._structured_output_has_data(merged):
            return recovery
        if merged.task_status == "failed" and recovery.task_status in {"partial", "success"}:
            return merged.model_copy(update={"task_status": recovery.task_status})
        return merged

    @classmethod
    def _reconcile_structured_output_with_tools(
        cls,
        structured_output: YuntuAgentStructuredOutput,
        profile,
        extracted_contents: list[Any] | None,
    ) -> YuntuAgentStructuredOutput:
        field_map, file_map = cls._collect_tool_results(extracted_contents)
        if not field_map and not file_map:
            return structured_output

        module_spec_map = {m.module_key: m for m in profile.module_specs}
        required_field_map = {f.field_key: bool(f.required) for f in profile.field_specs}
        required_file_map = {f.file_key: bool(f.required) for f in profile.file_specs}
        multi_report_mode = len(list(structured_output.reports or [])) > 1

        def _rebuild_module(module, *, apply_tool_overrides: bool):
            spec = module_spec_map.get(module.module_key)
            existing_fields = {item.field_key: item for item in module.fields}
            existing_files = {item.file_key: item for item in module.files}
            existing_period_results = cls._compact_period_results(getattr(module, "period_results", {}) or {})
            if spec is not None and not cls._module_supports_period(spec):
                existing_period_results = {}

            canonical_field_keys: list[str] = list(spec.field_keys) if spec else list(existing_fields.keys())
            canonical_file_keys: list[str] = list(spec.file_keys) if spec else list(existing_files.keys())

            merged_fields: list[FieldExtractionResult] = []
            seen_field_keys: set[str] = set()
            for key in canonical_field_keys:
                seen_field_keys.add(key)
                raw = field_map.get(key) if apply_tool_overrides else None
                if key in existing_fields:
                    merged_fields.append(existing_fields[key])
                elif raw is not None:
                    merged_fields.append(FieldExtractionResult.model_validate(raw))
            for key, item in existing_fields.items():
                if key not in seen_field_keys:
                    merged_fields.append(item)

            merged_files: list[FileDownloadResult] = []
            seen_file_keys: set[str] = set()
            for key in canonical_file_keys:
                seen_file_keys.add(key)
                raw = file_map.get(key) if apply_tool_overrides else None
                if key in existing_files:
                    merged_files.append(existing_files[key])
                elif raw is not None:
                    merged_files.append(FileDownloadResult.model_validate(raw))
            for key, item in existing_files.items():
                if key not in seen_file_keys:
                    merged_files.append(item)

            field_status = {item.field_key: item.status for item in merged_fields}
            file_status = {item.file_key: item.status for item in merged_files}
            required_field_keys = [k for k in canonical_field_keys if required_field_map.get(k, True)]
            required_file_keys = [k for k in canonical_file_keys if required_file_map.get(k, False)]
            all_required_fields_ok = all(field_status.get(k) == "ok" for k in required_field_keys)
            all_required_files_ok = all(file_status.get(k) == "ok" for k in required_file_keys)
            any_ok = any(item.status == "ok" for item in merged_fields) or any(item.status == "ok" for item in merged_files)
            if all_required_fields_ok and all_required_files_ok:
                merged_status = "success"
            elif any_ok:
                merged_status = "partial"
            else:
                merged_status = "failed" if (required_field_keys or required_file_keys) else module.status

            return module.model_copy(
                update={
                    "fields": merged_fields,
                    "files": merged_files,
                    "period_results": existing_period_results,
                    "status": merged_status,
                }
            )

        rebuilt_reports = []
        for report in structured_output.reports:
            # In multi-report mode, never apply global tool overrides to report modules,
            # otherwise same field_key values can leak across reports.
            rebuilt_modules = [
                _rebuild_module(module, apply_tool_overrides=not multi_report_mode)
                for module in report.modules
            ]
            report_status = report.status
            if rebuilt_modules:
                statuses = [m.status for m in rebuilt_modules]
                if all(s == "success" for s in statuses):
                    report_status = "success"
                elif any(s in {"success", "partial"} for s in statuses):
                    report_status = "partial"
                else:
                    report_status = "failed"
            rebuilt_reports.append(report.model_copy(update={"modules": rebuilt_modules, "status": report_status}))

        rebuilt_global_modules = [
            _rebuild_module(module, apply_tool_overrides=True) for module in structured_output.global_modules
        ]
        task_status = structured_output.task_status
        all_modules = [m for report in rebuilt_reports for m in report.modules] + rebuilt_global_modules
        if all_modules:
            statuses = [m.status for m in all_modules]
            if statuses and all(s == "success" for s in statuses):
                task_status = "success"
            elif any(s in {"success", "partial"} for s in statuses):
                task_status = "partial"
            else:
                task_status = "failed"

        return structured_output.model_copy(
            update={
                "reports": rebuilt_reports,
                "global_modules": rebuilt_global_modules,
                "task_status": task_status,
            }
        )

    async def _run_phase(
        self,
        *,
        agent: Agent,
        history: list,
        task_summary: str,
        phase_name: str,
        profile=None,
    ):
        log_messages: list[str] = []
        runtime_extracted_contents: list[Any] = []
        current_report_name: str | None = None
        last_log_count = 0
        last_stable_report_url: str | None = None
        report_phase_confirmed = False
        module_stage: int = -1
        empty_action_streak: int = 0
        hover_eval_seen_by_stage: set[int] = set()
        # Allow at most one retry for the same extract signature per stage.
        # This keeps flow stable while fixing first-pass false negatives.
        extract_signature_seen_by_stage: dict[int, dict[str, int]] = {}
        pending_click_target: str | None = None
        pending_click_retry_count: int = 0
        recent_success_click_targets: list[str] = []
        allowed_actions_in_report = {
            "click_in_content",
            "select_in_content",
            "evaluate",
            "extract",
            "download_files_by_spec",
            "sum_search_count_in_date_range",
            "wait",
            "done",
        }
        blocked_actions_in_report = {
            "navigate",
            "search",
            "input",
            "click",
            "go_back",
            "scroll",
            "find_text",
            "switch",
            "open_tab",
            "close_tab",
            "extract_fields_by_spec",
            "hover_in_content",
            "dropdown_options",
        }
        report_modules = [m for m in (getattr(profile, "module_specs", None) or []) if getattr(m, "requires_report", False)]
        all_modules = list(getattr(profile, "module_specs", None) or [])
        report_targets = self._build_report_targets(self.session.intent) if self.session.intent else []
        report_target_norm_map = {
            self._normalize_report_name_key(name): name for name in report_targets if self._normalize_report_name_key(name)
        }
        # If there is no report-dependent module, enforce guarded/tool-oriented mode from step 1
        # to avoid native-action drift (e.g. dropdown index probing).
        if not report_modules:
            report_phase_confirmed = True
        module_key_to_stage = {m.module_key: idx for idx, m in enumerate(report_modules)}
        module_route_tokens_by_key: dict[str, list[str]] = {}
        module_has_select_steps: dict[str, bool] = {}
        module_has_click_steps: dict[str, bool] = {}
        select_route_prereq_rules: list[dict[str, str]] = []
        click_route_prereq_rules: list[dict[str, str]] = []
        field_key_to_stage: dict[str, int] = {}
        file_key_to_stage: dict[str, int] = {}
        token_to_stages: dict[str, set[int]] = {}
        for module in all_modules:
            module_key = str(getattr(module, "module_key", "") or "").strip()
            route_tokens = [str(v).strip() for v in (getattr(module, "route_path", []) or []) if str(v).strip()]
            if module_key:
                module_route_tokens_by_key[module_key] = route_tokens
            has_select = False
            has_click = False
            route_tail = route_tokens[-1] if route_tokens else ""
            for step in (getattr(module, "interaction_steps", []) or []):
                op = str(getattr(step, "op", "") or "").strip()
                target = str(getattr(step, "target", "") or "").strip()
                if op in {"select", "select_path", "set_date_range"} and target and route_tail:
                    has_select = True
                    select_route_prereq_rules.append(
                        {
                            "module_key": module_key,
                            "target": target,
                            "route_tail": route_tail,
                        }
                    )
                if op in {"click", "ensure_visible"} and target and route_tail:
                    has_click = True
                    click_route_prereq_rules.append(
                        {
                            "module_key": module_key,
                            "target": target,
                            "route_tail": route_tail,
                        }
                    )
            if module_key:
                module_has_select_steps[module_key] = has_select
                module_has_click_steps[module_key] = has_click

        def _norm_token(value: str | None) -> str:
            text = str(value or "").strip().lower()
            if not text:
                return ""
            text = "".join(ch for ch in text if not ch.isspace())
            for ch in (".", ",", "_", "-", ":", ";", "/", "\\", "|", "(", ")", "[", "]", "{", "}", "。", "，", "：", "；"):
                text = text.replace(ch, "")
            return text

        def _register_token(stage: int, raw: str | None) -> None:
            token = _norm_token(raw)
            if not token:
                return
            token_to_stages.setdefault(token, set()).add(stage)

        def _resolve_token_stages(raw: str | None) -> set[int]:
            token = _norm_token(raw)
            if not token:
                return set()
            return set(token_to_stages.get(token, set()))

        for module in report_modules:
            stage = module_key_to_stage.get(module.module_key, -1)
            if stage >= 0:
                _register_token(stage, getattr(module, "module_key", None))
                _register_token(stage, getattr(module, "label", None))
            for key in getattr(module, "field_keys", []) or []:
                field_key_to_stage[str(key)] = max(stage, field_key_to_stage.get(str(key), -1))
                if stage >= 0:
                    _register_token(stage, str(key))
            for key in getattr(module, "file_keys", []) or []:
                file_key_to_stage[str(key)] = max(stage, file_key_to_stage.get(str(key), -1))
                if stage >= 0:
                    _register_token(stage, str(key))
            if stage >= 0:
                for target in getattr(module, "route_path", []) or []:
                    _register_token(stage, str(target))
                for step in getattr(module, "interaction_steps", []) or []:
                    _register_token(stage, str(getattr(step, "target", None) or ""))
                    _register_token(stage, str(getattr(step, "value", None) or ""))
                    for v in getattr(step, "path", []) or []:
                        _register_token(stage, str(v))

        overview_stage = module_key_to_stage.get("overview_post_report", 0)
        overview_click_tokens = {"活动总览", "流量分析", "转化分析"}
        for module in report_modules:
            if module_key_to_stage.get(module.module_key) != overview_stage:
                continue
            for step in getattr(module, "interaction_steps", []) or []:
                op = str(getattr(step, "op", "") or "")
                target = str(getattr(step, "target", "") or "").strip()
                if op in {"click", "ensure_visible"} and target:
                    overview_click_tokens.add(target)
            for target in getattr(module, "route_path", []) or []:
                t = str(target or "").strip()
                if t:
                    overview_click_tokens.add(t)

        def _update_stage_from_payload(payload_dict: dict[str, Any]) -> None:
            nonlocal module_stage
            payload_type = str(payload_dict.get("type") or "")
            if payload_type == "field_extraction_result":
                requested_keys = [str(k).strip() for k in (payload_dict.get("requested_field_keys") or []) if str(k).strip()]
                for key in requested_keys:
                    stage = field_key_to_stage.get(key, -1)
                    if stage >= 0:
                        module_stage = max(module_stage, stage)
                for item in payload_dict.get("items", []) or []:
                    if not isinstance(item, dict):
                        continue
                    key = str(item.get("field_key") or "").strip()
                    if not key:
                        continue
                    stage = field_key_to_stage.get(key, -1)
                    if stage >= 0:
                        module_stage = max(module_stage, stage)
            elif payload_type == "file_download_result":
                requested_file_keys = [str(k).strip() for k in (payload_dict.get("requested_file_keys") or []) if str(k).strip()]
                for key in requested_file_keys:
                    stage = file_key_to_stage.get(key, -1)
                    if stage >= 0:
                        module_stage = max(module_stage, stage)
                for item in payload_dict.get("items", []) or []:
                    if not isinstance(item, dict):
                        continue
                    key = str(item.get("file_key") or "").strip()
                    if not key:
                        continue
                    stage = file_key_to_stage.get(key, -1)
                    if stage >= 0:
                        module_stage = max(module_stage, stage)

        def _extract_action_name_and_params(action_obj) -> tuple[str, dict[str, Any]]:
            try:
                payload = action_obj.model_dump(exclude_none=True, mode="json")
            except Exception:
                payload = {}
            if isinstance(payload, dict):
                for key, value in payload.items():
                    if value is None:
                        continue
                    if isinstance(value, dict):
                        return str(key), value
                    return str(key), {}
            name = str(getattr(action_obj, "action_name", "") or "").strip()
            return name, {}

        def _max_stage_from_text_values(values: list[str | None]) -> int:
            stages: set[int] = set()
            for value in values:
                stages |= _resolve_token_stages(value)
            return max(stages) if stages else -1

        def _remember_success_click_target(target: str | None) -> None:
            text = str(target or "").strip()
            if not text:
                return
            recent_success_click_targets.append(text)
            deduped = list(dict.fromkeys([t for t in recent_success_click_targets if str(t).strip()]))
            recent_success_click_targets[:] = deduped[-12:]

        def _next_missing_route_token(route_tokens: list[str]) -> tuple[str | None, int]:
            if not route_tokens:
                return None, 0
            idx = 0
            for clicked in recent_success_click_targets:
                if idx >= len(route_tokens):
                    break
                if _same_target(clicked, route_tokens[idx]):
                    idx += 1
            if idx >= len(route_tokens):
                return None, idx
            return route_tokens[idx], idx

        def _pick_route_tail_for_select_target(select_target: str | None) -> str | None:
            target = str(select_target or "").strip()

            scored_candidates: list[tuple[int, str]] = []
            explicit_match_seen = False

            # 1) Prefer explicit select-target matched modules.
            if target:
                matched_module_keys: set[str] = set()
                for rule in select_route_prereq_rules:
                    if _same_target(target, rule.get("target")):
                        mk = str(rule.get("module_key") or "").strip()
                        if mk:
                            matched_module_keys.add(mk)
                if matched_module_keys:
                    explicit_match_seen = True
                    any_matched_route_complete = False
                for mk in matched_module_keys:
                    route_tokens = module_route_tokens_by_key.get(mk, [])
                    next_token, progress = _next_missing_route_token(route_tokens)
                    if next_token is None and progress >= len(route_tokens):
                        any_matched_route_complete = True
                        continue
                    if next_token:
                        scored_candidates.append((100 + progress, next_token))
                # Critical: when target maps to multiple modules (e.g. "日期类型"),
                # if any matched module route is already complete, do not force-jump
                # to another matched module's tail.
                if any_matched_route_complete:
                    return None

            # 2) Fallback by route progress (handles wrong target text like "对比周期").
            if not scored_candidates and not explicit_match_seen:
                for mk, route_tokens in module_route_tokens_by_key.items():
                    if not module_has_select_steps.get(mk, False):
                        continue
                    next_token, progress = _next_missing_route_token(route_tokens)
                    if not next_token:
                        continue
                    # Require at least one matched prefix to avoid unrelated jumps.
                    if progress <= 0:
                        continue
                    scored_candidates.append((progress, next_token))

            if not scored_candidates:
                return None
            scored_candidates.sort(key=lambda x: x[0], reverse=True)
            return scored_candidates[0][1]

        def _pick_route_tail_for_click_target(click_target: str | None) -> str | None:
            target = str(click_target or "").strip()
            if not target:
                return None
            matched_module_keys: set[str] = set()
            for rule in click_route_prereq_rules:
                if _same_target(target, rule.get("target")):
                    mk = str(rule.get("module_key") or "").strip()
                    if mk:
                        matched_module_keys.add(mk)
            scored_candidates: list[tuple[int, str]] = []
            if matched_module_keys:
                for mk in matched_module_keys:
                    route_tokens = module_route_tokens_by_key.get(mk, [])
                    next_token, progress = _next_missing_route_token(route_tokens)
                    if next_token is None and progress >= len(route_tokens):
                        continue
                    if next_token:
                        scored_candidates.append((100 + progress, next_token))
            else:
                # Fallback when click text matching is unreliable:
                # choose next missing route token from the module that has click steps
                # and currently has the deepest matched route prefix.
                for mk, route_tokens in module_route_tokens_by_key.items():
                    if not module_has_click_steps.get(mk, False):
                        continue
                    next_token, progress = _next_missing_route_token(route_tokens)
                    if not next_token:
                        continue
                    if progress <= 0:
                        continue
                    scored_candidates.append((progress, next_token))
            if not scored_candidates:
                return None
            scored_candidates.sort(key=lambda x: x[0], reverse=True)
            return scored_candidates[0][1]

        def _pick_recent_route_tail_for_failed_select(select_target: str | None) -> str | None:
            target = str(select_target or "").strip()
            matched_tails: list[str] = []
            if target:
                for rule in select_route_prereq_rules:
                    if _same_target(target, rule.get("target")):
                        rt = str(rule.get("route_tail") or "").strip()
                        if rt:
                            matched_tails.append(rt)

            # Prefer the most recently successful matched tail.
            if matched_tails:
                for clicked in reversed(recent_success_click_targets):
                    for tail in matched_tails:
                        if _same_target(clicked, tail):
                            return tail
                return matched_tails[0]

            # Fallback: most recent clicked token that is any module route tail with select steps.
            all_tails: list[str] = []
            for mk, route_tokens in module_route_tokens_by_key.items():
                if not module_has_select_steps.get(mk, False):
                    continue
                if route_tokens:
                    all_tails.append(route_tokens[-1])
            for clicked in reversed(recent_success_click_targets):
                for tail in all_tails:
                    if _same_target(clicked, tail):
                        return tail
            return None

        def _current_stage_key() -> int:
            return module_stage

        def _is_report_stage(stage_key: int) -> bool:
            return 0 <= stage_key < len(report_modules)

        def _norm_soft_text(value: str | None) -> str:
            text = str(value or "").strip().lower()
            if not text:
                return ""
            text = re.sub(r"\s+", "", text)
            text = re.sub(r"[\.銆?锛?_\-\(\)\[\]锛堬級銆愩€憒]+", "", text)
            return text

        def _same_target(a: str | None, b: str | None) -> bool:
            sa = _norm_soft_text(a)
            sb = _norm_soft_text(b)
            if not sa or not sb:
                return False
            return sa == sb or sa in sb or sb in sa

        def _normalize_extract_query(raw: str | None) -> str:
            text = str(raw or "").strip().lower()
            if not text:
                return ""
            text = re.sub(r"\s+", " ", text)
            return text

        def _extract_signature_for_stage(query: str | None, stage_key: int) -> str:
            normalized_query = _normalize_extract_query(query)
            if not normalized_query:
                return ""
            if stage_key < 0 or stage_key >= len(report_modules):
                return f"query:{normalized_query}"

            module = report_modules[stage_key]
            field_tokens: list[str] = []
            normalized_no_space = re.sub(r"\s+", "", normalized_query)
            for key in getattr(module, "field_keys", []) or []:
                raw_key = str(key or "").strip()
                if not raw_key:
                    continue
                norm_key = re.sub(r"\s+", "", raw_key.lower())
                if norm_key and norm_key in normalized_no_space:
                    field_tokens.append(raw_key)

            if field_tokens:
                dedup_sorted = sorted(set(field_tokens))
                return f"fields:{'|'.join(dedup_sorted)}"
            return f"query:{normalized_query}"

        def _is_hover_like_evaluate(code: str) -> bool:
            text = str(code or "").lower()
            if not text:
                return False
            mouse_like = any(token in text for token in ("mouseover", "mouseenter", "mousemove", "mouseevent"))
            selector_like = any(token in text for token in ("queryselector", "elementsfrompoint", "dispatchEvent".lower()))
            hover_words = any(token in text for token in ("hover", "tooltip", "popover", "question", "help"))
            return (mouse_like and selector_like) or (mouse_like and hover_words)

        def _evaluate_hover_guard_reason(code: str) -> str | None:
            normalized = re.sub(r"\s+", "", str(code or "").lower())
            if not normalized:
                return "empty_evaluate_code"
            banned_patterns = [
                "queryselector('svg')",
                "queryselector(\"svg\")",
                "queryselectorall('svg')",
                "queryselectorall(\"svg\")",
            ]
            for pattern in banned_patterns:
                if pattern in normalized:
                    return f"unsafe_selector:{pattern}"
            required_evidence_tokens = ("hover_ok", "tooltip_visible", "tooltip_text")
            if any(token not in normalized for token in required_evidence_tokens):
                return "missing_hover_evidence_fields"
            return None

        async def _inspect_page(agent_instance) -> dict:
            try:
                browser_session = getattr(agent_instance, "browser_session", None)
                if browser_session is None:
                    return {"ok": False, "error": "browser_session_missing"}
                page = await browser_session.must_get_current_page()
                js = r"""
() => {
  const clean = (v) => (v || "").replace(/\s+/g, " ").trim();
  const url = String(location.href || "");
  const title = clean(document.title || "");
  const body = clean((document.body?.innerText || "").slice(0, 2200));
  const merged = `${title} ${body}`.toLowerCase();
  const lowerUrl = url.toLowerCase();
  const isMessageRoute =
    /\/(message|notice|notification)(\/|$)/.test(lowerUrl) ||
    /#\/?(message|notice|notification)/.test(lowerUrl) ||
    /messagecenter|msg-?center/.test(lowerUrl);
  const hasMessageLabel = /消息中心|消息通知|站内信|通知中心/.test(`${title} ${body}`);
  const isReportLike =
    /结案报告|投后结案|报告列表|流量分析|人群分析|转化分析|画像分析|5a|营销决策/.test(merged) ||
    /\/(report|post|marketing|decision|insight)/.test(lowerUrl);
  const isUnexpected = !!isMessageRoute || (!!hasMessageLabel && !isReportLike);
  return { url, title, is_report_like: isReportLike, is_unexpected: isUnexpected };
}
"""
                raw = await page.evaluate(js)
                if isinstance(raw, dict):
                    return {"ok": True, **raw}
                return {"ok": False, "error": "invalid_js_result"}
            except Exception as exc:  # noqa: BLE001
                return {"ok": False, "error": str(exc)}

        async def _recover_if_unexpected(agent_instance) -> dict:
            nonlocal last_stable_report_url
            inspected = await _inspect_page(agent_instance)
            if not inspected.get("ok"):
                return {"recovered": False, "checked": inspected}

            if inspected.get("is_report_like") and inspected.get("url"):
                last_stable_report_url = str(inspected["url"])

            if not inspected.get("is_unexpected"):
                return {"recovered": False, "checked": inspected}

            browser_session = getattr(agent_instance, "browser_session", None)
            if browser_session is None:
                return {"recovered": False, "checked": inspected, "error": "browser_session_missing"}
            page = await browser_session.must_get_current_page()

            attempts: list[dict] = []
            for idx in range(2):
                back_ok = True
                try:
                    await page.go_back(wait_until="domcontentloaded")
                except Exception:
                    try:
                        await page.go_back()
                    except Exception as exc:  # noqa: BLE001
                        back_ok = False
                        attempts.append({"attempt": idx + 1, "op": "go_back", "ok": False, "error": str(exc)})
                await asyncio.sleep(0.35)
                checked = await _inspect_page(agent_instance)
                attempts.append({"attempt": idx + 1, "op": "go_back", "ok": back_ok, "checked": checked})
                if checked.get("ok") and not checked.get("is_unexpected"):
                    if checked.get("is_report_like") and checked.get("url"):
                        last_stable_report_url = str(checked["url"])
                    return {"recovered": True, "method": "go_back", "before": inspected, "after": checked, "attempts": attempts}

            if last_stable_report_url:
                try:
                    page = await browser_session.must_get_current_page()
                    await page.goto(last_stable_report_url)
                    await asyncio.sleep(0.5)
                    checked = await _inspect_page(agent_instance)
                    attempts.append({"op": "goto_last_stable", "url": last_stable_report_url, "checked": checked})
                    if checked.get("ok") and not checked.get("is_unexpected"):
                        return {
                            "recovered": True,
                            "method": "goto_last_stable",
                            "before": inspected,
                            "after": checked,
                            "attempts": attempts,
                        }
                except Exception as exc:  # noqa: BLE001
                    attempts.append({"op": "goto_last_stable", "url": last_stable_report_url, "error": str(exc)})

            return {"recovered": False, "before": inspected, "attempts": attempts}

        def _looks_like_report_detail_from_state(state_summary) -> bool:
            url = str(getattr(state_summary, "url", "") or "").lower()
            title = str(getattr(state_summary, "title", "") or "").lower()
            text = f"{url} {title}"
            analysis_tokens = [
                "流量分析",
                "人群分析",
                "转化分析",
                "画像分析",
                "5A关系资产分析",
            ]
            if any(tok in text for tok in [t.lower() for t in analysis_tokens]):
                return True
            if "结案报告" in text and not any(tok in text for tok in [t.lower() for t in analysis_tokens]):
                return False
            if any(seg in url for seg in ["/report/detail", "/report/view", "/post/report", "/insight/report"]):
                return True
            return False

        async def _pre_action_guard(browser_state_summary, model_output, step_no: int):
            try:
                if not report_phase_confirmed and not _looks_like_report_detail_from_state(browser_state_summary):
                    return
                actions = list(getattr(model_output, "action", []) or [])
                if not actions:
                    return

                # Global-only bootstrap escape hatch:
                # when page is blank/error, do not enforce report-phase action filtering,
                # otherwise model cannot perform initial navigate/open and gets stuck in wait loops.
                if not report_modules:
                    cur_url = str(getattr(browser_state_summary, "url", "") or "").strip().lower()
                    is_blank_or_error = (
                        (not cur_url)
                        or cur_url == "about:blank"
                        or cur_url.startswith("chrome-error://")
                        or cur_url.startswith("edge-error://")
                    )
                    if is_blank_or_error:
                        return

                filtered = []
                blocked = []
                forced_route_click_target: str | None = None
                for action in actions:
                    name, params = _extract_action_name_and_params(action)
                    if not name:
                        blocked.append("unknown_action")
                        continue

                    # Route confirmation guard:
                    # if a required click target failed in previous step, only allow retrying the same click target.
                    if pending_click_target:
                        if name != "click_in_content":
                            blocked.append(
                                f"{name}:pending_click_target:{pending_click_target}"
                            )
                            continue
                        current_target = str((params or {}).get("target_text") or "").strip()
                        if not _same_target(current_target, pending_click_target):
                            blocked.append(
                                f"{name}:wrong_target:{current_target or '<empty>'}:pending={pending_click_target}"
                            )
                            continue

                    # Route-path prerequisite guard for selects:
                    # if select target belongs to a module whose route tail is not yet reached,
                    # inject one click on that route tail before running select.
                    if name == "select_in_content":
                        select_target = str((params or {}).get("target_text") or "").strip()
                        route_tail = _pick_route_tail_for_select_target(select_target)
                        if route_tail and not any(_same_target(route_tail, t) for t in recent_success_click_targets):
                            blocked.append(f"{name}:route_tail_not_reached:{route_tail}")
                            if forced_route_click_target is None:
                                forced_route_click_target = route_tail
                            continue

                    if name == "evaluate":
                        code = str((params or {}).get("code") or "")
                        if _is_hover_like_evaluate(code):
                            stage_key = _current_stage_key()
                            if _is_report_stage(stage_key) and stage_key in hover_eval_seen_by_stage:
                                blocked.append(f"{name}:single_pass_hover_already_executed:stage={stage_key}")
                                continue
                            reason = _evaluate_hover_guard_reason(code)
                            if reason:
                                blocked.append(f"{name}:{reason}")
                                continue

                    if name == "extract":
                        stage_key = _current_stage_key()
                        signature = _extract_signature_for_stage((params or {}).get("query"), stage_key)
                        if signature and _is_report_stage(stage_key):
                            seen_map = extract_signature_seen_by_stage.get(stage_key, {})
                            attempt_count = int(seen_map.get(signature, 0))
                            if attempt_count >= 2:
                                blocked.append(f"{name}:single_pass_extract_already_executed:{signature}")
                                continue

                    # Stage lock: once agent has moved to later modules, don't allow fallback to overview module.
                    if module_stage > overview_stage:
                        if name in {"click_in_content", "hover_in_content"}:
                            target = str((params or {}).get("target_text") or "").strip()
                            if target and target in overview_click_tokens:
                                blocked.append(f"{name}:{target}:stage_lock")
                                continue
                        if name == "extract_fields_by_spec":
                            field_keys = [str(k).strip() for k in ((params or {}).get("field_keys") or []) if str(k).strip()]
                            if field_keys:
                                key_stages = [field_key_to_stage.get(key, module_stage) for key in field_keys]
                                if key_stages and max(key_stages) <= overview_stage:
                                    blocked.append(f"{name}:{','.join(field_keys)}:stage_lock")
                                    continue
                        if name == "download_files_by_spec":
                            file_keys = [str(k).strip() for k in ((params or {}).get("file_keys") or []) if str(k).strip()]
                            if file_keys:
                                key_stages = [file_key_to_stage.get(key, -1) for key in file_keys]
                                # Block downloading files that belong to future modules.
                                if any(stage > module_stage for stage in key_stages if stage >= 0):
                                    blocked.append(f"{name}:{','.join(file_keys)}:stage_lock")
                                    continue

                    # Generic no-backtracking guard:
                    # once module stage has advanced, don't execute actions that map only to earlier stages.
                    if module_stage >= 0:
                        if name in {"click_in_content", "hover_in_content"}:
                            target = str((params or {}).get("target_text") or "").strip()
                            target_stage = _max_stage_from_text_values([target])
                            if 0 <= target_stage < module_stage:
                                blocked.append(f"{name}:{target}:backtrack")
                                continue
                        elif name == "select_in_content":
                            target = str((params or {}).get("target_text") or "").strip()
                            value = str((params or {}).get("value") or "").strip()
                            path_values = [str(v).strip() for v in ((params or {}).get("path") or []) if str(v).strip()]
                            select_stage = _max_stage_from_text_values([target, value, *path_values])
                            if 0 <= select_stage < module_stage:
                                blocked.append(f"{name}:{target or value}:backtrack")
                                continue
                        elif name == "extract_fields_by_spec":
                            field_keys = [str(k).strip() for k in ((params or {}).get("field_keys") or []) if str(k).strip()]
                            if field_keys:
                                key_stages = [field_key_to_stage.get(key, -1) for key in field_keys]
                                known = [s for s in key_stages if s >= 0]
                                if known and max(known) < module_stage:
                                    blocked.append(f"{name}:{','.join(field_keys)}:backtrack")
                                    continue
                        elif name == "download_files_by_spec":
                            file_keys = [str(k).strip() for k in ((params or {}).get("file_keys") or []) if str(k).strip()]
                            if file_keys:
                                key_stages = [file_key_to_stage.get(key, -1) for key in file_keys]
                                known = [s for s in key_stages if s >= 0]
                                if known and max(known) < module_stage:
                                    blocked.append(f"{name}:{','.join(file_keys)}:backtrack")
                                    continue

                    if name in allowed_actions_in_report:
                        filtered.append(action)
                        continue
                    if name in blocked_actions_in_report:
                        blocked.append(name)
                        continue
                    # Unknown actions are blocked in report phase to avoid drift.
                    blocked.append(name)

                if blocked:
                    log_messages.append(
                        f"  pre_guard: blocked raw actions in report phase at step {step_no}: {', '.join(blocked)}"
                    )

                if forced_route_click_target:
                    try:
                        forced_route_selectors = [
                            "main [role='tab']",
                            "main .ant-tabs-tab",
                            "main [role='menuitem']",
                            "main .ant-menu-item",
                            "main a",
                            "main button",
                            ".ant-layout-content [role='tab']",
                            ".ant-layout-content .ant-tabs-tab",
                            ".ant-layout-content [role='menuitem']",
                            ".ant-layout-content .ant-menu-item",
                            ".ant-layout-content a",
                            ".ant-layout-content button",
                            "span",
                        ]
                        output_payload = model_output.model_dump(mode="json", exclude_none=False)
                        output_payload["action"] = [
                            {
                                "click_in_content": {
                                    "profile_name": str(getattr(profile, "profile_name", "generic") or "generic"),
                                    "target_text": forced_route_click_target,
                                    "selectors": forced_route_selectors,
                                    "wait_seconds": 2.0,
                                }
                            }
                        ]
                        rebuilt = type(model_output).model_validate(output_payload)
                        rebuilt_actions = list(getattr(rebuilt, "action", []) or [])
                        if rebuilt_actions:
                            model_output.action = rebuilt_actions
                            log_messages.append(
                                f"  pre_guard: enforce_route_tail_before_select -> click_in_content('{forced_route_click_target}')"
                            )
                            return
                    except Exception as exc:
                        log_messages.append(f"  pre_guard: enforce_route_tail_injection_failed: {exc}")

                if filtered:
                    model_output.action = filtered
                    return

                # No safe action left: inject a short wait to trigger replanning.
                safe_action = None
                try:
                    output_payload = model_output.model_dump(mode="json", exclude_none=False)
                    # Use integer seconds for wait action model compatibility.
                    output_payload["action"] = [{"wait": {"seconds": 1}}]
                    rebuilt = type(model_output).model_validate(output_payload)
                    rebuilt_actions = list(getattr(rebuilt, "action", []) or [])
                    if rebuilt_actions:
                        safe_action = rebuilt_actions[0]
                except Exception:
                    safe_action = None
                if safe_action is not None:
                    model_output.action = [safe_action]
                    log_messages.append(
                        f"  pre_guard: replaced blocked actions with wait at step {step_no} (report phase)"
                    )
                else:
                    # Global-only runs: keep one fallback action, but avoid unsafe raw click
                    # when pending click/date guards are active.
                    if not report_modules and actions:
                        first_name, _first_params = _extract_action_name_and_params(actions[0])
                        has_pending_block = any("pending_click_target:" in str(item) for item in blocked)
                        has_date_block = any("date_range_native_only" in str(item) for item in blocked)
                        if (has_pending_block or has_date_block) and first_name == "click":
                            model_output.action = []
                            log_messages.append(
                                f"  pre_guard: wait-injection failed at step {step_no}; blocked raw click fallback due to pending/date guard (global-only)"
                            )
                        else:
                            model_output.action = actions[:1]
                            log_messages.append(
                                f"  pre_guard: wait-injection failed at step {step_no}; fallback to original action (global-only)"
                            )
                    else:
                        # Report-phase: keep strict block to prevent drift.
                        model_output.action = []
                        log_messages.append(
                            f"  pre_guard: wait-injection failed at step {step_no}; blocked original unsafe action"
                        )
            except Exception as exc:  # noqa: BLE001
                log_messages.append(f"  pre_guard_error: {exc}")

        async def _pre_action_pending_click_retry_guard(browser_state_summary, model_output, step_no: int):
            nonlocal pending_click_target, pending_click_retry_count
            try:
                actions = list(getattr(model_output, "action", []) or [])
                route_click_selectors = [
                    "main [role='tab']",
                    "main .ant-tabs-tab",
                    "main [role='menuitem']",
                    "main .ant-menu-item",
                    "main a",
                    "main button",
                    ".ant-layout-content [role='tab']",
                    ".ant-layout-content .ant-tabs-tab",
                    ".ant-layout-content [role='menuitem']",
                    ".ant-layout-content .ant-menu-item",
                    ".ant-layout-content a",
                    ".ant-layout-content button",
                ]

                def _looks_like_date_range_text(value: str | None) -> bool:
                    text = str(value or "").strip()
                    if not text:
                        return False
                    # Examples:
                    # 2026.01.23 ~ 2026.02.14
                    # 2026-01-23 ~ 2026-02-14
                    # 2026/01/23-2026/02/14
                    return bool(
                        re.search(
                            r"\d{4}[./-]\d{1,2}[./-]\d{1,2}\s*[~\-至到]+\s*\d{4}[./-]\d{1,2}[./-]\d{1,2}",
                            text,
                        )
                    )

                # Always-on date typing guard:
                # forbid manual typing for date range; force opening picker by clicking the same control.
                if actions:
                    first_name, first_params = _extract_action_name_and_params(actions[0])
                    if first_name == "input":
                        typed_text = str((first_params or {}).get("text") or "").strip()
                        if _looks_like_date_range_text(typed_text):
                            idx = (first_params or {}).get("index")
                            try:
                                output_payload = model_output.model_dump(mode="json", exclude_none=False)
                                if isinstance(idx, int):
                                    output_payload["action"] = [{"click": {"index": int(idx)}}]
                                else:
                                    output_payload["action"] = [{"wait": {"seconds": 1}}]
                                rebuilt = type(model_output).model_validate(output_payload)
                                rebuilt_actions = list(getattr(rebuilt, "action", []) or [])
                                if rebuilt_actions:
                                    model_output.action = rebuilt_actions
                                    log_messages.append(
                                        "  date_guard: blocked manual date typing; replaced with picker-open action"
                                    )
                                    return
                            except Exception as exc:
                                log_messages.append(f"  date_guard_injection_failed: {exc}")

                # Always-on minimal route-tail guard:
                # pre_guard is usually disabled, so enforce route completion for select steps here.
                enable_route_tail_guard = str(os.getenv("YUNTU_ENABLE_ROUTE_TAIL_GUARD", "1")).strip().lower() in {
                    "1",
                    "true",
                    "yes",
                    "on",
                }
                if enable_route_tail_guard and not pending_click_target and actions:
                    first_name, first_params = _extract_action_name_and_params(actions[0])
                    if first_name == "select_in_content":
                        select_target = str((first_params or {}).get("target_text") or "").strip()
                        route_tail = _pick_route_tail_for_select_target(select_target)
                        if route_tail and not any(_same_target(route_tail, t) for t in recent_success_click_targets):
                            try:
                                output_payload = model_output.model_dump(mode="json", exclude_none=False)
                                output_payload["action"] = [
                                    {
                                        "click_in_content": {
                                            "profile_name": str(getattr(profile, "profile_name", "generic") or "generic"),
                                            "target_text": route_tail,
                                            "selectors": route_click_selectors,
                                            "wait_seconds": 2.0,
                                        }
                                    }
                                ]
                                rebuilt = type(model_output).model_validate(output_payload)
                                rebuilt_actions = list(getattr(rebuilt, "action", []) or [])
                                if rebuilt_actions:
                                    model_output.action = rebuilt_actions
                                    pending_click_target = route_tail
                                    pending_click_retry_count = 0
                                    log_messages.append(
                                        f"  click_retry_guard: enforce_route_tail_before_select -> click_in_content('{route_tail}')"
                                    )
                                    return
                            except Exception as exc:
                                log_messages.append(f"  click_retry_guard: enforce_route_tail_injection_failed: {exc}")
                    # Intentionally do not force route-tail injection for click steps.
                    # Click-step routing is already encoded in DSL order; forcing route tail
                    # here can cause ping-pong loops between route tail and next click target.

                if not pending_click_target:
                    return

                # Do not hard-lock forever on one missing node.
                if pending_click_retry_count >= 2:
                    log_messages.append(
                        f"  click_retry_guard: retry limit reached for '{pending_click_target}', continue next steps"
                    )
                    pending_click_target = None
                    pending_click_retry_count = 0
                    return

                if actions:
                    name, params = _extract_action_name_and_params(actions[0])
                    current_target = str((params or {}).get("target_text") or "").strip()
                    if name == "click_in_content" and _same_target(current_target, pending_click_target):
                        return

                output_payload = model_output.model_dump(mode="json", exclude_none=False)
                output_payload["action"] = [
                    {
                        "click_in_content": {
                            "profile_name": str(getattr(profile, "profile_name", "generic") or "generic"),
                            "target_text": pending_click_target,
                            "selectors": route_click_selectors,
                            "wait_seconds": 2.0,
                        }
                    }
                ]
                rebuilt = type(model_output).model_validate(output_payload)
                rebuilt_actions = list(getattr(rebuilt, "action", []) or [])
                if rebuilt_actions:
                    model_output.action = rebuilt_actions
                    log_messages.append(
                        f"  click_retry_guard: enforce retry click_in_content('{pending_click_target}')"
                    )
            except Exception as exc:  # noqa: BLE001
                log_messages.append(f"  click_retry_guard_error: {exc}")

        async def on_step_start(agent_instance):
            log_messages.append(f"{phase_name} step {agent_instance.state.n_steps} started")

        async def on_step_end(agent_instance):
            nonlocal report_phase_confirmed, empty_action_streak, pending_click_target, pending_click_retry_count, current_report_name
            step_num = agent_instance.state.n_steps
            output = agent_instance.state.last_model_output

            def _detect_report_from_text(text: str | None) -> str | None:
                value = str(text or "").strip()
                if not value:
                    return None
                norm_text = self._normalize_report_name_key(value)
                if not norm_text:
                    return None
                for norm_target, report_name in report_target_norm_map.items():
                    if norm_target and (norm_target in norm_text or norm_text in norm_target):
                        return report_name
                return None

            # Capture active report from model narrative when available.
            if output is not None:
                for attr in ("memory", "next_goal", "evaluation_previous_goal", "thinking"):
                    detected = _detect_report_from_text(getattr(output, attr, None))
                    if detected:
                        current_report_name = detected
                        break

            should_checkpoint = False
            if output and output.action:
                empty_action_streak = 0
                for action in output.action:
                    action_name, action_params = _extract_action_name_and_params(action)
                    if not action_name:
                        action_name = str(type(action).__name__)
                    log_messages.append(f"  action: {action_name}")

                    stage_key = _current_stage_key()
                    if action_name == "evaluate":
                        code = str((action_params or {}).get("code") or "")
                        if _is_hover_like_evaluate(code):
                            if _is_report_stage(stage_key):
                                hover_eval_seen_by_stage.add(stage_key)
                    elif action_name in {
                        "extract",
                        "click_in_content",
                        "select_in_content",
                        "download_files_by_spec",
                        "done",
                    }:
                        # Single-pass mode: do not reset hover execution flags.
                        pass

                    if action_name == "extract":
                        signature = _extract_signature_for_stage((action_params or {}).get("query"), stage_key)
                        if signature and _is_report_stage(stage_key):
                            stage_map = extract_signature_seen_by_stage.setdefault(stage_key, {})
                            stage_map[signature] = int(stage_map.get(signature, 0)) + 1

                    if action_name in {
                        "click_in_content",
                        "hover_in_content",
                        "select_in_content",
                        "extract",
                        "extract_fields_by_spec",
                        "download_files_by_spec",
                    }:
                        report_phase_confirmed = True
            elif output is not None:
                empty_action_streak += 1
                log_messages.append(
                    f"  guard: empty_action_detected streak={empty_action_streak}"
                )
                # Prevent no-op infinite loops near tail steps.
                if report_phase_confirmed and empty_action_streak >= 3:
                    log_messages.append(
                        "  guard: stopping due to repeated empty actions in report phase"
                    )
                    try:
                        agent_instance.stop()
                    except Exception:
                        pass
            results = agent_instance.state.last_result
            if results:
                for result in results:
                    if result.extracted_content:
                        runtime_extracted_contents.append(result.extracted_content)
                        payload = result.extracted_content
                        # Track successful native click labels (non-JSON text) for route-progress guards.
                        try:
                            if isinstance(payload, str) and payload.startswith("Clicked "):
                                m = re.search(r"\"([^\"]+)\"", payload)
                                if m:
                                    _remember_success_click_target(str(m.group(1) or "").strip())
                        except Exception:
                            pass
                        payload_dict = self._parse_json_payload(payload)
                        if payload_dict:
                            _update_stage_from_payload(payload_dict)
                            payload_type = str(payload_dict.get("type") or "")
                            if payload_type in {
                                "field_extraction_result",
                                "file_download_result",
                                "module_collection_result",
                                "open_report_result",
                            }:
                                should_checkpoint = True
                            if payload_type in {"open_report_result", "module_collection_result"}:
                                detected_report = _detect_report_from_text(str(payload_dict.get("report_name") or ""))
                                if detected_report:
                                    current_report_name = detected_report
                            if payload_type == "click_in_content_result":
                                target_text = str(payload_dict.get("target_text") or "").strip()
                                clicked_ok = bool(payload_dict.get("clicked"))
                                click_error = str(payload_dict.get("error") or "").strip()
                                log_messages.append(
                                    f"  result: click target='{target_text}' clicked={clicked_ok} error='{click_error or '-'}'"
                                )
                                if target_text and not clicked_ok:
                                    retry_target = target_text
                                    if _same_target(retry_target, pending_click_target):
                                        pending_click_retry_count += 1
                                    else:
                                        pending_click_target = retry_target
                                        pending_click_retry_count = 0
                                    log_messages.append(
                                        f"  guard: pending_click_target set to '{pending_click_target}' (click failed, retry_count={pending_click_retry_count})"
                                    )
                                elif target_text and clicked_ok and _same_target(target_text, pending_click_target):
                                    pending_click_target = None
                                    pending_click_retry_count = 0
                                    log_messages.append("  guard: pending_click_target cleared (click confirmed)")
                                if target_text and clicked_ok:
                                    _remember_success_click_target(target_text)
                            elif payload_type == "select_in_content_result":
                                target_text = str(payload_dict.get("target_text") or "").strip()
                                status = str(payload_dict.get("status") or "").strip().lower()
                                select_error = str(payload_dict.get("error") or "").strip()
                                if (
                                    target_text
                                    and status in {"failed", "blocked_repeated_failure"}
                                    and select_error in {"target_control_not_found", "target_label_mismatch"}
                                ):
                                    retry_tail = _pick_recent_route_tail_for_failed_select(target_text)
                                    if retry_tail:
                                        pending_click_target = retry_tail
                                        pending_click_retry_count = 0
                                        # Invalidate optimistic route-progress memory for this tail.
                                        recent_success_click_targets[:] = [
                                            t for t in recent_success_click_targets if not _same_target(t, retry_tail)
                                        ]
                                        log_messages.append(
                                            f"  guard: select failed ({select_error}), force re-position route tail '{retry_tail}'"
                                        )
                        if (
                            "field_extraction_result" in payload
                            or "file_download_result" in payload
                            or "module_route_navigation_result" in payload
                            or "click_in_content_result" in payload
                            or "hover_in_content_result" in payload
                            or "select_in_content_result" in payload
                        ):
                            report_phase_confirmed = True
                        # native extract payload
                        if "<query>" in payload and "<result>" in payload:
                            should_checkpoint = True
                        preview = result.extracted_content[:160]
                        log_messages.append(f"  extracted: {preview}")
                        if module_stage >= 0 and report_modules:
                            stage_name = report_modules[module_stage].module_key if module_stage < len(report_modules) else str(module_stage)
                            log_messages.append(f"  stage_lock: current_module_stage={module_stage}({stage_name})")
                    if result.error:
                        log_messages.append(f"  error: {result.error}")
            recovery = await _recover_if_unexpected(agent_instance)
            if recovery.get("recovered"):
                log_messages.append("  guard: unexpected page detected and auto-recovered")
            elif recovery.get("before", {}).get("is_unexpected"):
                log_messages.append("  guard: unexpected page detected but recovery failed")

            if should_checkpoint and self.session.task_dir and profile is not None:
                try:
                    self._persist_incremental_outputs(
                        output_dir=self.session.task_dir,
                        profile=profile,
                        task_intent=self.session.intent,
                        extracted_contents=runtime_extracted_contents,
                        log_messages=log_messages,
                        current_report=current_report_name,
                        step_no=step_num,
                    )
                    log_messages.append(
                        f"  checkpoint: saved incremental outputs at step {step_num} (current_report={current_report_name or 'N/A'})"
                    )
                except Exception as exc:  # noqa: BLE001
                    log_messages.append(f"  checkpoint_error: {exc}")
            log_messages.append(f"{phase_name} step {step_num - 1} finished")

        async def run_agent():
            # Disable pre-action guard by default to avoid over-filtering first-step actions.
            # Re-enable only when explicitly needed via env.
            enable_pre_guard = str(os.getenv("YUNTU_ENABLE_PRE_GUARD", "0")).strip().lower() in {
                "1",
                "true",
                "yes",
                "on",
            }
            if not enable_pre_guard:
                log_messages.append("  pre_guard: disabled")
                old_callback = getattr(agent, "register_new_step_callback", None)
                agent.register_new_step_callback = _pre_action_pending_click_retry_guard
                try:
                    return await agent.run(on_step_start=on_step_start, on_step_end=on_step_end)
                finally:
                    agent.register_new_step_callback = old_callback

            old_callback = getattr(agent, "register_new_step_callback", None)
            async def _combined_pre_step_guard(browser_state_summary, model_output, step_no: int):
                await _pre_action_pending_click_retry_guard(browser_state_summary, model_output, step_no)
                await _pre_action_guard(browser_state_summary, model_output, step_no)

            agent.register_new_step_callback = _combined_pre_step_guard
            try:
                log_messages.append("  pre_guard: enabled")
                return await agent.run(on_step_start=on_step_start, on_step_end=on_step_end)
            finally:
                agent.register_new_step_callback = old_callback

        agent_task = asyncio.create_task(run_agent())
        self.session.agent = agent
        self.session.agent_task = agent_task
        stopped = False

        while not agent_task.done():
            if agent.state.stopped:
                stopped = True
                log_messages.append(f"{phase_name} stopped by user")
                if not agent_task.done():
                    agent_task.cancel()
                    try:
                        await agent_task
                    except (asyncio.CancelledError, Exception):
                        pass
                break

            if len(log_messages) > last_log_count:
                log_text = "\n".join(log_messages[-40:])
                history[-1]["content"] = (
                    task_summary + f"\n\nRunning {phase_name}...\n\n---\n**Realtime logs**\n{log_text}"
                )
                yield history.copy(), None, stopped
                last_log_count = len(log_messages)
            await asyncio.sleep(0.5)

        result = None
        if not stopped:
            try:
                result = await agent_task
            except asyncio.CancelledError:
                stopped = True

        final_log_text = "\n".join(log_messages[-60:])
        return_text = task_summary + f"\n\n{phase_name} finished.\n\n---\n**Execution logs**\n{final_log_text}"
        yield history.copy(), (result, final_log_text, return_text), stopped

    async def process(self, history: list):
        await self.session.stop_current_agent()
        if not history:
            return
        last_msg = history[-1]
        if not isinstance(last_msg, dict) or last_msg.get("role") != "user":
            return
        user_message = last_msg.get("content", "")
        if not user_message:
            return

        history.append({"role": "assistant", "content": "..."})
        yield history.copy()

        history[-1]["content"] = "Analyzing intent with multi-turn context..."
        yield history.copy()

        conversation_history = [
            {"role": msg.get("role", "unknown"), "content": msg.get("content", "")}
            for msg in history[:-1]
            if isinstance(msg, dict) and msg.get("role") in ("user", "assistant")
        ]
        previous_intent = self.session.intent
        task_intent: YuntuTask = await extract_intent(
            user_message,
            conversation_history=conversation_history,
            previous_intent=previous_intent,
        )
        self.session.intent = task_intent
        self._print_intent_debug(task_intent)

        validation = validate_intent(task_intent)
        if not validation.is_valid:
            history[-1]["content"] = validation.message
            yield history.copy()
            return

        history[-1]["content"] = "Checking login state..."
        yield history.copy()
        if not await check_profile_exists():
            history[-1]["content"] = (
                "No persisted login detected.\n\n"
                "Please login manually in the opened browser window, then execution will continue."
            )
            yield history.copy()
            await perform_manual_login()
            history[-1]["content"] = "Login completed. Preparing opt execution..."
            yield history.copy()

        profile_name = resolve_profile_name(task_intent)
        profile = build_opt_profile(profile_name)
        brand_name_en = get_brand_english_name(task_intent.brand_name)
        report_targets = self._build_report_targets(task_intent)
        report_targets_text = ", ".join(report_targets) if report_targets else "none"

        task_timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        task_dir = TASKS_BASE_DIR / f"opt_task_{task_timestamp}"
        task_dir.mkdir(parents=True, exist_ok=True)
        self.session.task_dir = task_dir

        task_summary = (
            "opt run started\n\n"
            f"- profile: {profile.profile_name}\n"
            f"- brand: {task_intent.brand_name or 'N/A'}\n"
            f"- report: {task_intent.report_name or 'from categories'}\n"
            f"- report_targets: {report_targets_text}\n"
            f"- kol_content_date_type: {task_intent.kol_content_date_type or 'near30(default)'}\n"
            f"- kol_content_date_range: {task_intent.kol_content_date_range or 'N/A'}\n"
            "- mode: AGENT(tool-driven)"
        )
        history[-1]["content"] = task_summary
        yield history.copy()

        browser = None
        run_agent = None
        try:
            browser_env = os.environ.copy()
            for key in ("HTTP_PROXY", "HTTPS_PROXY", "http_proxy", "https_proxy"):
                browser_env.pop(key, None)
            no_proxy_value = "127.0.0.1,localhost,yuntu.oceanengine.com,.oceanengine.com"
            for key in ("NO_PROXY", "no_proxy"):
                existing = browser_env.get(key, "")
                browser_env[key] = no_proxy_value if not existing else f"{existing},{no_proxy_value}"

            browser = Browser(
                user_data_dir=str(PROFILE_DIR),
                storage_state=str(STORAGE_STATE_FILE),
                downloads_path=str(task_dir),
                headless=False,
                # Keep BrowserSession alive until the end of this single-agent run.
                keep_alive=True,
                allowed_domains=["yuntu.oceanengine.com", "*.oceanengine.com", "oceanengine.com"],
                args=["--no-proxy-server", "--disable-quic"],
                env=browser_env,
                cross_origin_iframes=True,
                max_iframes=30,
                wait_between_actions=0.6,
                minimum_wait_page_load_time=0.35,
                wait_for_network_idle_page_load_time=0.6,
            )
            (task_dir / "conversations").mkdir(parents=True, exist_ok=True)

            run_agent = Agent(
                task=build_agent_prompt(task_intent, profile, brand_name_en),
                llm=ChatGoogle(
                    model=GEMINI_MODEL,
                    api_key=API_KEY,
                    temperature=0.0,
                    thinking_budget=0,
                    max_output_tokens=6144,
                ),
                browser=browser,
                # Keep navigation capability stable; module phase is constrained by prompt/tools.
                tools=self._build_tools("agent_first"),
                max_actions_per_step=1,
                flash_mode=True,
                use_thinking=False,
                use_vision=False,
                max_history_items=6,
                llm_timeout=180,
                step_timeout=210,
                file_system_path=str(task_dir),
                save_conversation_path=str(task_dir / "conversations"),
            )

            run_result = None
            run_log = ""
            run_text = task_summary + "\n\nAGENT running..."
            async for updated, phase_result, stopped in self._run_phase(
                agent=run_agent,
                history=history,
                task_summary=task_summary,
                phase_name="AGENT",
                profile=profile,
            ):
                yield updated
                if phase_result is not None:
                    run_result, run_log, run_text = phase_result
                    history[-1]["content"] = run_text
                    yield history.copy()
                if stopped:
                    partial_history = getattr(run_agent, "history", None)
                    if partial_history is not None:
                        await self._handle_post_processing(
                            partial_history,
                            history,
                            task_summary,
                            run_log,
                            profile,
                            stop_reason="stopped_by_user",
                        )
                    else:
                        history[-1]["content"] = run_text + "\n\nStopped."
                    yield history.copy()
                    return

            if run_result is not None:
                await self._handle_post_processing(run_result, history, task_summary, run_log, profile)
                yield history.copy()
            else:
                partial_history = getattr(run_agent, "history", None)
                if partial_history is not None:
                    await self._handle_post_processing(
                        partial_history,
                        history,
                        task_summary,
                        run_log,
                        profile,
                        stop_reason="result_none_or_interrupted",
                    )
                else:
                    run_log_preview = "\n".join(run_log.splitlines()[-40:]) if run_log else "(no logs)"
                    history[-1]["content"] = (
                        task_summary
                        + "\n\nAGENT failed to produce result (result is None)."
                        + "\n\n---\n**Logs (tail)**\n"
                        + run_log_preview
                    )
                yield history.copy()
        except Exception as e:
            partial_history = getattr(run_agent, "history", None) if run_agent is not None else None
            if partial_history is not None and self.session.task_dir:
                await self._handle_post_processing(
                    partial_history,
                    history,
                    task_summary,
                    log_text=f"Exception: {type(e).__name__}: {e}",
                    profile=profile,
                    stop_reason=f"exception:{type(e).__name__}",
                )
                history[-1]["content"] += f"\n\nException captured:\n```\n{str(e)}\n```"
            else:
                history[-1]["content"] = f"Task execution failed:\n```\n{str(e)}\n```"
            yield history.copy()
        finally:
            self.session.clear()
            if browser:
                try:
                    await browser.stop()
                except Exception:
                    pass

    async def _handle_post_processing(
        self,
        result,
        history: list,
        task_summary: str,
        log_text: str,
        profile,
        stop_reason: str | None = None,
    ) -> None:
        output_dir = self.session.task_dir
        if not output_dir:
            history[-1]["content"] = task_summary + "\n\nNo task output directory available."
            return

        conversations_dir = output_dir / "conversations"
        execution_log_path = output_dir / "execution_logs.txt"
        try:
            conversations_dir.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass
        try:
            execution_log_path.write_text(log_text or "", encoding="utf-8")
        except Exception:
            execution_log_path = Path("N/A")

        try:
            payload = {
                "final_result": result.final_result(),
                "extracted_content": result.extracted_content(),
            }
            (output_dir / "agent_output.json").write_text(
                json.dumps(payload, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception:
            pass

        raw_json_path = export_agent_output_to_json(
            agent_final_result=result.final_result(),
            agent_extracted_content=result.extracted_content(),
            output_dir=output_dir,
            report_name=(self.session.intent.report_name if self.session.intent else None),
            brand_name=(self.session.intent.brand_name if self.session.intent else None),
        )

        structured_output = None
        try:
            structured_output = result.structured_output
            if structured_output is None:
                structured_output = result.get_structured_output(YuntuAgentStructuredOutput)
        except Exception:
            structured_output = None

        structured_path = output_dir / "agent_structured_output.json"
        business_summary_path = output_dir / "business_summary.json"
        completeness_text = "\n- completeness: unavailable (structured output parse failed)"
        report_targets = self._build_report_targets(self.session.intent) if self.session.intent else []
        final_result_text = result.final_result() if hasattr(result, "final_result") else ""
        hint_text = "\n".join([str(final_result_text or ""), str(log_text or "")]).strip()

        extracted_contents = result.extracted_content() if hasattr(result, "extracted_content") else []
        report_period_map = self._extract_report_period_map(extracted_contents)
        self._sync_period_map_to_intent(self.session.intent, report_period_map)
        default_period_label = self._infer_default_period_label_from_intent(self.session.intent)
        recovery_output = self._build_partial_structured_output(
            profile=profile,
            extracted_contents=extracted_contents,
            output_dir=output_dir,
            report_name=(self.session.intent.report_name if self.session.intent else None),
            brand_name=(self.session.intent.brand_name if self.session.intent else None),
            stop_reason=stop_reason or "recovery",
            default_period_label=default_period_label,
        )
        if structured_output is not None:
            structured_output = self._reconcile_structured_output_with_tools(
                structured_output=structured_output,
                profile=profile,
                extracted_contents=extracted_contents,
            )
            structured_output = self._merge_structured_with_recovery(structured_output, recovery_output)
            structured_output = self._ensure_report_targets_in_structured_output(structured_output, report_targets)
            structured_output = self._apply_report_field_hints_to_structured_output(
                structured_output=structured_output,
                profile=profile,
                report_targets=report_targets,
                hint_text=hint_text,
            )
            print(
                "[structured_output.report_names]",
                [str(r.report_name) for r in (structured_output.reports or [])],
            )
            structured_path.write_text(structured_output.model_dump_json(indent=2), encoding="utf-8")
            completeness = check_completeness(structured_output, profile)
            if completeness.ok:
                completeness_text = "\n- completeness: PASS"
            else:
                missing_lines = []
                if completeness.missing_modules:
                    missing_lines.append("missing_modules=" + ", ".join(completeness.missing_modules))
                if completeness.missing_fields:
                    missing_lines.append("missing_fields=" + ", ".join(completeness.missing_fields))
                if completeness.missing_files:
                    missing_lines.append("missing_files=" + ", ".join(completeness.missing_files))
                completeness_text = "\n- completeness: FAIL\n- " + "\n- ".join(missing_lines)
        else:
            structured_output = recovery_output
            structured_output = self._ensure_report_targets_in_structured_output(structured_output, report_targets)
            structured_output = self._apply_report_field_hints_to_structured_output(
                structured_output=structured_output,
                profile=profile,
                report_targets=report_targets,
                hint_text=hint_text,
            )
            print(
                "[structured_output.report_names]",
                [str(r.report_name) for r in (structured_output.reports or [])],
            )
            structured_path.write_text(structured_output.model_dump_json(indent=2), encoding="utf-8")
            completeness_text = "\n- completeness: PARTIAL_RECOVERY"

        try:
            business_summary = self._build_business_summary_json(
                structured_output=structured_output,
                profile=profile,
                task_intent=self.session.intent,
                report_period_map=report_period_map,
                output_dir=output_dir,
            )
            business_summary = self._merge_business_summary_with_progress(output_dir, business_summary)
            self._update_report_progress_map(output_dir, business_summary)
            business_summary_path.write_text(
                json.dumps(business_summary, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception:
            business_summary_path = Path("N/A")

        final_output = result.final_result() or "(no final text)"
        history[-1]["content"] = (
            task_summary
            + f"\n\nTask finished.\n\n{final_output}\n\n---\n"
            f"**Execution logs**\n{log_text}\n\n---\n"
            "**Output files**\n"
            f"- conversations_dir: `{conversations_dir}`\n"
            f"- execution_log_txt: `{execution_log_path}`\n"
            f"- raw_agent_json: `{raw_json_path}`\n"
            f"- structured_output_json: `{structured_path}`"
            f"\n- business_summary_json: `{business_summary_path}`"
            f"{completeness_text}\n"
        )

        self.session.task_dir = None
