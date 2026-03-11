from __future__ import annotations

from typing import Any, Literal

from intent import YuntuTask
from pydantic import BaseModel, ConfigDict, Field

from dsl import parse_route_dsl


class FieldSpec(BaseModel):
    model_config = ConfigDict(extra="forbid")

    field_key: str
    label: str | None = None
    alt_labels: list[str] = Field(default_factory=list)
    required: bool = True
    mode: Literal["direct", "hover", "auto"] = "auto"
    value_type: Literal["number", "text"] = "number"
    value_selectors: list[str] = Field(default_factory=list)
    hover_selectors: list[str] = Field(default_factory=list)
    tooltip_selectors: list[str] = Field(
        default_factory=lambda: ["[role='tooltip']", ".ant-tooltip-inner", ".tippy-content"]
    )
    regex: str | None = None


class FileSpec(BaseModel):
    model_config = ConfigDict(extra="forbid")

    file_key: str
    required: bool = False
    trigger_selectors: list[str] = Field(default_factory=list)
    expected_name_keywords: list[str] = Field(default_factory=list)
    timeout_seconds: int = 20


class InteractionStep(BaseModel):
    model_config = ConfigDict(extra="forbid")

    op: Literal[
        "ensure_visible",
        "click",
        "hover",
        "select",
        "select_path",
        "input",
        "set_date_range",
        "extract",
        "extract_fields",
        "download",
        "wait",
    ]
    target: str | None = None
    selector: str | None = None
    value: str | None = None
    value_from: str | None = None
    path: list[str] = Field(default_factory=list)
    path_from: str | None = None
    field_keys: list[str] = Field(default_factory=list)
    file_key: str | None = None
    file_keys: list[str] = Field(default_factory=list)
    seconds: float | None = None
    when: str | None = None
    required: bool = True
    retries: int = 1


class ModuleSpec(BaseModel):
    model_config = ConfigDict(extra="forbid")

    module_key: str
    label: str
    required: bool = True
    requires_report: bool = True
    route_path: list[str] = Field(default_factory=list)
    route_selectors: list[str] = Field(
        default_factory=lambda: [
            "[role='menuitem']",
            "aside a",
            "aside button",
            ".ant-menu-item",
            ".ant-tabs-tab",
            "a",
            "button",
            "li",
            "span",
        ]
    )
    page_ready_selectors: list[str] = Field(default_factory=lambda: ["body"])
    interaction_steps: list[InteractionStep] = Field(default_factory=list)
    acquisition_mode: Literal["direct", "download", "mixed"] = "mixed"
    field_keys: list[str] = Field(default_factory=list)
    file_keys: list[str] = Field(default_factory=list)
    touchpoint_candidates: list[str] = Field(default_factory=list)
    default_touchpoint: str | None = None
    date_type_candidates: list[str] = Field(default_factory=list)
    date_range_hint: str | None = None


class BrandProfileSpec(BaseModel):
    model_config = ConfigDict(extra="forbid")

    profile_name: str
    brand_aliases: list[str] = Field(default_factory=list)
    menu_path: list[str] = Field(default_factory=list)
    field_specs: list[FieldSpec] = Field(default_factory=list)
    file_specs: list[FileSpec] = Field(default_factory=list)
    module_specs: list[ModuleSpec] = Field(default_factory=list)
    report_search_selectors: list[str] = Field(
        default_factory=lambda: [
            "input[placeholder*='报告']",
            "input[placeholder*='搜索']",
            ".ant-input",
            "input[type='search']",
            "input",
        ]
    )
    report_row_selectors: list[str] = Field(
        default_factory=lambda: [
            "table tbody tr",
            ".ant-table-row",
            ".report-row",
            "[data-test*='report-row']",
            "[role='row']",
            ".list-item",
        ]
    )
    report_name_cell_selectors: list[str] = Field(
        default_factory=lambda: [
            "td:first-child",
            ".report-name",
            "[data-test*='report-name']",
            "[title]",
        ]
    )
    report_view_button_selectors: list[str] = Field(
        default_factory=lambda: ["button", "a", "[role='button']"]
    )
    report_list_tab_candidates: list[str] = Field(
        default_factory=list
    )


COMMON_DOWNLOAD_SELECTORS = [
    "[data-test*='下载']",
    "[data-testid*='下载']",
    "[data-test*='导出']",
    "[data-testid*='导出']",
    "button[title*='下载']",
    "button[aria-label*='下载']",
    "button[title*='导出']",
    "button[aria-label*='导出']",
    "[class*='download']",
    "[class*='export']",
    ".download-btn",
]

# Single source of truth in opt (DSL-only mode):
# Keep only route_dsl (and optional module-level metadata).
# Field/file specs are auto-derived from [extract]/[hover]/[download] in route_dsl.
DSL_MODULE_OVERRIDES: dict[str, dict[str, Any]] = {
    # "overview_post_report": {
    #     "label": "项目整体Overview",
    #     "required": True,
    #     "requires_report": True,
    #     "route_dsl": (
    #         "活动总览 > [extract]互动率,拉新人群规模,新客人数,新客占比 > "
    #         "流量分析 > [extract]曝光次数,曝光人数,完播率,完播数,7日回搜人数,7日回搜率 > "
    #         "转化分析 > [extract]转化金额,转化人数 > "
    #         "品牌形象 > 搜索分析 > 搜索趋势分析 > [download]"
    #     ),
    # },
    # "asset_5a_flow": {
    #     "label": "5A人群资产流转",
    #     "required": True,
    #     "requires_report": True,
    #     "route_dsl": (
    #         "人群分析 > 5A关系资产分析 > 5A人群资产 > "
    #         "[select]触点=全部/全部 > "
    #         "[extract]拉新人群规模,拉新比例,A1流转人群,A1流转率,A2流转人群,A2流转率,A3流转人群,A3流转率,A4流转人群,A4流转率,A5流转人群,A5流转率 > "
    #         "[hover]A1流转率/行业TOP5%品牌均值,A2流转率/行业TOP5%品牌均值,A3流转率/行业TOP5%品牌均值,A4流转率/行业TOP5%品牌均值,A5流转率/行业TOP5%品牌均值"
    #     ),
    # },
    # "人群包投放效果分析": {
    #     "label": "人群包投放效果分析",
    #     "required": False,
    #     "requires_report": True,
    #     "route_dsl": (
    #         "人群分析 > 人群画像分析 > 人群包投放效果分析 > [download]"
    #     ),
    # },
    # "audience_profile_5a": {
    #     "label": "投放人群画像-5A人群总资产",
    #     "required": False,
    #     "requires_report": True,
    #     "route_dsl": (
    #         "人群分析 > 人群画像分析 > 画像分析 > "
    #         "[select]5A人群资产=投后 > "
    #         "[select]触点=params.touchpoint > "
    #         "5A人群总资产 > [download]"
    #     ),
    # },
    # "audience_profile_a3": {
    #     "label": "投放人群画像-A3问询",
    #     "required": False,
    #     "requires_report": True,
    #     "route_dsl": (
    #         "人群分析 > 人群画像分析 > 画像分析 > "
    #         "[select]5A人群资产=投后 > "
    #         "[select]触点=params.touchpoint > "
    #         "A3问询 > [download]"
    #     ),
    # },
    "行业搜索洞察": {
        "label": "搜索与溢出价值",
        "required": False,
        "requires_report": False,
        "route_dsl": (
            "营销触点 > 行业搜索洞察 > "
            "[select]日期类型=params.current_insight.date_range_type > "
            "[set_date_range]日期=params.current_insight.insight_date_range > "
            "行业搜索趋势 > "
            "[extract]搜索次数,搜索人数,搜索次数环比,搜索人数环比 > "
            "行业核心品牌 > "
            "[extract]搜索流量SOV排名,搜索流量SOV（指数）"
        ),
    },
    "达人及内容复盘": {
        "label": "达人及内容复盘",
        "required": False,
        "requires_report": False,
        "route_dsl": (
            "营销触点 > 营销概览 > 爆文加热 > "
            "[select]日期类型=params.kol_content_date_type > "
            "[set_date_range]日期=params.kol_content_date_range > "
            "视频榜单 > [download]"
        ),
    },
}


def hydrate_field_spec(spec: FieldSpec) -> FieldSpec:
    label = (spec.label or spec.field_key).strip()
    return spec.model_copy(update={"label": label})


def _normalize_field_entry(entry: str | dict[str, Any]) -> FieldSpec:
    if isinstance(entry, str):
        return FieldSpec(field_key=entry, label=entry, required=True)
    data = dict(entry)
    data.setdefault("label", data.get("field_key"))
    return FieldSpec(**data)


def _collect_field_keys_from_steps(steps: list[InteractionStep]) -> tuple[list[str], set[str]]:
    ordered: list[str] = []
    hover_keys: set[str] = set()
    seen: set[str] = set()
    for step in steps:
        if step.op not in {"extract", "extract_fields", "hover"}:
            continue
        for key in step.field_keys:
            k = (key or "").strip()
            if not k:
                continue
            if k not in seen:
                seen.add(k)
                ordered.append(k)
            if step.op == "hover":
                hover_keys.add(k)
    return ordered, hover_keys


def _normalize_step_field_keys(steps: list[InteractionStep], explicit_specs: list[FieldSpec]) -> list[InteractionStep]:
    alias_to_key: dict[str, str] = {}
    for spec in explicit_specs:
        alias_to_key[spec.field_key] = spec.field_key
        if spec.label:
            alias_to_key[spec.label] = spec.field_key
        for alt in spec.alt_labels:
            alias_to_key[alt] = spec.field_key

    normalized_steps: list[InteractionStep] = []
    for step in steps:
        if step.op not in {"extract", "extract_fields", "hover"} or not step.field_keys:
            normalized_steps.append(step)
            continue
        mapped_keys: list[str] = []
        seen: set[str] = set()
        for token in step.field_keys:
            text = (token or "").strip()
            if not text:
                continue
            key = alias_to_key.get(text, text)
            if key in seen:
                continue
            seen.add(key)
            mapped_keys.append(key)
        normalized_steps.append(step.model_copy(update={"field_keys": mapped_keys}))
    return normalized_steps


def _normalize_file_entry(entry: str | dict[str, Any], module_key: str, index: int) -> FileSpec:
    default_file_key = f"{module_key}_file" if index == 0 else f"{module_key}_file_{index + 1}"
    if isinstance(entry, str):
        file_key = entry.strip() or default_file_key
        return FileSpec(
            file_key=file_key,
            required=False,
            trigger_selectors=list(COMMON_DOWNLOAD_SELECTORS),
            timeout_seconds=30,
        )
    data = dict(entry)
    if not str(data.get("file_key", "")).strip():
        data["file_key"] = default_file_key
    data.setdefault("trigger_selectors", list(COMMON_DOWNLOAD_SELECTORS))
    data.setdefault("timeout_seconds", 30)
    return FileSpec(**data)


def _infer_acquisition_mode(field_keys: list[str], file_keys: list[str]) -> Literal["direct", "download", "mixed"]:
    if field_keys and file_keys:
        return "mixed"
    if field_keys:
        return "direct"
    if file_keys:
        return "download"
    return "direct"


def _build_module(module_key: str, cfg: dict[str, Any]) -> tuple[ModuleSpec, list[FieldSpec], list[FileSpec]]:
    route_dsl = str(cfg.get("route_dsl", "") or "")
    route_path, interaction_steps = parse_route_dsl(route_dsl)

    # DSL-only mode: field specs are derived from [extract]/[hover] only.
    # Ignore cfg["fields"] overrides to keep configuration minimal and consistent.
    explicit_field_specs: list[FieldSpec] = []
    interaction_steps = _normalize_step_field_keys(interaction_steps, explicit_field_specs)
    step_field_keys, hover_keys = _collect_field_keys_from_steps(interaction_steps)
    merged_keys = step_field_keys
    field_specs: list[FieldSpec] = []
    for key in merged_keys:
        mode = "hover" if key in hover_keys else "auto"
        field_specs.append(FieldSpec(field_key=key, label=key, required=True, mode=mode))

    file_specs = [_normalize_file_entry(item, module_key, idx) for idx, item in enumerate(cfg.get("files") or [])]
    field_keys = [item.field_key for item in field_specs]
    file_keys = [item.file_key for item in file_specs]

    has_extract_step = any(step.op in {"extract", "extract_fields", "hover"} for step in interaction_steps)
    has_download_step = any(step.op == "download" for step in interaction_steps)

    if field_keys and not has_extract_step:
        interaction_steps.append(InteractionStep(op="extract", field_keys=field_keys, required=True))
    if has_download_step and not file_keys:
        auto_file_spec = _normalize_file_entry({}, module_key, 0)
        file_specs.append(auto_file_spec)
        file_keys.append(auto_file_spec.file_key)
    if file_keys and not has_download_step:
        interaction_steps.append(InteractionStep(op="download", file_keys=file_keys, required=True))

    # If DSL uses [download] without explicit key, bind to module file keys.
    if file_keys:
        for idx, step in enumerate(interaction_steps):
            if step.op != "download":
                continue
            if step.file_key or step.file_keys:
                continue
            interaction_steps[idx] = step.model_copy(update={"file_keys": list(file_keys)})

    module = ModuleSpec(
        module_key=module_key,
        label=str(cfg.get("label") or module_key),
        required=bool(cfg.get("required", True)),
        requires_report=bool(cfg.get("requires_report", True)),
        route_path=route_path,
        page_ready_selectors=list(cfg.get("page_ready_selectors") or ["body"]),
        interaction_steps=interaction_steps,
        acquisition_mode=_infer_acquisition_mode(field_keys, file_keys),
        field_keys=field_keys,
        file_keys=file_keys,
        touchpoint_candidates=list(cfg.get("touchpoint_candidates") or []),
        default_touchpoint=cfg.get("default_touchpoint"),
        date_type_candidates=list(cfg.get("date_type_candidates") or []),
        date_range_hint=cfg.get("date_range_hint"),
    )
    return module, field_specs, file_specs


def resolve_profile_name(task_intent: YuntuTask) -> str:
    _ = task_intent
    return "generic"


def build_opt_profile(profile_name: str) -> BrandProfileSpec:
    modules: list[ModuleSpec] = []
    field_specs: list[FieldSpec] = []
    file_specs: list[FileSpec] = []

    for module_key, cfg in DSL_MODULE_OVERRIDES.items():
        module, fields, files = _build_module(module_key, cfg)
        modules.append(module)
        field_specs.extend(fields)
        file_specs.extend(files)

    dedup_fields: dict[str, FieldSpec] = {}
    for item in field_specs:
        dedup_fields[item.field_key] = hydrate_field_spec(item)
    dedup_files: dict[str, FileSpec] = {}
    for item in file_specs:
        dedup_files[item.file_key] = item

    return BrandProfileSpec(
        profile_name=profile_name or "generic",
        brand_aliases=[],
        menu_path=[],
        field_specs=list(dedup_fields.values()),
        file_specs=list(dedup_files.values()),
        module_specs=modules,
    )


def get_profile_spec(profile_name: str) -> BrandProfileSpec:
    return build_opt_profile(profile_name)


def get_module_spec(profile_name: str, module_key: str) -> ModuleSpec | None:
    profile = get_profile_spec(profile_name)
    for module in profile.module_specs:
        if module.module_key == module_key:
            return module
    return None
