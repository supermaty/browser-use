from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, ConfigDict, Field


class MenuTraceItem(BaseModel):
    model_config = ConfigDict(extra="forbid")

    step_no: int
    action: str
    target: str
    success: bool
    note: str | None = None


class RouteTraceItem(BaseModel):
    model_config = ConfigDict(extra="forbid")

    step_no: int
    target: str
    clicked: bool
    selector: str | None = None
    error: str | None = None


class FieldExtractionResult(BaseModel):
    model_config = ConfigDict(extra="forbid")

    field_key: str
    status: Literal["ok", "not_found", "hover_failed", "parse_failed", "error", "skipped"]
    source: Literal["direct", "hover", "auto", "none"]
    value: float | int | str | None = None
    raw_text: str | None = None
    error: str | None = None


class FileDownloadResult(BaseModel):
    model_config = ConfigDict(extra="forbid")

    file_key: str
    status: Literal["ok", "not_found", "download_failed", "timeout", "error", "skipped"]
    file_path: str | None = None
    file_name: str | None = None
    clicked_selector: str | None = None
    error: str | None = None


class ModuleExecutionResult(BaseModel):
    model_config = ConfigDict(extra="forbid")

    module_key: str
    module_label: str | None = None
    status: Literal["success", "partial", "failed", "skipped"]
    route_trace: list[RouteTraceItem] = Field(default_factory=list)
    fields: list[FieldExtractionResult] = Field(default_factory=list)
    period_results: dict[str, list[FieldExtractionResult]] = Field(default_factory=dict)
    files: list[FileDownloadResult] = Field(default_factory=list)
    touchpoint: str | None = None
    note: str | None = None


class ReportExecutionResult(BaseModel):
    model_config = ConfigDict(extra="forbid")

    report_name: str
    report_type: Literal["single", "star", "bid", "unknown"] = "unknown"
    status: Literal["success", "partial", "failed"]
    modules: list[ModuleExecutionResult] = Field(default_factory=list)
    fields: list[FieldExtractionResult] = Field(default_factory=list)
    files: list[FileDownloadResult] = Field(default_factory=list)
    note: str | None = None


class YuntuAgentStructuredOutput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    task_status: Literal["success", "partial", "failed"]
    switched_brand: str | None = None
    menu_trace: list[MenuTraceItem] = Field(default_factory=list)
    reports: list[ReportExecutionResult] = Field(default_factory=list)
    global_modules: list[ModuleExecutionResult] = Field(default_factory=list)
    summary: str | None = None


class CompletenessResult(BaseModel):
    model_config = ConfigDict(extra="forbid")

    ok: bool
    missing_modules: list[str] = Field(default_factory=list)
    missing_fields: list[str] = Field(default_factory=list)
    missing_files: list[str] = Field(default_factory=list)
