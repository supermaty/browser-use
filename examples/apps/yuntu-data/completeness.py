from __future__ import annotations

from schemas import CompletenessResult, YuntuAgentStructuredOutput
from specs import BrandProfileSpec, ModuleSpec


def _collect_enabled_module_file_keys(profile: BrandProfileSpec) -> set[str]:
    keys: list[str] = []
    for module in profile.module_specs:
        if module.file_keys:
            keys.extend(module.file_keys)
        for step in module.interaction_steps:
            if step.file_key:
                keys.append(step.file_key)
            if step.file_keys:
                keys.extend(step.file_keys)
    return set(k for k in keys if k)


def _check_module(
    module,
    module_spec: ModuleSpec,
    field_required: dict[str, bool],
    file_required: dict[str, bool],
    prefix: str,
) -> tuple[list[str], list[str], list[str]]:
    missing_modules: list[str] = []
    missing_fields: list[str] = []
    missing_files: list[str] = []

    if module.status in {"failed", "skipped"}:
        missing_modules.append(f"{prefix}:{module_spec.module_key}")

    field_status = {item.field_key: item.status for item in module.fields}
    file_status = {item.file_key: item.status for item in module.files}

    for key in module_spec.field_keys:
        if field_required.get(key, True) and field_status.get(key) != "ok":
            missing_fields.append(f"{prefix}:{module_spec.module_key}:{key}")

    for key in module_spec.file_keys:
        if file_required.get(key, False) and file_status.get(key) != "ok":
            missing_files.append(f"{prefix}:{module_spec.module_key}:{key}")

    return missing_modules, missing_fields, missing_files


def check_completeness(
    structured_output: YuntuAgentStructuredOutput,
    profile: BrandProfileSpec,
) -> CompletenessResult:
    field_required = {item.field_key: item.required for item in profile.field_specs}
    enabled_file_keys = _collect_enabled_module_file_keys(profile)
    file_required = {item.file_key: item.required for item in profile.file_specs if item.file_key in enabled_file_keys}

    required_report_modules = [m for m in profile.module_specs if m.required and m.requires_report]
    required_global_modules = [m for m in profile.module_specs if m.required and not m.requires_report]

    missing_modules: list[str] = []
    missing_fields: list[str] = []
    missing_files: list[str] = []

    # Report-level module completeness
    if required_report_modules and not structured_output.reports:
        for module_spec in required_report_modules:
            missing_modules.append(f"NO_REPORTS:{module_spec.module_key}")

    for report in structured_output.reports:
        module_map = {module.module_key: module for module in report.modules}

        for module_spec in required_report_modules:
            module = module_map.get(module_spec.module_key)
            prefix = report.report_name
            if module is None:
                missing_modules.append(f"{prefix}:{module_spec.module_key}")
                continue

            mm, mf, mfi = _check_module(
                module=module,
                module_spec=module_spec,
                field_required=field_required,
                file_required=file_required,
                prefix=prefix,
            )
            missing_modules.extend(mm)
            missing_fields.extend(mf)
            missing_files.extend(mfi)

    # Global module completeness
    global_map = {module.module_key: module for module in structured_output.global_modules}
    for module_spec in required_global_modules:
        prefix = "GLOBAL"
        module = global_map.get(module_spec.module_key)
        if module is None:
            missing_modules.append(f"{prefix}:{module_spec.module_key}")
            continue

        mm, mf, mfi = _check_module(
            module=module,
            module_spec=module_spec,
            field_required=field_required,
            file_required=file_required,
            prefix=prefix,
        )
        missing_modules.extend(mm)
        missing_fields.extend(mf)
        missing_files.extend(mfi)

    return CompletenessResult(
        ok=not missing_modules and not missing_fields and not missing_files,
        missing_modules=sorted(set(missing_modules)),
        missing_fields=sorted(set(missing_fields)),
        missing_files=sorted(set(missing_files)),
    )
