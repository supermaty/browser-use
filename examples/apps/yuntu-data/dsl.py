from __future__ import annotations

import re


def _split_dsl(route_dsl: str) -> list[str]:
    text = (route_dsl or "").strip()
    if not text:
        return []
    # Prefer ">" as separator for readability and stability.
    parts = [item.strip() for item in re.split(r"\s*>\s*", text) if item.strip()]
    return parts


def parse_route_dsl(route_dsl: str) -> tuple[list[str], list[InteractionStep]]:
    """
    Parse compact route DSL into route_path + interaction_steps.

    Examples:
    - 人群分析 > 人群画像分析 > 画像分析 > [select]5A人群资产=投后 > [click]5A人群总资产 > [download]audience_profile_5a_asset
    """
    parts = _split_dsl(route_dsl)
    # Delayed import to avoid circular dependency with specs.py
    from specs import InteractionStep

    route_path: list[str] = []
    steps: list[InteractionStep] = []
    encountered_explicit_op = False

    for part in parts:
        m = re.match(r"^\[(?P<op>[a-zA-Z_]+)\](?P<body>.*)$", part)
        if not m:
            if not encountered_explicit_op:
                route_path.append(part)
            else:
                steps.append(InteractionStep(op="click", target=part, required=True, retries=2))
            continue

        encountered_explicit_op = True
        op = m.group("op").strip().lower()
        body = m.group("body").strip()

        if op == "click":
            steps.append(InteractionStep(op="click", target=body, required=True, retries=2))
            continue
        if op == "download":
            file_key = body.strip() or None
            steps.append(InteractionStep(op="download", file_key=file_key, required=True, retries=2))
            continue
        if op == "wait":
            try:
                seconds = float(body)
            except Exception:
                seconds = 0.8
            steps.append(InteractionStep(op="wait", seconds=seconds, required=False))
            continue
        if op in {"select", "select_path", "input", "set_date_range"}:
            if "=" in body:
                target, value = [x.strip() for x in body.split("=", 1)]
            else:
                target, value = body.strip(), None
            value_from = value if value and value.startswith(("params.", "module.", "context.")) else None
            literal_value = None if value_from else value
            if op in {"select", "select_path"}:
                steps.append(
                    InteractionStep(
                        op="select",
                        target=target,
                        value=literal_value,
                        value_from=value_from,
                        required=True,
                        retries=2,
                    )
                )
            elif op == "input":
                steps.append(
                    InteractionStep(
                        op="input",
                        target=target,
                        value=literal_value,
                        value_from=value_from,
                        required=True,
                        retries=2,
                    )
                )
            else:
                steps.append(
                    InteractionStep(
                        op="set_date_range",
                        target=target,
                        value=literal_value,
                        value_from=value_from,
                        required=True,
                        retries=2,
                    )
                )
            continue
        if op in {"extract", "extract_fields"}:
            keys = [k.strip() for k in re.split(r"[,\s]+", body) if k.strip()]
            steps.append(InteractionStep(op="extract_fields", field_keys=keys, required=False, retries=1))
            continue
        if op in {"hover", "hover_extract"}:
            keys = [k.strip() for k in re.split(r"[,\s]+", body) if k.strip()]
            steps.append(InteractionStep(op="hover", field_keys=keys, required=False, retries=1))
            continue

        # Unknown op fallback to explicit click on full token.
        steps.append(InteractionStep(op="click", target=part, required=True, retries=2))

    return route_path, steps
