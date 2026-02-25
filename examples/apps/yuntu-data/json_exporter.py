"""
JSON 导出模块 - 将 Agent 输出保存为 JSON 文件

阶段一：直接按提示词中要求的 JSON 格式保存 Agent 输出。
阶段二（待实现）：将 JSON 数据写入 Excel 模板。
"""
from __future__ import annotations

import json
import re
from pathlib import Path
from datetime import datetime

from data_models import YuntuDataReport
from json_utils import extract_json_from_text


def export_agent_output_to_json(
	agent_final_result: str | None,
	agent_extracted_content: list[str] | None,
	output_dir: Path | str,
	report_name: str | None = None,
	brand_name: str | None = None,
) -> Path:
	"""
	将 Agent 输出保存为 JSON 文件。

	优先尝试从 Agent 输出中解析结构化 JSON（与提示词中要求的格式一致）；
	若解析失败，则保存原始文本到 raw_text 字段。

	Args:
		agent_final_result: Agent.run() 返回的 final_result
		agent_extracted_content: Agent.run() 返回的 extracted_content 列表
		output_dir: 输出目录
		report_name: 报告名（可选，用于文件命名）
		brand_name: 品牌名（可选，用于文件命名）

	Returns:
		生成的 JSON 文件路径
	"""
	output_dir = Path(output_dir)
	output_dir.mkdir(parents=True, exist_ok=True)

	# 拼接所有文本
	all_text = agent_final_result or ""
	if agent_extracted_content:
		all_text += "\n" + "\n".join(c for c in agent_extracted_content if c)

	# 尝试解析 JSON
	parsed = extract_json_from_text(all_text)

	if parsed is not None:
		output_data = parsed
	else:
		# 解析失败 — 保留原始文本，方便后续人工检查或二次处理
		output_data = {
			"_parse_warning": "无法从 Agent 输出中解析 JSON，已保存原始文本",
			"raw_text": all_text,
		}

	# 生成文件名
	timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
	brand = brand_name or "未知品牌"
	report = report_name or "report"
	# 清理文件名中的非法字符
	brand = re.sub(r'[<>:"/\\|?*]', '_', brand)
	report = re.sub(r'[<>:"/\\|?*]', '_', report)
	filename = f"{brand}_{report}_{timestamp}.json"

	file_path = output_dir / filename
	file_path.write_text(
		json.dumps(output_data, ensure_ascii=False, indent=2),
		encoding="utf-8",
	)
	return file_path


def export_structured_reports_to_json(
	reports: list[YuntuDataReport],
	output_dir: Path | str,
	filename: str | None = None,
) -> Path:
	"""
	将 Pydantic 结构化数据导出为 JSON（作为补充/调试用途）。

	Args:
		reports: YuntuDataReport 列表
		output_dir: 输出目录
		filename: 自定义文件名（可选）

	Returns:
		生成的 JSON 文件路径
	"""
	output_dir = Path(output_dir)
	output_dir.mkdir(parents=True, exist_ok=True)

	if not filename:
		timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
		filename = f"structured_data_{timestamp}.json"
	if not filename.endswith(".json"):
		filename += ".json"

	data = [r.model_dump(mode="json") for r in reports]
	file_path = output_dir / filename
	file_path.write_text(
		json.dumps(data, ensure_ascii=False, indent=2),
		encoding="utf-8",
	)
	return file_path
