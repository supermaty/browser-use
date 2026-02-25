"""
JSON 解析与正则提取的公共工具模块。

提供：
- extract_json_from_text: 从 Agent 输出文本中提取 JSON 对象
- extract_fields_by_patterns: 数据驱动的正则字段提取
"""
import json
import re
from typing import Any

from utils import parse_number


# 匹配 N/A、null 等无效值
_NA_VALUES = frozenset({"n/a", "na", "-", "--", "null", "none", ""})


def extract_json_from_text(text: str) -> dict | list | None:
	"""
	从 Agent 输出文本中提取 JSON 对象。

	按优先级依次尝试：
	1. 直接 json.loads 整段文本
	2. ```json ... ``` code block
	3. ``` ... ``` code block
	4. 最外层 { ... } 或 [ ... ]

	Returns:
		解析出的 dict/list，或 None。
	"""
	if not text or not text.strip():
		return None

	# 1. 直接解析
	try:
		return json.loads(text.strip())
	except (json.JSONDecodeError, ValueError):
		pass

	# 2. ```json ... ```
	m = re.search(r"```json\s*\n?(.*?)```", text, re.DOTALL)
	if m:
		try:
			return json.loads(m.group(1).strip())
		except (json.JSONDecodeError, ValueError):
			pass

	# 3. ``` ... ```
	m = re.search(r"```\s*\n?(.*?)```", text, re.DOTALL)
	if m:
		try:
			return json.loads(m.group(1).strip())
		except (json.JSONDecodeError, ValueError):
			pass

	# 4. 最外层 { } 或 [ ]
	for opener, closer in [("{", "}"), ("[", "]")]:
		start = text.find(opener)
		if start == -1:
			continue
		end = text.rfind(closer)
		if end > start:
			try:
				return json.loads(text[start : end + 1])
			except (json.JSONDecodeError, ValueError):
				pass

	return None


def extract_fields_by_patterns(
	text: str,
	patterns: dict[str, str],
) -> dict[str, Any]:
	"""
	数据驱动的正则字段提取。

	对 patterns 中的每个 (field_name, regex) 执行 re.search，
	清理匹配值（移除方括号标记、过滤 N/A），并用 parse_number 转换。

	Args:
		text: 待搜索文本
		patterns: {字段名: 正则表达式}，正则需包含一个捕获组

	Returns:
		{字段名: 解析后的值} 字典（仅包含成功提取的字段）
	"""
	data: dict[str, Any] = {}
	for key, pattern in patterns.items():
		match = re.search(pattern, text)
		if not match:
			continue
		value_str = match.group(1).strip()
		# 移除可能的方括号标记，如 [待确认]
		value_str = re.sub(r"^\[.*?\]", "", value_str).strip()
		if value_str.lower() in _NA_VALUES:
			continue
		data[key] = parse_number(value_str)
	return data
