"""
数据提取模块 - 从Agent执行结果中提取结构化数据

提供两种解析策略：
1. JSON 解析（优先）：当 Agent 按照提示词输出 JSON 时，直接映射到 Pydantic 模型。
2. 正则解析（兜底）：从自由文本中用正则提取字段。
"""
from __future__ import annotations

import re
import os

from browser_use.agent.views import AgentHistoryList

from data_models import (
	ProjectOverview,
	A5AssetFlow,
	KOLContentReview,
	SearchInsight,
	AdFlowReview,
	TAPortraitReview,
	YuntuDataReport,
)
from utils import parse_number, parse_chinese_number
from json_utils import extract_json_from_text, extract_fields_by_patterns


def _build_report_from_json_item(
	item: dict,
	brand_name: str | None = None,
	date_range: str | None = None,
) -> YuntuDataReport:
	"""
	将提示词 JSON 格式中的单个报告对象映射到 YuntuDataReport。

	提示词要求的 JSON 格式:
	{
	  "品牌分类": "...",
	  "报告名": "...",
	  "项目整体": { "消耗金额": ..., "曝光次数": ..., ... },
	  "5A人群资产": { "投后人群规模": ..., ... }
	}
	"""
	report_name = item.get("报告名") or item.get("report_name")

	# --- 项目整体 ---
	proj_raw = item.get("项目整体") or item.get("project_overview") or {}
	project_overview = None
	if proj_raw:
		# 字段名映射：提示词 JSON key -> Pydantic field name
		field_map = {
			"消耗金额": "消耗金额",
			"曝光次数": "曝光次数",
			"曝光人数": "曝光人数",
			"互动率": "互动率",
			"互动量": "互动量",
			"流转A3人数": "A3流转人数",
			"A3流转人数": "A3流转人数",
			"7日回搜人数": "七日回搜人数",
			"七日回搜人数": "七日回搜人数",
			"七日回搜率": "七日回搜率",
			"完播率": "完播率",
			"完播数": "完播数",
			"回搜次数": "回搜次数",
			"搜索次数": "回搜次数",
			"本次活动转化金额": "本次活动转化金额",
			"转化金额": "本次活动转化金额",
			"拉新人群规模": "拉新人群规模",
			"拉新人数量级": "拉新人群规模",
			"新客占比": "新客占比",
			"CPM": "CPM",
			"CPE": "CPE",
			"CPS(消耗金额/搜索次数)": "CPS",
			"CPS": "CPS",
			"CP流转A3": "CPA3",
			"CPA3": "CPA3",
			"CP5A": "CP5A",
			"新增A3回搜率": "A3流转率",
			"A3流转率": "A3流转率",
			"行业TOP5A3流转率": "行业TOP5_A3流转率",
			"行业TOP5%A3流转率": "行业TOP5_A3流转率",
			"A3行业TOP5%流转均值": "行业TOP5_A3流转率",
			"ROI": "ROI",
		}
		mapped: dict = {}
		for src_key, dst_key in field_map.items():
			val = proj_raw.get(src_key)
			if val is not None and dst_key not in mapped:
				mapped[dst_key] = parse_number(val) if isinstance(val, str) else val
		if mapped:
			project_overview = ProjectOverview(**{k: v for k, v in mapped.items() if v is not None})

	# --- 5A 人群资产 ---
	a5_raw = item.get("5A人群资产") or item.get("a5_asset_flow") or {}
	a5_asset_flow = None
	if a5_raw:
		a5_map = {
			"投后人群规模": "投后人群规模",
			"投后人群增长率": "投后增长率",
			"投后增长率": "投后增长率",
			"拉新人群规模": "拉新人群规模",
			"拉新比例": "拉新比例",
			"拉新比例(行业TOP5%品牌均值)": "拉新比例_行业TOP5均值",
			"A1流转人群": "A1新增量级",
			"A1新增量级": "A1新增量级",
			"A1流转率": "A1流转率",
			"A2流转人群": "A2新增量级",
			"A2新增量级": "A2新增量级",
			"A2流转率": "A2流转率",
			"A3流转人群": "A3新增量级",
			"A3新增量级": "A3新增量级",
			"A3流转率": "A3流转率",
			"A4流转人群": "A4新增量级",
			"A4新增量级": "A4新增量级",
			"A4流转率": "A4流转率",
			"A5流转人群": "A5新增量级",
			"A5新增量级": "A5新增量级",
			"A5流转率": "A5流转率",
		}
		mapped_a5: dict = {}
		for src_key, dst_key in a5_map.items():
			val = a5_raw.get(src_key)
			if val is not None and dst_key not in mapped_a5:
				mapped_a5[dst_key] = parse_number(val) if isinstance(val, str) else val
		if mapped_a5:
			a5_asset_flow = A5AssetFlow(**{k: v for k, v in mapped_a5.items() if v is not None})

	return YuntuDataReport(
		报告名称=report_name,
		品牌名称=item.get("品牌分类") or brand_name,
		日期范围=date_range,
		项目整体=project_overview,
		五A人群资产流转=a5_asset_flow,
		达人及内容复盘=None,
		搜索与溢出价值=None,
		投流数据精细化复盘=None,
		TA人群画像精准度复盘=None,
	)


def _try_parse_json_reports(
	text: str,
	brand_name: str | None = None,
	date_range: str | None = None,
) -> list[YuntuDataReport] | None:
	"""
	尝试从 Agent 输出文本中解析 JSON 并映射到 YuntuDataReport 列表。
	成功返回报告列表，失败返回 None（调用方可回退到正则解析）。
	"""
	parsed = extract_json_from_text(text)
	if parsed is None:
		return None

	# 提示词格式: {"投后报告": [...], "文档下载地址": [...]}
	if isinstance(parsed, dict):
		reports_raw = parsed.get("投后报告")
		if isinstance(reports_raw, list) and reports_raw:
			return [
				_build_report_from_json_item(item, brand_name=brand_name, date_range=date_range)
				for item in reports_raw
				if isinstance(item, dict)
			]
		# 也许 Agent 直接返回了单个报告对象
		if "项目整体" in parsed or "5A人群资产" in parsed or "报告名" in parsed:
			return [_build_report_from_json_item(parsed, brand_name=brand_name, date_range=date_range)]

	if isinstance(parsed, list):
		return [
			_build_report_from_json_item(item, brand_name=brand_name, date_range=date_range)
			for item in parsed
			if isinstance(item, dict)
		] or None

	return None


def _split_text_by_report_names(
	text: str,
	report_names: list[str],
) -> list[tuple[str | None, str]]:
	"""
	按报告名称将文本拆分为多段。查找包含「报告名：X」或「报告名称：X」或【X】的行，用其分割。
	返回 [(report_name, segment_text), ...]，若无法拆分则返回 [(None, text)] 表示整段。
	"""
	report_names = [n for n in report_names if n and n.strip()]
	if not report_names:
		return [(None, text)]
	escaped = [re.escape(n.strip()) for n in report_names]
	name_alt = "|".join(escaped)
	# 匹配：行首/前有换行，然后「报告名：X」或「报告名称：X」或「【X】」，X 为 report_names 之一
	pattern = r"(?:\n|^)(?:(?:报告名|报告名称)[：:]\s*(" + name_alt + r")|【(" + name_alt + r")】)"
	matches = list(re.finditer(pattern, text))
	if not matches:
		return [(None, text)]
	segments: list[tuple[str | None, str]] = []
	for i, m in enumerate(matches):
		report_name = (m.group(1) or m.group(2) or "").strip() if m.lastindex else None
		start = m.end()
		end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
		segment = text[start:end].strip()
		if report_name and segment:
			segments.append((report_name, segment))
	if not segments:
		return [(None, text)]
	return segments


def extract_from_agent_history(
	history: AgentHistoryList,
	report_name: str | None = None,
	brand_name: str | None = None,
	date_range: str | None = None,
	report_names: list[str] | None = None,
) -> YuntuDataReport | list[YuntuDataReport]:
	"""
	从Agent执行结果中提取结构化数据。

	解析策略：
	1. 优先尝试 JSON 解析(Agent 按提示词输出 JSON 时直接映射）。
	2. JSON 解析失败则回退到正则提取。

	若提供 report_names（黑玩多品牌分类：每分类星图+竞价），则尝试按报告名拆分文本并返回多份报告。
	否则返回单份报告。
	"""
	final_result = history.final_result() or ""
	extracted_contents = history.extracted_content()
	all_text = final_result
	if extracted_contents:
		all_text += "\n" + "\n".join(c for c in extracted_contents if c)

	# ---- 策略 1: JSON 解析 ----
	json_reports = _try_parse_json_reports(all_text, brand_name=brand_name, date_range=date_range)
	if json_reports:
		return json_reports if len(json_reports) > 1 else json_reports[0]

	# ---- 策略 2: 正则兜底 ----

	def _build_report(
		seg_name: str | None,
		seg_text: str,
	) -> YuntuDataReport:
		name = seg_name or report_name
		project_overview = _extract_project_overview(seg_text)
		a5_asset_flow = _extract_a5_asset_flow(seg_text)
		search_insight = _extract_search_insight(seg_text)
		ta_portrait = _extract_ta_portrait(seg_text)
		kol_content = _extract_kol_content_info(seg_text)
		ad_flow = _extract_ad_flow_info(seg_text)
		return YuntuDataReport(
			报告名称=name,
			品牌名称=brand_name,
			日期范围=date_range,
			项目整体=project_overview,
			五A人群资产流转=a5_asset_flow,
			达人及内容复盘=kol_content,
			搜索与溢出价值=search_insight,
			投流数据精细化复盘=ad_flow,
			TA人群画像精准度复盘=ta_portrait,
		)

	if report_names:
		segments = _split_text_by_report_names(all_text, report_names)
		if len(segments) > 1 or (len(segments) == 1 and segments[0][0] is not None):
			reports = [
				_build_report(seg_name, seg_text)
				for seg_name, seg_text in segments
			]
			return reports
	# 单报告
	return _build_report(None, all_text)


def _extract_project_overview(text: str) -> ProjectOverview | None:
	"""提取项目整体数据"""
	patterns = {
		'消耗金额': r'消耗金额[：:]\s*([^\n]+)',
		'曝光次数': r'曝光次数[：:]\s*([^\n]+)',
		'曝光人数': r'曝光人数[：:]\s*([^\n]+)',
		'互动率': r'互动率[：:]\s*([^\n]+)',
		'互动量': r'互动量[：:]\s*([^\n]+)',
		'七日回搜人数': r'7日回搜人数[：:]\s*([^\n]+)',
		'七日回搜率': r'七日回搜率[：:]\s*([^\n]+)',
		'回搜次数': r'回搜次数[：:]\s*([^\n]+)',
		'完播率': r'完播率[：:]\s*([^\n]+)',
		'完播数': r'完播数[：:]\s*([^\n]+)',
		'A3流转人数': r'A3流转人数[：:]\s*([^\n]+)',
		'A3流转率': r'A3流转率[：:]\s*([^\n]+)',
		'本次活动转化金额': r'本次活动转化金额[：:]\s*([^\n]+)',
		'拉新人群规模': r'拉新人群规模[：:]\s*([^\n]+)',
		'新客占比': r'新客占比[：:]\s*([^\n]+)',
		'CPM': r'CPM[：:]\s*([^\n]+)',
		'CPE': r'CPE[：:]\s*([^\n]+)',
		'CPS': r'CPS[：:]\s*([^\n]+)',
		'CPA3': r'CPA3[：:]\s*([^\n]+)',
		'CP5A': r'CP5A[：:]\s*([^\n]+)',
		'ROI': r'ROI[：:]\s*([^\n]+)',
	}
	data = extract_fields_by_patterns(text, patterns)
	return ProjectOverview(**data) if data else None


def _extract_a5_asset_flow(text: str) -> A5AssetFlow | None:
	"""提取5A人群资产流转数据"""
	patterns = {
		'投后人群规模': r'投后人群规模[：:]\s*([^\n]+)',
		'投后增长率': r'投后增长率[：:]\s*([^\n]+)',
		'A1新增量级': r'A1新增量级[：:]\s*([^\n]+)',
		'A1流转率': r'A1流转率[：:]\s*([^\n]+)',
		'A2新增量级': r'A2新增量级[：:]\s*([^\n]+)',
		'A2流转率': r'A2流转率[：:]\s*([^\n]+)',
		'A3新增量级': r'A3新增量级[：:]\s*([^\n]+)',
		'A3流转率': r'A3流转率[：:]\s*([^\n]+)',
		'A4新增量级': r'A4新增量级[：:]\s*([^\n]+)',
		'A4流转率': r'A4流转率[：:]\s*([^\n]+)',
		'A5新增量级': r'A5新增量级[：:]\s*([^\n]+)',
		'A5流转率': r'A5流转率[：:]\s*([^\n]+)',
	}
	data = extract_fields_by_patterns(text, patterns)
	return A5AssetFlow(**data) if data else None


def _extract_search_insight(text: str) -> SearchInsight | None:
	"""提取搜索与溢出价值数据"""
	patterns = {
		'SOV': r'SOV[：:]\s*([^\n]+)',
		'搜索人数变化趋势': r'搜索人数变化趋势[：:]\s*([^\n]+)',
		'搜索次数变化趋势': r'搜索次数变化趋势[：:]\s*([^\n]+)',
	}
	data = extract_fields_by_patterns(text, patterns)
	return SearchInsight(**data) if data else None


def _extract_ta_portrait(text: str) -> TAPortraitReview | None:
	"""提取TA人群画像数据"""
	# 尝试提取基础画像和内容偏好信息
	# 这些数据通常是结构化的，需要更复杂的解析
	# 这里先提取文本标记
	
	全触达人群基础画像_match = re.search(r'全触达人群基础画像[：:]\s*([^\n]+(?:\n[^\n]+)*?)(?=全触达人群内容偏好|投后A3人群|$)', text, re.MULTILINE)
	全触达人群内容偏好_match = re.search(r'全触达人群内容偏好[：:]\s*([^\n]+(?:\n[^\n]+)*?)(?=投后A3人群|$)', text, re.MULTILINE)
	投后A3人群基础画像_match = re.search(r'投后A3人群基础画像[：:]\s*([^\n]+(?:\n[^\n]+)*?)(?=投后A3人群内容偏好|$)', text, re.MULTILINE)
	投后A3人群内容偏好_match = re.search(r'投后A3人群内容偏好[：:]\s*([^\n]+)', text, re.MULTILINE)
	
	data = {}
	
	if 全触达人群基础画像_match:
		content = 全触达人群基础画像_match.group(1).strip()
		content = re.sub(r'^\[.*?\]', '', content).strip()
		if content:
			data['全触达人群基础画像'] = {'原始文本': content}
	
	if 全触达人群内容偏好_match:
		content = 全触达人群内容偏好_match.group(1).strip()
		content = re.sub(r'^\[.*?\]', '', content).strip()
		if content:
			data['全触达人群内容偏好'] = {'原始文本': content}
	
	if 投后A3人群基础画像_match:
		content = 投后A3人群基础画像_match.group(1).strip()
		content = re.sub(r'^\[.*?\]', '', content).strip()
		if content:
			data['投后A3人群基础画像'] = {'原始文本': content}
	
	if 投后A3人群内容偏好_match:
		content = 投后A3人群内容偏好_match.group(1).strip()
		content = re.sub(r'^\[.*?\]', '', content).strip()
		if content:
			data['投后A3人群内容偏好'] = {'原始文本': content}
	
	if not data:
		return None
	
	return TAPortraitReview(**data)


def _extract_kol_content_info(text: str) -> KOLContentReview | None:
	"""提取达人及内容复盘信息（主要是Excel文件下载状态）"""
	# 检查是否有Excel文件下载的标记
	excel_patterns = [
		r'已下载[：:]\s*([^\n]+)',
		r'文件下载路径[：:]\s*([^\n]+)',
		r'下载.*?Excel[：:]\s*([^\n]+)',
		r'Excel.*?路径[：:]\s*([^\n]+)',
	]
	
	excel_path = None
	for pattern in excel_patterns:
		match = re.search(pattern, text, re.IGNORECASE)
		if match:
			excel_path = match.group(1).strip()
			excel_path = re.sub(r'^\[.*?\]', '', excel_path).strip()
			if excel_path:
				break
	
	# 检查是否有"已下载"标记
	downloaded = bool(excel_path) or '已下载' in text or '下载完成' in text
	
	return KOLContentReview(
		excel_file_path=excel_path,
		excel_downloaded=downloaded,
		爆文数据=None,
		达人效率对比=None,
		内容数量=None,
		爆文率=None,
		看后搜索率=None,
	)


def _extract_ad_flow_info(text: str) -> AdFlowReview | None:
	"""提取投流数据精细化复盘信息（主要是Excel文件下载状态）"""
	# 使用类似的模式提取Excel文件信息
	excel_patterns = [
		r'已下载[：:]\s*([^\n]+)',
		r'文件下载路径[：:]\s*([^\n]+)',
		r'下载.*?Excel[：:]\s*([^\n]+)',
		r'Excel.*?路径[：:]\s*([^\n]+)',
		r'人群包.*?Excel[：:]\s*([^\n]+)',
	]

	excel_path = None
	for pattern in excel_patterns:
		match = re.search(pattern, text, re.IGNORECASE)
		if match:
			excel_path = match.group(1).strip()
			excel_path = re.sub(r'^\[.*?\]', '', excel_path).strip()
			if excel_path:
				break

	downloaded = bool(excel_path) or '已下载' in text or '下载完成' in text

	# 如果Excel文件存在，尝试读取并提取数据
	extracted_data: dict = {}
	if excel_path and os.path.exists(excel_path):
		try:
			import pandas as pd
			df = pd.read_excel(excel_path)

			if '人群包表现' in df.columns:
				extracted_data['人群包表现'] = df['人群包表现'].to_dict()
			if '投放A3流转率' in df.columns:
				value = df['投放A3流转率'].iloc[0]
				extracted_data['投放A3流转率'] = parse_number(value) if pd.notna(value) else None
			if '各人群包CPA3' in df.columns:
				extracted_data['各人群包CPA3'] = df['各人群包CPA3'].to_dict()
			if '人群精准度' in df.columns:
				value = df['人群精准度'].iloc[0]
				extracted_data['人群精准度'] = parse_number(value) if pd.notna(value) else None
		except Exception as e:
			print(f"Failed to read Excel file: {e}")

	return AdFlowReview(
		excel_file_path=excel_path,
		excel_downloaded=downloaded,
		人群包表现=extracted_data.get('人群包表现'),
		投放A3流转率=extracted_data.get('投放A3流转率'),
		各人群包CPA3=extracted_data.get('各人群包CPA3'),
		人群精准度=extracted_data.get('人群精准度'),
	)
