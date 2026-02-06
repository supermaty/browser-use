"""
数据提取模块 - 从Agent执行结果中提取结构化数据
"""
import re
from typing import Optional

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
	report_name: Optional[str] = None,
	brand_name: Optional[str] = None,
	date_range: Optional[str] = None,
	report_names: Optional[list[str]] = None,
) -> YuntuDataReport | list[YuntuDataReport]:
	"""
	从Agent执行结果中提取结构化数据。
	若提供 report_names（黑玩多品牌分类：每分类星图+竞价），则尝试按报告名拆分文本并返回多份报告。
	否则返回单份报告。
	"""
	final_result = history.final_result() or ""
	extracted_contents = history.extracted_content()
	all_text = final_result
	if extracted_contents:
		all_text += "\n" + "\n".join(extracted_contents)

	def _build_report(
		seg_name: Optional[str],
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


def _extract_project_overview(text: str) -> Optional[ProjectOverview]:
	"""提取项目整体数据"""
	# 使用正则表达式提取各个字段
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
	
	data = {}
	for key, pattern in patterns.items():
		match = re.search(pattern, text)
		if match:
			value_str = match.group(1).strip()
			# 移除可能的方括号标记
			value_str = re.sub(r'^\[.*?\]', '', value_str).strip()
			if value_str and value_str.lower() not in ['n/a', 'na', '-', '--', 'null', 'none']:
				data[key] = parse_number(value_str) if key in ['互动率', '七日回搜率', '完播率', 'A3流转率', '新客占比'] else parse_number(value_str)
	
	# 如果没有提取到任何数据，返回None
	if not data:
		return None
	
	return ProjectOverview(**data)


def _extract_a5_asset_flow(text: str) -> Optional[A5AssetFlow]:
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
	
	data = {}
	for key, pattern in patterns.items():
		match = re.search(pattern, text)
		if match:
			value_str = match.group(1).strip()
			value_str = re.sub(r'^\[.*?\]', '', value_str).strip()
			if value_str and value_str.lower() not in ['n/a', 'na', '-', '--', 'null', 'none']:
				data[key] = parse_number(value_str) if '流转率' in key or '增长率' in key else parse_number(value_str)
	
	if not data:
		return None
	
	return A5AssetFlow(**data)


def _extract_search_insight(text: str) -> Optional[SearchInsight]:
	"""提取搜索与溢出价值数据"""
	patterns = {
		'SOV': r'SOV[：:]\s*([^\n]+)',
		'搜索人数变化趋势': r'搜索人数变化趋势[：:]\s*([^\n]+)',
		'搜索次数变化趋势': r'搜索次数变化趋势[：:]\s*([^\n]+)',
	}
	
	data = {}
	for key, pattern in patterns.items():
		match = re.search(pattern, text)
		if match:
			value_str = match.group(1).strip()
			value_str = re.sub(r'^\[.*?\]', '', value_str).strip()
			if value_str and value_str.lower() not in ['n/a', 'na', '-', '--', 'null', 'none']:
				data[key] = parse_number(value_str)
	
	if not data:
		return None
	
	return SearchInsight(**data)


def _extract_ta_portrait(text: str) -> Optional[TAPortraitReview]:
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


def _extract_kol_content_info(text: str) -> Optional[KOLContentReview]:
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
	)


def _extract_ad_flow_info(text: str) -> Optional[AdFlowReview]:
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
	
	return AdFlowReview(
		excel_file_path=excel_path,
		excel_downloaded=downloaded,
	)
