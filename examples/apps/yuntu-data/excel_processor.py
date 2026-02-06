"""
Excel处理模块 - 读取下载的Excel文件并计算指标
"""
import re
import pandas as pd
from pathlib import Path
from typing import Optional, Any

from data_models import KOLContentReview, AdFlowReview
from utils import find_downloaded_excel_files, identify_excel_file_type, parse_number


def _normalize_date_to_yyyymmdd(value: str) -> int:
	"""将日期字符串转为 YYYYMMDD 整数，便于与 Excel 中的日期列比较。支持 2025-12-01、2025/12/01、20251201。"""
	value = (value or "").strip()
	# 已是 8 位数字
	m = re.match(r"^(\d{4})[-/]?(\d{2})[-/]?(\d{2})$", value)
	if m:
		return int(m.group(1) + m.group(2) + m.group(3))
	if re.match(r"^\d{8}$", value):
		return int(value)
	raise ValueError(f"无法解析日期: {value!r}，请使用 YYYY-MM-DD 或 YYYYMMDD 格式")


def sum_search_count_from_excel(
	file_path: str | Path,
	date_start: str,
	date_end: str,
) -> int:
	"""
	从 Excel 文件中按计算时间区间累加「搜索次数」并返回。

	Excel 格式（参考搜索趋势分析导出）：列名为「日期」「搜索人数」「搜索次数」；
	日期列为 YYYYMMDD 数字（如 20251124）。

	Args:
		file_path: Excel 文件路径（下载后的文件地址）
		date_start: 计算区间起始日期，支持 "2025-12-01" 或 "20251201"
		date_end: 计算区间结束日期，同上

	Returns:
		区间内「搜索次数」的合计值
	"""
	path = Path(file_path)
	if not path.exists():
		raise FileNotFoundError(f"Excel 文件不存在: {path}")
	start_int = _normalize_date_to_yyyymmdd(date_start)
	end_int = _normalize_date_to_yyyymmdd(date_end)
	if start_int > end_int:
		raise ValueError(f"起始日期 {date_start} 不能晚于结束日期 {date_end}")

	# 读取第一个 sheet（搜索趋势通常在第一页）
	df = pd.read_excel(path, sheet_name=0)
	# 兼容列名：日期 / 搜索次数（允许前后空格）
	date_col = None
	search_count_col = None
	for c in df.columns:
		cn = str(c).strip()
		if cn == "日期":
			date_col = c
		if cn == "搜索次数":
			search_count_col = c
	if date_col is None or search_count_col is None:
		raise ValueError(
			f"Excel 中未找到「日期」或「搜索次数」列。当前列名: {list(df.columns)}"
		)

	# 日期列可能是 int 或 float（如 20251124.0）
	dates = pd.to_numeric(df[date_col], errors="coerce").fillna(0).astype(int)
	mask = (dates >= start_int) & (dates <= end_int)
	subset = df.loc[mask, search_count_col]
	# 搜索次数为整数，兼容空单元格
	total = pd.to_numeric(subset, errors="coerce").fillna(0).astype(int).sum()
	return int(total)


def sum_search_count_from_downloads(
	downloads_path: str | Path,
	date_start: str,
	date_end: str,
	file_path: str | Path | None = None,
	max_age_hours: int = 24,
) -> int:
	"""
	从下载目录或指定路径读取 Excel，按计算时间区间累加「搜索次数」并返回。

	若未指定 file_path，则在 downloads_path 下查找最近下载的 Excel（.xlsx/.xls），
	逐个尝试直到找到包含「日期」「搜索次数」列的文件并成功累加。

	Args:
		downloads_path: 下载目录路径
		date_start: 计算区间起始日期（如 "2025-12-01" 或 "20251201"）
		date_end: 计算区间结束日期
		file_path: 若提供则直接使用该文件，否则从 downloads_path 中查找
		max_age_hours: 未指定 file_path 时，只考虑最近 max_age_hours 小时内修改的文件

	Returns:
		区间内「搜索次数」的合计值
	"""
	if file_path is not None:
		return sum_search_count_from_excel(file_path, date_start, date_end)
	files = find_downloaded_excel_files(downloads_path, max_age_hours=max_age_hours)
	last_error: Exception | None = None
	for path in files:
		try:
			return sum_search_count_from_excel(path, date_start, date_end)
		except (ValueError, FileNotFoundError) as e:
			last_error = e
			continue
	if last_error is not None:
		raise last_error
	raise FileNotFoundError(
		f"在下载目录中未找到符合条件的 Excel 文件（需包含「日期」「搜索次数」列）: {downloads_path}"
	)


def read_kol_content_excel(file_path: Path | str) -> dict[str, Any]:
	"""
	读取达人及内容复盘Excel文件
	
	Args:
		file_path: Excel文件路径
		
	Returns:
		提取的数据字典
	"""
	try:
		df = pd.read_excel(file_path, sheet_name=None)  # 读取所有sheet
		
		result = {
			'文件路径': str(file_path),
			'sheets': {},
			'爆文数据': {},
			'达人效率对比': {},
		}
		
		# 处理每个sheet
		for sheet_name, sheet_df in df.items():
			result['sheets'][sheet_name] = {
				'行数': len(sheet_df),
				'列数': len(sheet_df.columns),
				'列名': list(sheet_df.columns),
			}
			
			# 尝试识别关键列并提取数据
			# 根据实际Excel结构调整
			if '曝光' in sheet_name or '曝光' in str(sheet_df.columns):
				# 提取曝光相关数据
				exposure_cols = [col for col in sheet_df.columns if '曝光' in str(col)]
				if exposure_cols:
					result['爆文数据']['曝光数据'] = sheet_df[exposure_cols].to_dict('records')
			
			if '互动' in sheet_name or '互动' in str(sheet_df.columns):
				# 提取互动相关数据
				interaction_cols = [col for col in sheet_df.columns if '互动' in str(col)]
				if interaction_cols:
					result['爆文数据']['互动数据'] = sheet_df[interaction_cols].to_dict('records')
			
			# 查找CPM、CPE、CPA3相关列
			cost_cols = [col for col in sheet_df.columns if any(keyword in str(col) for keyword in ['CPM', 'CPE', 'CPA3', '成本'])]
			if cost_cols:
				result['达人效率对比'] = sheet_df[cost_cols].to_dict('records')
		
		# 计算总体指标
		if df:
			# 尝试从第一个sheet计算总体数据
			first_sheet = list(df.values())[0]
			result['内容数量'] = len(first_sheet)
			
			# 尝试计算爆文率（如果有相关列）
			for col in first_sheet.columns:
				if '爆文' in str(col) or '完播' in str(col):
					# 假设有完播率列，计算爆文率
					completion_col = [c for c in first_sheet.columns if '完播率' in str(c) or '完播' in str(c)]
					if completion_col:
						completion_rates = first_sheet[completion_col[0]].apply(parse_number)
						valid_rates = completion_rates.dropna()
						if len(valid_rates) > 0:
							# 假设完播率>50%为爆文
							explosive_count = len(valid_rates[valid_rates > 0.5])
							result['爆文率'] = explosive_count / len(valid_rates) if len(valid_rates) > 0 else None
		
		return result
		
	except Exception as e:
		return {
			'错误': str(e),
			'文件路径': str(file_path),
		}


def read_ad_flow_excel(file_path: Path | str) -> dict[str, Any]:
	"""
	读取投流数据精细化复盘Excel文件
	
	Args:
		file_path: Excel文件路径
		
	Returns:
		提取的数据字典
	"""
	try:
		df = pd.read_excel(file_path, sheet_name=None)  # 读取所有sheet
		
		result = {
			'文件路径': str(file_path),
			'sheets': {},
			'人群包表现': {},
			'各人群包CPA3': {},
		}
		
		# 处理每个sheet
		for sheet_name, sheet_df in df.items():
			result['sheets'][sheet_name] = {
				'行数': len(sheet_df),
				'列数': len(sheet_df.columns),
				'列名': list(sheet_df.columns),
			}
			
			# 尝试识别人群包相关列
			crowd_cols = [col for col in sheet_df.columns if any(keyword in str(col) for keyword in ['人群', '包', '人群包'])]
			if crowd_cols:
				result['人群包表现'][sheet_name] = sheet_df[crowd_cols].to_dict('records')
			
			# 查找A3流转率相关列
			a3_cols = [col for col in sheet_df.columns if 'A3' in str(col) or '流转' in str(col)]
			if a3_cols:
				# 计算平均A3流转率
				for col in a3_cols:
					if '率' in str(col) or '流转率' in str(col):
						rates = sheet_df[col].apply(parse_number)
						valid_rates = rates.dropna()
						if len(valid_rates) > 0:
							result['投放A3流转率'] = valid_rates.mean()
							break
			
			# 查找CPA3相关列
			cpa3_cols = [col for col in sheet_df.columns if 'CPA3' in str(col) or 'CPA' in str(col)]
			if cpa3_cols:
				result['各人群包CPA3'][sheet_name] = sheet_df[cpa3_cols].to_dict('records')
		
		return result
		
	except Exception as e:
		return {
			'错误': str(e),
			'文件路径': str(file_path),
		}


def process_downloaded_excel_files(
	downloads_path: str | Path,
	report: Any  # YuntuDataReport类型，但避免循环导入
) -> tuple[Optional[KOLContentReview], Optional[AdFlowReview]]:
	"""
	处理下载的Excel文件，更新报告中的达人内容和投流数据
	
	Args:
		downloads_path: 下载目录路径
		report: 数据报告对象（将被更新）
		
	Returns:
		更新后的KOLContentReview和AdFlowReview对象
	"""
	# 查找Excel文件
	excel_files = find_downloaded_excel_files(downloads_path, max_age_hours=24)
	
	kol_content_review = None
	ad_flow_review = None
	
	for excel_file in excel_files:
		file_type = identify_excel_file_type(excel_file)
		
		if file_type == 'kol_content':
			# 处理达人内容Excel
			data = read_kol_content_excel(excel_file)
			if '错误' not in data:
				kol_content_review = KOLContentReview(
					excel_file_path=str(excel_file),
					excel_downloaded=True,
					爆文数据=data.get('爆文数据'),
					达人效率对比=data.get('达人效率对比'),
					内容数量=data.get('内容数量'),
					爆文率=data.get('爆文率'),
					看后搜索率=None,  # 需要根据实际Excel结构调整
				)
		
		elif file_type == 'ad_flow':
			# 处理投流数据Excel
			data = read_ad_flow_excel(excel_file)
			if '错误' not in data:
				ad_flow_review = AdFlowReview(
					excel_file_path=str(excel_file),
					excel_downloaded=True,
					人群包表现=data.get('人群包表现'),
					投放A3流转率=data.get('投放A3流转率'),
					各人群包CPA3=data.get('各人群包CPA3'),
					人群精准度=None,  # 需要根据实际Excel结构调整
				)
	
	return kol_content_review, ad_flow_review


def calculate_metrics(project_overview: Any) -> dict[str, float | None]:
	"""
	根据项目整体数据计算成本指标
	
	Args:
		project_overview: ProjectOverview对象
		
	Returns:
		计算后的指标字典
	"""
	if not project_overview:
		return {}
	
	metrics = {}
	
	# CPM = 实际消耗金额 / 曝光次数 * 1000
	if project_overview.消耗金额 and project_overview.曝光次数 and project_overview.曝光次数 > 0:
		metrics['CPM'] = (project_overview.消耗金额 / project_overview.曝光次数) * 1000
	else:
		metrics['CPM'] = project_overview.CPM
	
	# CPE = 实际消耗金额 / 互动量
	if project_overview.消耗金额 and project_overview.互动量 and project_overview.互动量 > 0:
		metrics['CPE'] = project_overview.消耗金额 / project_overview.互动量
	else:
		metrics['CPE'] = project_overview.CPE
	
	# CPS = 实际消耗金额 / 回搜次数
	if project_overview.消耗金额 and project_overview.回搜次数 and project_overview.回搜次数 > 0:
		metrics['CPS'] = project_overview.消耗金额 / project_overview.回搜次数
	else:
		metrics['CPS'] = project_overview.CPS
	
	# CPA3 = 实际消耗金额 / A3流转人数
	if project_overview.消耗金额 and project_overview.A3流转人数 and project_overview.A3流转人数 > 0:
		metrics['CPA3'] = project_overview.消耗金额 / project_overview.A3流转人数
	else:
		metrics['CPA3'] = project_overview.CPA3
	
	# ROI = 本次活动转化金额 / 实际消耗金额
	if project_overview.本次活动转化金额 and project_overview.消耗金额 and project_overview.消耗金额 > 0:
		metrics['ROI'] = project_overview.本次活动转化金额 / project_overview.消耗金额
	else:
		metrics['ROI'] = project_overview.ROI
	
	return metrics
