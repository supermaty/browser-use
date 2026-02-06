"""
工具函数 - 辅助数据处理和文件操作
"""
import re
from pathlib import Path
from typing import Optional
from datetime import datetime, timedelta


def find_downloaded_excel_files(
	downloads_path: str | Path,
	max_age_hours: int = 24,
	file_patterns: list[str] | None = None
) -> list[Path]:
	"""
	在下载目录中查找最近下载的Excel文件
	
	Args:
		downloads_path: 下载目录路径
		max_age_hours: 最大文件年龄（小时），默认24小时
		file_patterns: 文件名模式列表，用于匹配特定文件（如包含"爆文"、"人群包"等）
		
	Returns:
		按修改时间排序的Excel文件路径列表（最新的在前）
	"""
	downloads_dir = Path(downloads_path).expanduser()
	if not downloads_dir.exists():
		return []
	
	# 默认文件模式
	if file_patterns is None:
		file_patterns = []
	
	# 查找Excel文件
	excel_files = []
	cutoff_time = datetime.now() - timedelta(hours=max_age_hours)
	
	for file_path in downloads_dir.glob("*.xlsx"):
		if file_path.is_file():
			# 检查文件修改时间
			mtime = datetime.fromtimestamp(file_path.stat().st_mtime)
			if mtime >= cutoff_time:
				# 如果指定了文件模式，检查文件名是否匹配
				if file_patterns:
					file_name = file_path.name.lower()
					if any(pattern.lower() in file_name for pattern in file_patterns):
						excel_files.append((file_path, mtime))
					else:
						excel_files.append((file_path, mtime))
				else:
					excel_files.append((file_path, mtime))
	
	# 也查找.xls文件（旧格式）
	for file_path in downloads_dir.glob("*.xls"):
		if file_path.is_file():
			mtime = datetime.fromtimestamp(file_path.stat().st_mtime)
			if mtime >= cutoff_time:
				if file_patterns:
					file_name = file_path.name.lower()
					if any(pattern.lower() in file_name for pattern in file_patterns):
						excel_files.append((file_path, mtime))
				else:
					excel_files.append((file_path, mtime))
	
	# 按修改时间排序（最新的在前）
	excel_files.sort(key=lambda x: x[1], reverse=True)
	return [file_path for file_path, _ in excel_files]


def parse_number(value: str | int | float | None) -> float | int | None:
	"""
	安全的数字解析，处理百分比、千分位等格式
	
	Args:
		value: 要解析的值（字符串、数字或None）
		
	Returns:
		解析后的数字（int或float），如果无法解析则返回None
	"""
	if value is None:
		return None
	
	if isinstance(value, (int, float)):
		return value
	
	if not isinstance(value, str):
		return None
	
	# 移除空白字符
	value = value.strip()
	if not value or value.lower() in ['n/a', 'na', '-', '--', 'null', 'none']:
		return None
	
	# 处理百分比
	if '%' in value:
		value = value.replace('%', '').strip()
		try:
			num = float(value)
			return num / 100.0  # 转换为小数
		except ValueError:
			return None
	
	# 移除千分位分隔符（逗号）
	value = value.replace(',', '').replace('，', '')
	
	# 移除其他常见符号
	value = re.sub(r'[^\d.\-+]', '', value)
	
	try:
		# 尝试解析为整数
		if '.' not in value:
			return int(value)
		# 解析为浮点数
		return float(value)
	except ValueError:
		return None


def format_percentage(value: float | int | None, decimals: int = 2) -> str:
	"""
	格式化百分比显示
	
	Args:
		value: 要格式化的值（小数形式，如0.15表示15%）
		decimals: 小数位数
		
	Returns:
		格式化后的百分比字符串（如"15.00%"）
	"""
	if value is None:
		return "N/A"
	
	try:
		percentage = float(value) * 100
		return f"{percentage:.{decimals}f}%"
	except (ValueError, TypeError):
		return "N/A"


def parse_chinese_number(value: str | None) -> float | int | None:
	"""
	解析中文数字格式（如"1.5万"、"1000"等）
	
	Args:
		value: 包含中文数字的字符串
		
	Returns:
		解析后的数字
	"""
	if value is None or not isinstance(value, str):
		return None
	
	value = value.strip()
	if not value:
		return None
	
	# 处理中文单位
	multipliers = {
		'万': 10000,
		'萬': 10000,
		'亿': 100000000,
		'億': 100000000,
		'千': 1000,
		'百': 100,
	}
	
	for unit, multiplier in multipliers.items():
		if unit in value:
			# 提取数字部分
			num_str = re.sub(r'[^\d.]', '', value)
			try:
				num = float(num_str)
				return int(num * multiplier) if num * multiplier == int(num * multiplier) else num * multiplier
			except ValueError:
				return None
	
	# 如果没有单位，直接解析
	return parse_number(value)


def identify_excel_file_type(file_path: Path) -> str | None:
	"""
	根据文件名识别Excel文件类型
	
	Args:
		file_path: Excel文件路径
		
	Returns:
		文件类型：'kol_content'（达人内容）、'ad_flow'（投流数据）或None
	"""
	file_name = file_path.name.lower()
	
	# 达人及内容复盘相关关键词
	kol_keywords = ['爆文', '达人', 'kol', '内容', '视频', '加热', '星图']
	if any(keyword in file_name for keyword in kol_keywords):
		return 'kol_content'
	
	# 投流数据精细化复盘相关关键词
	ad_keywords = ['人群包', '投流', '投放', '触达', '人群画像']
	if any(keyword in file_name for keyword in ad_keywords):
		return 'ad_flow'
	
	return None
