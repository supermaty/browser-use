"""
数据模型定义 - 用于结构化存储从云图提取的数据
"""
from pydantic import BaseModel, Field, ConfigDict


class ProjectOverview(BaseModel):
	"""项目整体数据"""
	model_config = ConfigDict(extra='forbid')

	# 基础指标
	消耗金额: float | None = Field(None, description="实际消耗金额")
	曝光次数: int | None = Field(None, description="曝光次数")
	曝光人数: int | None = Field(None, description="曝光人数")
	互动率: float | None = Field(None, description="互动率（百分比）")
	互动量: int | None = Field(None, description="互动次数")
	七日回搜人数: int | None = Field(None, description="7日回搜人数")
	七日回搜率: float | None = Field(None, description="七日回搜率（百分比）")
	回搜次数: int | None = Field(None, description="回搜次数")
	完播率: float | None = Field(None, description="完播率（百分比）")
	完播数: int | None = Field(None, description="完播数")
	A3流转人数: int | None = Field(None, description="A3流转人数")
	A3流转率: float | None = Field(None, description="A3流转率（百分比）")
	行业TOP5_A3流转率: float | None = Field(None, description="A3行业TOP5%流转均值（百分比）")
	本次活动转化金额: float | None = Field(None, description="本次活动转化金额")
	拉新人群规模: int | None = Field(None, description="拉新人群规模")
	新客占比: float | None = Field(None, description="新客占比（百分比）")

	# 成本指标
	CPM: float | None = Field(None, description="CPM（每千次曝光成本）")
	CPE: float | None = Field(None, description="CPE（每次互动成本）")
	CPS: float | None = Field(None, description="CPS（每次回搜成本）")
	CPA3: float | None = Field(None, description="CPA3（每个A3流转成本）")
	CP5A: float | None = Field(None, description="CP5A（每个5A人群成本）")
	ROI: float | None = Field(None, description="ROI（投资回报率）")


class A5AssetFlow(BaseModel):
	"""5A人群资产流转数据"""
	model_config = ConfigDict(extra='forbid')

	投后人群规模: int | None = Field(None, description="投后人群规模")
	投后增长率: float | None = Field(None, description="投后人群增长率（百分比）")
	拉新人群规模: int | None = Field(None, description="拉新人群规模")
	拉新比例: float | None = Field(None, description="拉新比例（百分比）")
	拉新比例_行业TOP5均值: float | None = Field(None, description="拉新比例（行业TOP5%品牌均值，百分比）")

	# A1-A5各层级数据
	A1新增量级: int | None = Field(None, description="A1新增量级")
	A1流转率: float | None = Field(None, description="A1流转率（百分比）")
	A2新增量级: int | None = Field(None, description="A2新增量级")
	A2流转率: float | None = Field(None, description="A2流转率（百分比）")
	A3新增量级: int | None = Field(None, description="A3新增量级")
	A3流转率: float | None = Field(None, description="A3流转率（百分比）")
	A4新增量级: int | None = Field(None, description="A4新增量级")
	A4流转率: float | None = Field(None, description="A4流转率（百分比）")
	A5新增量级: int | None = Field(None, description="A5新增量级")
	A5流转率: float | None = Field(None, description="A5流转率（百分比）")


class KOLContentReview(BaseModel):
	"""达人及内容复盘数据（从Excel读取）"""
	model_config = ConfigDict(extra='forbid')

	excel_file_path: str | None = Field(None, description="Excel文件路径")
	excel_downloaded: bool = Field(False, description="是否已下载Excel文件")

	# 从Excel中提取的数据（可选，如果成功读取Excel则填充）
	爆文数据: dict | None = Field(None, description="爆文数据详情")
	达人效率对比: dict | None = Field(None, description="不同层级达人效率对比（CPM/CPE/CPA3）")
	内容数量: int | None = Field(None, description="内容数量")
	爆文率: float | None = Field(None, description="爆文率（百分比）")
	看后搜索率: float | None = Field(None, description="看后搜索率（百分比）")


class SearchInsight(BaseModel):
	"""搜索与溢出价值数据"""
	model_config = ConfigDict(extra='forbid')

	SOV: float | None = Field(None, description="SOV（声量份额，指数）")
	搜索人数变化趋势: float | None = Field(None, description="搜索人数变化环比（百分比）")
	搜索次数变化趋势: float | None = Field(None, description="搜索次数变化环比（百分比）")


class AdFlowReview(BaseModel):
	"""投流数据精细化复盘数据（从Excel读取）"""
	model_config = ConfigDict(extra='forbid')

	excel_file_path: str | None = Field(None, description="Excel文件路径")
	excel_downloaded: bool = Field(False, description="是否已下载Excel文件")

	# 从Excel中提取的数据（可选，如果成功读取Excel则填充）
	人群包表现: dict | None = Field(None, description="人群包表现详情")
	投放A3流转率: float | None = Field(None, description="投放A3流转率（百分比）")
	各人群包CPA3: dict | None = Field(None, description="各人群包CPA3数据")
	人群精准度: float | None = Field(None, description="人群精准度")


class TAPortraitReview(BaseModel):
	"""TA人群画像精准度复盘数据"""
	model_config = ConfigDict(extra='forbid')

	# 全触达人群画像
	全触达人群基础画像: dict | None = Field(None, description="全触达人群基础画像（八大人群、年龄、性别、城市线级等）")
	全触达人群内容偏好: dict | None = Field(None, description="全触达人群内容偏好")

	# 投后A3人群画像
	投后A3人群基础画像: dict | None = Field(None, description="投后A3人群基础画像（八大人群、年龄、性别、城市线级等）")
	投后A3人群内容偏好: dict | None = Field(None, description="投后A3人群内容偏好")


class YuntuDataReport(BaseModel):
	"""整合所有数据的根模型"""
	model_config = ConfigDict(extra='forbid')

	报告名称: str | None = Field(None, description="报告名称")
	品牌名称: str | None = Field(None, description="品牌名称")
	日期范围: str | None = Field(None, description="日期范围")

	项目整体: ProjectOverview | None = Field(None, description="项目整体数据")
	五A人群资产流转: A5AssetFlow | None = Field(None, description="5A人群资产流转数据")
	达人及内容复盘: KOLContentReview | None = Field(None, description="达人及内容复盘数据")
	搜索与溢出价值: SearchInsight | None = Field(None, description="搜索与溢出价值数据")
	投流数据精细化复盘: AdFlowReview | None = Field(None, description="投流数据精细化复盘数据")
	TA人群画像精准度复盘: TAPortraitReview | None = Field(None, description="TA人群画像精准度复盘数据")
