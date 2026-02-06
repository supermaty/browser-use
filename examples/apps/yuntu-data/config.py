import os
from pathlib import Path
from dotenv import load_dotenv


load_dotenv()

API_KEY = os.getenv('GOOGLE_API_KEY')
if not API_KEY:
	raise ValueError('GOOGLE_API_KEY is not set')

OUTPUT_BASE_DIR = './data-output'

# Excel输出配置
EXCEL_OUTPUT_DIR = Path(OUTPUT_BASE_DIR) / 'excel_exports'
EXCEL_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# 每次任务使用独立子目录：下载与导出均写入该目录，任务间隔离
TASKS_BASE_DIR = Path(OUTPUT_BASE_DIR) / 'tasks'
TASKS_BASE_DIR.mkdir(parents=True, exist_ok=True)

# Excel文件命名规则
def generate_excel_filename(brand_name: str | None = None, report_name: str | None = None) -> str:
	"""生成Excel文件名"""
	from datetime import datetime
	timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
	brand = brand_name or "未知品牌"
	report = report_name or "未知报告"
	# 清理文件名中的非法字符
	import re
	brand = re.sub(r'[<>:"/\\|?*]', '_', brand)
	report = re.sub(r'[<>:"/\\|?*]', '_', report)
	return f"{brand}_{report}_{timestamp}.xlsx"

# Project Root and Profile Path
# Assuming this file is in examples/apps/yuntu-data/
PROJECT_ROOT = Path(__file__).resolve().parents[3]

# 如果 .env 中配置了绝对路径, 直接使用；否则使用项目根目录下的相对路径
PROFILE_DIR_STR = os.getenv('BROWSER_PROFILE_DIR')
if PROFILE_DIR_STR:
    # 如果是绝对路径, 直接使用
    if os.path.isabs(PROFILE_DIR_STR):
        PROFILE_DIR = Path(PROFILE_DIR_STR)
    else:
        # 相对路径, 相对于项目根目录
        PROFILE_DIR = PROJECT_ROOT / PROFILE_DIR_STR
else:
    # 默认路径
    PROFILE_DIR = PROJECT_ROOT / 'browser_profiles' / 'yuntu_profile'
# Ensure profile directory exists

STORAGE_STATE_FILE = PROFILE_DIR / 'storage_state.json'
PROFILE_DIR.mkdir(parents=True, exist_ok=True)



# 从 .env 读取, 如果没有则使用默认值
DOWNLOADS_PATH = os.getenv(
    'BROWSER_DOWNLOADS_PATH',
    str(PROJECT_ROOT / 'downloads')
)
Path(DOWNLOADS_PATH).expanduser().mkdir(parents=True, exist_ok=True)

# Models
GEMINI_MODEL = "gemini-3-pro-preview"

# URLs
YUNTU_LOGIN_URL = "https://yuntu.oceanengine.com/account/login" # Adjust if necessary, main entry point

# 品牌名 -> 英文名（第一步右上角切换品牌时，界面可能显示英文名，便于 Agent 定位）
BRAND_NAME_MAPPING: dict[str, str] = {
	"黑玩": "Heyone",
	"施华蔻": "Schwarzkopf",
}


def get_brand_english_name(brand_name: str | None) -> str:
	"""返回品牌对应的英文名；未在 mapping 中则返回原品牌名。"""
	if not brand_name or not brand_name.strip():
		return "未指定"
	return BRAND_NAME_MAPPING.get(brand_name.strip(), brand_name.strip())

# Q0: 互动率在哪里？
# Q1: 竞价投放的转化金额怎么获取？
# Prompts
HEYONE_PROMPT_TEMPLATE = """
    当前使用案例在巨量云图(https://yuntu.oceanengine.com)复盘的SOP。
    【必须】完成下方品牌分类列表中每一个报告的取数（星图+竞价均需完成），并完整执行第四、五、六步；不得只完成一个报告或跳过步骤4/5/6就结束。

    打开巨量云图网站后, 请在右上角用户信息处切换到用户指定的品牌（{brand_name}；若中文名字查不到则查找英文名 {brand_name_en}）, 然后按照以下取数路径获取核心指标（百分比全部使用带%符号的数值）并以json的格式输出。

    品牌分类与报告名（本任务需对每个品牌分类分别取星图报告与竞价报告；若列表为空则为单报告模式）:
    {brand_categories_formatted}

    将以下取数步骤重复执行第一到第四步，直到所有品牌分类的星图报告与竞价报告都取数完成（品牌分类列表为空则为单报告模式），再执行第五步到第六步：

    一、项目整体 Overview
        汇总全案数据(KOL+KOC+投放)主要去看整体的表现以及KPI完成度。
        [目标数据]: 曝光次数、曝光人数、互动率、7日回搜人数、七日回搜率、完播率、完播数、A3 流转人数、A3 流转率、A3行业TOP5%流转均值、本次活动转化金额、拉新人群规模、
        拉新比例、CPM、CPE、CP流转A3、CP5A、CPS、新增A3回搜率、ROI
        [数据计算说明]: 
        - 互动次数 = 互动率 * 曝光次数
        - 回搜次数: 需要从品牌形象-搜索分析-搜索趋势分析中选取对应天数, 将每天的搜索次数相加得到总回搜次数
        - 完播数 = 曝光次数 * 完播率
        - 成本指标(CP开头): 消耗金额（星图、竞价）由用户必填提供；本次活动整体为星图金额+投放金额、竞价为竞价实际金额
        - CPM = 实际消耗金额 / 曝光次数 * 1000
        - CPE = 实际消耗金额 / 互动次数
        - CPS = 实际消耗金额 / 回搜次数
        - CPA3 = 实际消耗金额 / A3流转人数
        - ROI = 本次活动转化金额 / 实际消耗金额
        注意: 目标数据分两份, 分别为星图(达人营销)和竞价(竞价投放), 需分别获取。

        [详细取数路径]: 
        1. 进入巨量云图首页, 在右上角切换到用户指定的品牌（{brand_name}；中文名查不到则选 {brand_name_en}）
        2. 点击"营销决策"菜单
        3. 点击"投后结案"
        4. 在升级版报告列表中搜索并找到报告名为"{report_name}"的报告
        5. 点击"查看报告"进入报告详情页
        6. 在报告详情页中, 提取计算时间区间(sample:  2025.12.01 ~ 2025.12.31)用于后续数据筛选
        7. 在报告详情页中, 依次查看以下模块并提取两份数据: 
            - 品牌形象: 
                - 搜索分析 - 搜索趋势分析: 点击右侧下载按钮下载文件, 然后调用回搜次数Tool计算回搜次数
            - 流量分析: 
                - 触点: 全部/竞价投放: 提取曝光次数、曝光人数、完播率、7日回搜人数、七日回搜率
                - 触点: 全部/达人营销: 提取曝光次数、曝光人数、完播率、7日回搜人数、七日回搜率
            - 人群分析: 
                - 触点: 全部/达人营销: 提取A3流转人数、A3流转率、A3行业TOP5%流转均值(需鼠标挪到问号上查看)、拉新人群规模、拉新比例
                - 触点: 全部/竞价投放: 提取A3流转人数、A3流转率、A3行业TOP5%流转均值(需鼠标挪到问号上查看)、拉新人群规模、拉新比例
            - 转化分析: 提取本次活动转化金额 (此数值为达人营销的转化金额)
        8. 如果某些数据在页面上直接显示, 请直接提取；如果数据需要计算, 请按照上述计算公式进行计算
    二、5A人群资产流转
        评估营销对品牌资产的贡献度。
        
        [目标数据]: 投后人群规模、投后人群增长率、A1-A5人群增长人数、A1-A5人群流转率。
        
        [详细取数路径]:
        1. 在报告详情页(报告名: {report_name})中, 点击左侧菜单"人群分析"
        2. 在人群分析页面中, 找到并点击"5A关系资产分析"
        3. 在5A关系资产分析页面中, 点击"5A人群资产"标签
        4. 在触点选择中, 选择"全部"(如果有多级选择, 都选择"全部")
        5. 等待数据加载完成后, 提取以下数据: 
           - 拉新人群规模
           - 拉新比例
           - 拉新比例(行业TOP5%品牌均值) (需鼠标挪到附近的问号上从后弹框内容查找)
           - A1-A5各层级数据:
             * A1流转人群
             * A1流转率
             * A1流转率(行业TOP5%品牌均值) (需鼠标挪到附近的问号上从后弹框内容查找)
             * A2流转人群
             * A2流转率
             * A2流转率(行业TOP5%品牌均值) (需鼠标挪到附近的问号上从后弹框内容查找)
             * A3流转人群
             * A3流转率
             * A3流转率(行业TOP5%品牌均值) (需鼠标挪到附近的问号上从后弹框内容查找)
             * A4流转人群
             * A4流转率
             * A4流转率(行业TOP5%品牌均值) (需鼠标挪到附近的问号上从后弹框内容查找)
             * A5流转人群
             * A5流转率
             * A5流转率(行业TOP5%品牌均值) (需鼠标挪到附近的问号上从后弹框内容查找)
    三、投放人群画像
        看本次营销整体、投放的曝光人群、A3人群画像是否精准。
        [数据维度]：八大人群&年龄&性别&城市线级等基础数据
        [分析建议]: 本次投放是否精准,曝光TA浓度
        [取数路径]:
        投放曝光人群：
        第二步执行完毕后，保持在人群分析页面, 
        选择 画像分析-5A人群资产: 投后-确保"5A人群总资产"卡片被选中-根据业务选择触点({report_name}包含"竞价"，则选择 "全部/竞价投放"；报告名包含"星图"，则选择 "全部/达人营销")-点击下载按钮
        投后A3人群: 
        上一部执行完毕后，保持在当前页面
        选择 "A3问询"卡片-点击下载按钮(记录文件下载地址并添加在输出json中)
    四、投放人群复盘
        拆分投放效果。

        [目标数据]: 人群包表现、投放A3 流转率、各人群包 CPA3、渗透率...
        [分析建议]: 为下次投放提供人群包优化迭代建议。
        [详细取数路径]: 
        1. 在报告详情页(报告名: {report_name})中, 点击左侧菜单"人群分析"
        2. 在"人群分析"页中, 找到并点击"人群画像分析"
        3. 在 找到"人群包投放效果分析"区域, 确保触点为"全部-全部", 点击下载按钮
        4. 下载完成后, 请确认文件已成功下载, 然后调用处理人群包效果分析的Tool(传递对应的参数)

    五、达人及内容复盘(Content & KOL)
        基于个体表现去劣存优。
        
        [分析重点]: 爆文数据特征、不同层级达人效率对比(CPM/CPE/CPA3)。
        [视觉内容]: 星图曝光人群画像与品牌 TA 匹配度分析。
        
        [详细取数路径]:
        1. 在巨量云图首页, 点击左侧菜单"营销触点"
        2. 在营销触点页面中, 点击"营销概览"
        3. 在营销概览页面中, 找到并点击"爆文加热"模块
        4. 在爆文加热页面中, 找到"视频发布日期"筛选条件
        5. 根据用户输入的爆文加热时间：类型 {kol_content_date_type}，时间范围 {kol_content_date_range}；选择时间类型和时间范围(不要手动输入日期, 使用下拉菜单选择) ；
        若出现对比周期，确保对比周期时间范围是{kol_content_date_range}前一天开始的上一周期
        6. 等待数据加载后, 找到"视频榜单"标签
        7. 在视频榜单页面中, 找到"导出数据"或"下载"按钮(通常在页面右上角或表格上方)
        8. 点击导出按钮后, 会弹出下载选项对话框
        9. 在下载选项中, 选择"下载类型: 全部内容"(确保选择全部内容, 而不是部分内容)
        10. 确认下载后, 等待Excel文件下载完成
        11. 下载完成后, 请确认文件已成功下载, 并在输出中说明: "达人及内容复盘Excel文件已下载, 文件路径: [文件路径]"
        12. 注意: 如果下载按钮不可见或无法点击, 请先滚动页面查看, 或检查是否有权限限制
    六、搜索与溢出价值
        复盘"看搜"链路。
        [数据维度]: SOV(声量份额)、搜索人数/次数变化趋势(活动期 vs 活动前eg: 活动期为1.20-1.29, 活动前则为1.10-1.19)
        
        [详细取数路径]: 
        1. 在巨量云图首页, 点击左侧菜单"营销触点"
        2. 在营销触点页面中, 找到并点击"行业搜索洞察"
        3. 在行业搜索洞察页面中, 使用系统默认行业（无需切换行业）, 找到时间范围选择器
        4.根据用户输入的一到多个时间范围, 逐个进行以下操作(时间范围列表见下方): 
            时间范围列表: 
            {insight_time_ranges_formatted}
            对每个时间范围依次执行: 
            4.1 选择时间范围类型(从上方列表取当前项；选项包括: 近7天、近30天、自定义、按周、按月)
            4.2 如果选择的是"自定义", 请点击日期选择器, 通过日历控件选择日期范围(从上方列表取当前项的具体日期)
            - 重要: 不要直接在输入框中输入日期, 必须通过点击日历控件来选择日期
            - 先点击开始日期, 再点击结束日期
            4.3 等待数据加载完成后, 找到"行业核心品牌"模块
            4.4 在行业核心品牌列表中, 找到{brand_name}品牌(通常为第一行)
            4.5 在该行数据中, 找到排名列提取排名, 找到环比排名列提取环比排名变化, 找到搜索流量SOV(指数)提取SOV数值
            4.6 切换到"行业搜索趋势"模块
            4.77 在核心指标区域, 提取搜索次数和搜索次数环比

    输出数据格式（字段名按照实例中的字段名，如果没找到对应字段名，则使用中文字段名）: 
    {   
        "投后报告": [{
            "品牌分类": "品牌分类名",
            "报告名": "星图报告名或竞价报告名",
            "项目整体": {
                "消耗金额": 100000,
                "曝光次数": 100000,
                "曝光人数": 100000,
                "互动率": "0.1%",
                "互动量": 12000,
                "流转A3人数": 12313,
                "行业TOP5A3流转率": "0.1%",
                "7日回搜人数": 100000,
                "七日回搜率": "0.1%",
                "CPM": 100,
                "CPE": 100,
                "CP流转A3": 100,
                "CP5A": 100,
                "CPS(消耗金额/搜索次数)": 100,
                "新增A3回搜率": "0.1%",
                "完播率": "0.23%",
                "新客占比": "0.1%",
                "拉新人数量级": 100000,
                "转化金额": 100000,
                "ROI": 100,
                "搜索次数": 100000
            },
            "5A人群资产": {
                "投后人群规模": 100000,
                "投后人群增长率": "0.1%",
                "A1流转人群": 100000,
                "A1流转率": "0.1%",
                "A2流转人群": 100000,
                "A2流转率": "0.1%",
                "A3流转人群": 100000,
                "A3流转率": "0.1%",
                "A4流转人群": 100000,
                "A4流转率": "0.1%",
                "A5流转人群": 100000,
                "A5流转率": "0.1%"
            }
        }, ...],
        "文档下载地址": ["趋势分析下载地址", "5A人群资产下载地址", "A3问询下载地址", ...]
    }
    (请全部用中文回答)
"""

GENERIC_PROMPT_TEMPLATE = """
你是巨量云图数据获取助手。请按照以下步骤操作: 

1. 确认已登录巨量云图。
2. 查找报告名为 "{report_name}" 的报告。
3. 如果需要获取"行业搜索洞察", 请设置时间范围为: {insight_date_range}(类型: {date_range_type})。
4. 提取关键数据并生成摘要。

"""

INTENT_EXTRACTION_SYSTEM_PROMPT = """
You are an intent extraction specialist. Your task is to analyze the user's request and extract specific parameters for a data scraping task on Yuntu (巨量云图).

Extract the following fields into a JSON object:
1. 报告名（二选一）:
   - `report_name` (string, optional): 单个报告名。当未使用品牌分类时必填。
   - `brand_categories` (array, optional): 黑玩品牌案例：多个品牌分类，每分类含星图报告名+消耗金额、竞价报告名+消耗金额。每项: {"category_name": "分类名可选", "report_name_star": "星图报告名", "consumption_amount_star": 100000, "report_name_bid": "竞价报告名", "consumption_amount_bid": 200000}。用户输入示例：品牌分类：分类A 星图报告xxx 10万 竞价报告yyy 20万；分类B 星图报告xxx 消耗金额 竞价报告yyy 消耗金额。与 report_name 二选一。
2. 爆文加热时间（第五步 达人及内容复盘，单值）:
   - `kol_content_date_type` (string, optional): 时间类型。Must be one of: "实时", "按周", "按月", "近7天", "近30天"。未提及则 null。
   - `kol_content_date_range` (string, optional): 时间范围（按周/按月/自定义时的具体值，如 "2025/12/01-2025/12/31"）；实时/近7天/近30天可为 null。
3. 行业搜索洞察时间（第六步，一到多值；行业使用系统默认，无需抽取）:
   - `insight_time_ranges` (array, optional): 一到多个时间范围。每项: {"date_range_type": "按周"|"按月"|"近7天"|"近30天"|"自定义", "insight_date_range": "具体日期或 null"}。若用户只说一个范围可只填一个元素。未提及则 null。
4. `brand_name` (string, optional): 用户指定的品牌名称；打开云图后需在右上角切换到此品牌（如 "黑玩"）。未提供则 null。
5. `consumption_amount_star` (number): 星图(达人营销)消耗金额，单位元。当使用单个 report_name 时必填；当使用 brand_categories 时由各分类内 consumption_amount_star 提供，此项可为 null。
6. `consumption_amount_bid` (number): 竞价(竞价投放)消耗金额，单位元。当使用单个 report_name 时必填；当使用 brand_categories 时由各分类内 consumption_amount_bid 提供，此项可为 null。

Important rules:
- If previous values are provided in the context, preserve them unless the current message provides new values.
- Analyze the entire conversation history to extract information, not just the current message.
- If a field is missing in the current message but was mentioned in previous messages, use the previous value.
- When using brand_categories, each item must include report_name_star, consumption_amount_star, report_name_bid, consumption_amount_bid. When using report_name, top-level consumption_amount_star and consumption_amount_bid are required.
- Always return valid JSON.
"""


PITCHING_PROMPT_TEMPLATE = """
    打开巨量云图(https://yuntu.oceanengine.com), 右上角用户信息切换到"{brand_name}"（英文名: {brand_name_en}）, 按照以下取数路径获取核心指标并按照以下格式输出: 


    一、市场大盘&竞争格局分析(同周期对比, MAT25VSMAT24, eg: 25.1月-12月 VS 24.1月-12月)
        目的: 界定增长赛道, 行业的增长动因、卖点、TOP品牌、商品是什么;
        [数据维度]: 
        大盘趋势: MAT(年度移动总额)销售金额、销售量、购买人数、商品数量、商品平均价格、人均购买频次、人均购买量
        品牌竞争格局: 销售额排名、品牌名称、销售额, 同比。
        行业标题关键词: 增速&量级
        取数路径: 
        巨量云图-商品-细分市场-查看对应的细分市场报告-市场概览: MAT(年度移动总额)销售金额、销售量、购买人数、商品数量、商品平均价格、人均购买频次、人均购买量
        巨量云图-商品-细分市场-查看对应的细分市场报告-商品洞察-TOP品牌分析-点击下载按钮下载数据/todo
        巨量云图-商品-细分市场-查看对应的细分市场报告-商品洞察-商品标题关键词分析-趋势分析-点击下载按钮下载数据todo？
    二、 品牌/SPU 5A人群资产健康度诊断
        目的: 量化品牌在行业中的位置, 找出资产断层(是种草A3不足, 还是转化A4太慢)。
        数据维度: 
        5A规模及占比: A1-A5各层级人群量级。
        流转率对比: 本品 vs 行业Top 5均值(A2→A3、A3→A4)。
        Benchmark偏差: 单品差距倍数。
        取数路径: 
        巨量云图-人群-5A关系资产-A1-A5各层级人群量级。
        巨量云图-人群-5A关系流转-明细表格中每行的人群分类和本品牌的流转率和对比品牌均值的流转率
        巨量云图-人群-SPU5A分布
    三、搜索洞察 
        目的: 识别竞对的防线和漏洞, 判断用户在搜索谁, 以及竞对是通过"达人内容"还是"广告加热"在赢。
        数据维度: 
        搜索SOV/排名: by月品牌搜索SOV变化, 搜索词TOP 3覆盖。
        取数路径: 
        巨量云图-营销触点-搜索概览-数据概览
    四、本竞品/行业核心人群画像与TA内容偏好洞察 
        目的: 分析我们的核心人群是怎么样的, 喜欢看什么内容；识别竞品和本品人群上的差异, 达成差异化运营
        数据维度: 
        人群画像 : 八大消费群体占比、性别、年龄、消费力分布。
        场景心智: TA 内容偏好一级/二级类目(如: 休闲娱乐、生活记录等高 TGI 类目)。
        取数路径: 
        巨量云图-人群-人群列表-查看画像
        巨量云图-数据工厂-标签工厂
        巨量云图-内容-TA内容洞察-自定义人群
    五、星推比+本竞品爆文加热LIST
    目的: 看行业&竞品会做什么样的内容, 用的怎么样的星推比, 给出参考方向
        数据维度: 
        星推比: 星图任务金额 vs 加热投放金额。
        爆文效率: 本竞品内容数量、爆文率、看后搜索率对比。
        取数路径: 
        巨量云图-营销触点-营销概览-爆文加热(竞品对比)
"""