# Yuntu Data Scraper (巨量云图数据助手)

A Gradio-based application for extracting data from Yuntu (巨量云图) using Browser Use and Gemini 3 Pro.

## Features

- **Chat Interface**: Natural language task descriptions
- **Intent Recognition**: Automatic extraction of report name, date range, and brand
- **Persistent Login**: Browser profile saved to `browser_profiles/yuntu_profile`
- **Brand-Specific Templates**: Special handling for Heyone (黑玩)
- **Excel Export**: Automatically extracts data, processes downloaded Excel files, and generates formatted Excel reports
- **Data Processing**: Reads Excel files for KOL content review and ad flow analysis, calculates metrics (CPM, CPE, CPS, CPA3, ROI)

## Requirements

- Python 3.11+
- `GOOGLE_API_KEY` environment variable set

## Installation

```bash
# From project root
# Install all dependencies including Excel processing libraries (pandas, openpyxl)
uv sync --all-extras

# Or if you only need code dependencies
uv sync --extra code
```

## Usage

```bash
# Run the application
python examples/apps/yuntu-data/app.py
```

Then open your browser to the displayed URL (typically `http://127.0.0.1:7860`).

## Example Task

```
从巨量云图中帮我获取一下黑玩的复盘数据，报告名：ZY-Heyone12.1-12.31；行业搜索洞察的日期范围是：自定义（2025/12/01-2025/12/31）
```

## Intent Fields

The system extracts the following from your task description:

| Field | Description | Example |
|-------|-------------|---------|
| `report_name` | Name of the report | ZY-SHK12.1-12.31 |
| `brand_name` | Brand name (optional) | 黑玩 |
| `date_range_type` | Type of date selection | 近7天, 近30天, 自定义, 按周, 按月 |
| `insight_date_range` | Specific date range | 2025/12/01-2025/12/31 |

## First Run

On first run, if no login profile exists, the application will:

1. Open a browser window
2. Navigate to the Yuntu login page
3. Wait for you to manually log in
4. Save your session for future runs

## Output Files

After running a task, the system will:

1. Extract structured data from Agent execution results
2. Process downloaded Excel files (KOL content review, ad flow analysis)
3. Calculate metrics (CPM, CPE, CPS, CPA3, ROI)
4. Generate a formatted Excel report with multiple sheets:
   - 项目整体 (Project Overview)
   - 5A人群资产流转 (5A Asset Flow)
   - 达人及内容复盘 (KOL Content Review)
   - 搜索与溢出价值 (Search Insight)
   - 投流数据精细化复盘 (Ad Flow Review)
   - TA人群画像精准度复盘 (TA Portrait Review)
   - 汇总 (Summary)

Excel reports are saved to `data-output/excel_exports/` directory.

## Configuration

Edit `config.py` to customize:

- `GEMINI_MODEL`: The Gemini model to use (default: `gemini-3-pro-preview`)
- `YUNTU_LOGIN_URL`: The login URL
- `EXCEL_OUTPUT_DIR`: Excel report output directory
- `DOWNLOADS_PATH`: Directory for downloaded Excel files
- Prompt templates for different brands

## Optional: Update Prompts

For better accuracy, you can manually read the Word document (`docs/复盘路径拆解.docx`) and update the prompt templates in `config.py` based on the detailed data extraction paths.
