# 运行指南

## 快速开始

### 1. 安装依赖

确保已安装所有必需的依赖，包括Excel处理相关的库：

```bash
# 从项目根目录执行
uv sync --all-extras

# 或者如果只需要code依赖（包含pandas和openpyxl）
uv sync --extra code
```

### 2. 配置环境变量

创建或更新 `.env` 文件（在项目根目录），确保包含：

```bash
GOOGLE_API_KEY=your_google_api_key_here
```

### 3. （可选）更新提示词模板

根据Word文档（`docs/复盘路径拆解.docx`）更新 `config.py` 中的提示词模板，以获得更准确的取数路径。这是一次性任务，如果不更新也可以运行，但Agent可能无法完全按照Word文档中的详细步骤执行。

### 4. 运行应用

```bash
# 从项目根目录执行
python examples/apps/yuntu-data/app.py
```

或者：

```bash
# 进入应用目录
cd examples/apps/yuntu-data
python app.py
```

### 5. 访问界面

应用启动后，会显示一个URL（通常是 `http://127.0.0.1:7860`），在浏览器中打开即可使用。

## 首次运行

### 登录巨量云图

首次运行时，如果没有登录信息，应用会：

1. 自动打开浏览器窗口
2. 导航到巨量云图登录页面
3. 等待您手动登录
4. 登录成功后自动保存会话（保存在 `browser_profiles/yuntu_profile`）

**注意**：请确保完全关闭Chrome浏览器后再运行应用，否则可能无法连接到浏览器。

## 使用示例

在Gradio界面中输入任务描述，例如：

```
从巨量云图中帮我获取一下黑玩的复盘数据，报告名：ZY-Heyone12.1-12.31；行业搜索洞察的日期范围是：自定义（2025/12/01-2025/12/31）
```

系统会：
1. 自动识别报告名、品牌名、日期范围
2. 使用Agent执行数据抓取
3. 自动处理下载的Excel文件
4. 生成格式化的Excel报告

## 输出文件

- **Excel报告**：保存在 `data-output/excel_exports/` 目录
- **对话记录**：保存在 `data-output/conversations/` 目录
- **下载的Excel文件**：保存在 `downloads/` 目录（项目根目录）

## 故障排除

### 1. 依赖缺失错误

如果遇到 `ModuleNotFoundError`，确保已安装所有依赖：

```bash
uv sync --all-extras
```

### 2. Excel文件读取失败

- 检查下载目录中是否有Excel文件
- 确认文件是最近24小时内下载的
- 检查文件是否损坏

### 3. Agent执行失败

- 检查提示词模板是否正确
- 确认已登录巨量云图
- 查看浏览器窗口中的错误信息

### 4. 数据提取失败

- 检查Agent输出格式是否匹配预期
- 查看 `data-output/conversations/` 中的对话记录
- 可能需要调整 `data_extractor.py` 中的正则表达式

## 下一步优化

1. **更新提示词**：根据Word文档完善取数路径
2. **调整数据提取**：根据实际Agent输出调整 `data_extractor.py` 中的解析逻辑
3. **优化Excel处理**：根据实际Excel文件结构调整 `excel_processor.py` 中的读取逻辑
4. **添加用户输入**：在UI中添加选项值输入功能（如人群包选择、SKU选择等）
