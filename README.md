# Excel 翻译工具 (Excel Translation Tool)

一个使用 Google 翻译的 Excel 文件翻译工具，支持中文到英文的自动翻译。

## 功能特性

- **Google 翻译集成** - 使用 `deep_translator` 库调用 Google 翻译 API
- **多线程并发** - 支持多线程并发翻译，提升处理速度（默认 8 线程）
- **翻译缓存** - 自动缓存已翻译内容，避免重复翻译
- **词典匹配** - 内置专业术语词典，优先使用预定义翻译
- **格式保留** - 保留 Excel 原格式和合并单元格
- **列选择** - 支持指定列翻译，跳过表头
- **重试机制** - 翻译失败自动重试（默认 2 次）

## 依赖安装

```bash
pip install openpyxl deep-translator
```

## 配置说明

### 核心配置

在脚本顶部修改以下参数：

```python
DEBUG = True          # True 时打印调试信息
MAX_RETRIES = 2       # 翻译失败重试次数
RETRY_DELAY = 0.5     # 重试间隔（秒）
```

### 词典配置

`test_dictionary` 包含预定义的中文到英文映射：

```python
test_dictionary = {
    # 组合词（优先匹配）
    "新增成功": "Successfully Added",
    "导入成功": "Import Successful",
    "保存成功": "Save Successful",
    # 单个词
    "保存": "Save",
    "删除": "Delete",
    "确认": "Confirm",
    # 专业术语
    "设备台账": "Equipment Ledger",
    "点位管理": "IoT Management",
    # ... 更多术语
}
```

## 使用方法

### 基本用法

```python
from excel_translation import translate_excel

translate_excel(
    input_file='input.xlsx',      # 输入文件
    output_file='output.xlsx',    # 输出文件
    columns_to_translate=[2, 4],  # 要翻译的列（从 1 开始）
    skip_header=True,             # 是否跳过表头
    max_workers=8,                # 并发线程数
    dedupe=True                   # 是否去重翻译
)
```

### 命令行运行

```bash
python excel_translation.py
```

### 参数说明

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `input_file` | str | 必需 | 输入 Excel 文件路径 |
| `output_file` | str | 必需 | 输出 Excel 文件路径 |
| `columns_to_translate` | list | None | 要翻译的列索引列表（从 1 开始），None 表示翻译所有列 |
| `skip_header` | bool | True | 是否跳过第一行（表头） |
| `max_workers` | int | 8 | 并发线程数 |
| `dedupe` | bool | True | 是否对相同文本去重后翻译 |

## 内置词典

当前内置词典包含以下类别：

| 类别 | 示例 |
|------|------|
| 操作状态 | 新增成功、保存成功、导入成功 |
| 基础操作 | 保存、删除、确认、取消、新增、编辑 |
| 测试相关 | 测试用例、前置条件、测试步骤、预期结果 |
| 测试类型 | 功能测试、回归测试、冒烟测试、性能测试 |
| 设备相关 | 设备、设备类型、设备厂商、传感器 |
| 平台相关 | 运营平台、点位管理、物联网管理 |

## 工作流程

1. **加载 Excel** - 保留所有格式和样式
2. **解除合并** - 记录合并单元格信息后解除
3. **收集内容** - 收集含中文的单元格
4. **翻译处理** - 先词典匹配，再 Google 翻译
5. **回填结果** - 将翻译结果写回单元格
6. **恢复格式** - 重新合并单元格
7. **保存文件** - 输出翻译后的 Excel

## 注意事项

- 需要网络连接才能使用 Google 翻译
- 大量翻译时可能被 Google 限流，可通过 `RETRY_DELAY` 调整
- 词典中的组合词优先于单个词匹配（按长度从长到短）
- 翻译结果会缓存到内存，处理大文件时注意内存占用
