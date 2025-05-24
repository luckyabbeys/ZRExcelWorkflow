# Excel自动化处理工作流

## 项目简介

这是一个用于自动化处理Excel文件的工作流系统，可以批量处理多个Excel文件并将结果合并。整个工作流分为三个阶段：

1. **单文件处理**：处理单个Excel文件中的多个sheet
2. **批量处理**：批量处理多个Excel文件
3. **结果合并**：合并所有处理结果到一个最终文件

## 功能特点

- 🚀 **高效处理**：支持并行处理多个文件
- 📊 **数据整合**：自动合并多个Excel文件的数据
- 🔍 **错误追踪**：详细的日志记录和错误报告
- 🛠️ **灵活配置**：可以选择运行特定阶段或处理特定sheet
- 📈 **报告生成**：自动生成处理报告

## 快速开始

### 安装依赖

```bash
pip install -r requirements.txt
```

### 准备数据

1. 将原始数据文件放入 `data/input` 目录
2. 确保原始数据文件包含必要的sheet：
   - 门急诊信息
   - 住院信息
   - 药物医嘱信息
   - 吸氧信息
   - 检查信息
   - 统计数据

### 运行工作流

```bash
python main.py
```

这将运行完整的三阶段工作流。

### 查看结果

处理完成后，可以在以下位置查看结果：

- 第一阶段结果：`data/output/测试合并.xlsx`
- 最终合并结果：`data/final/merged_results.xlsx`
- 处理报告：`data/output/batch_process_report.xlsx` 和 `data/final/merge_report.xlsx`

## 高级用法

### 运行特定阶段

```bash
python main.py --phase 1  # 只运行第一阶段
python main.py --phase 2  # 只运行第二阶段
python main.py --phase 3  # 只运行第三阶段
```

### 处理特定sheet

```bash
python main.py --phase 1 --sheet sheet1  # 只处理第一个sheet
```

### 指定输入输出路径

```bash
python main.py --input "path/to/input" --output "path/to/output" --final "path/to/final"
```

## 项目结构

```
ZRExcelWorkflow/
├── data/                # 数据目录
│   ├── input/           # 存放原始数据文件
│   ├── output/          # 存放第二阶段处理后的文件
│   └── final/           # 存放最终合并结果
├── scripts/             # 脚本目录
│   ├── phase1/          # 第一阶段：处理单个Excel文件的脚本
│   ├── phase2/          # 第二阶段：批量处理多个Excel文件的脚本
│   └── phase3/          # 第三阶段：合并处理结果的脚本
├── utils/               # 工具函数模块
├── main.py              # 主程序入口
└── requirements.txt     # 依赖包列表
```

## 处理流程图

```
原始数据文件 → 第一阶段处理 → 单文件处理结果
                   ↓
多个原始数据文件 → 第二阶段处理 → 多个处理结果文件
                   ↓
多个处理结果文件 → 第三阶段处理 → 最终合并结果
```

## 常见问题

**Q: 如何添加新的sheet处理？**

A: 在 `scripts/phase1` 目录下创建新的处理脚本，然后在 `scripts/phase2/batch_process.py` 和 `scripts/phase3/merge_results.py` 中更新相关函数。

**Q: 处理大文件时出现内存错误怎么办？**

A: 尝试使用分块读取和处理功能，或增加系统内存。

**Q: 如何查看处理日志？**

A: 日志文件保存在项目根目录下的 `batch_process.log` 和 `merge_results.log`。

## 更多信息

详细的项目规则和开发指南，请参阅 [PROJECT_RULES.md](PROJECT_RULES.md)。

## 许可证

本项目采用 MIT 许可证。详见 LICENSE 文件。

---

*开发者：您的团队*