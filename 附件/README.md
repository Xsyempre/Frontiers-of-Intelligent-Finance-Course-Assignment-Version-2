# 附件：Python 脚本、提示词与可视化工具

## 目录结构

```
附件/
├── py_ver/
│   ├── config.py          # API 配置（模型、密钥、路径）
│   ├── annotate_api.py    # 全量标注脚本（并行版）
│   ├── quality_check.py   # 质量检查脚本（三层检查）
│   ├── calc_f1.py         # F1 Score 计算脚本
│   └── test_api.py        # API 连通性测试脚本
├── prompts/
│   └── system_prompt.md   # 标注用 System Prompt（外部文件）
├── demo.html              # 标注界面 Demo（浏览器直接打开）
├── dashboard.html         # 数据可视化仪表盘（浏览器直接打开）
└── README.md              # 本文件
```

## 设计说明

### 为什么选择 HTML + API 方案？

在完成本作业的过程中，我尝试了两种技术路线：

1. **Python 脚本**：通过 `requests` 库直接调用 API，适合批量处理和自动化
2. **HTML 界面**：通过浏览器 `fetch` API 调用，可视化操作，无需配置 Python 环境

HTML 方案的优势在于：
- **零配置**：双击即可打开，无需安装 Python 或依赖库
- **可视化**：实时进度条、日志滚动、统计卡片，操作直观
- **可分享**：一个 HTML 文件即可交付，老师可直接打开验证

最终我选择 Python 脚本完成全量标注（因为需要并行处理和断点续跑），但 HTML 界面作为 Demo 保留，展示了另一种可行的技术路线。

### 提示词管理

根据老师的要求，提示词与代码分离：
- `prompts/system_prompt.md`：标注用 System Prompt，脚本通过 `load_prompt()` 动态加载
- 便于版本管理和迭代更新（v1 → v2 补充边界规则）

## 使用方法

### Python 脚本

```bash
cd py_ver

# 1. 测试 API 连通性（前 5 行）
python test_api.py

# 2. 全量标注（3721 行，约 35 分钟）
python annotate_api.py

# 3. 质量检查
python quality_check.py --phase all

# 4. 计算 F1 Score
python calc_f1.py
```

### HTML 工具

- **demo.html**：双击打开，配置 API → 上传 xlsx → 批量标注 → 下载结果
- **dashboard.html**：双击打开，查看标注数据的可视化分析图表

## 依赖

- Python 3.8+
- openpyxl
- requests
- 浏览器（Chrome/Edge/Firefox）
