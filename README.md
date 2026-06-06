# AI辅助财务舞弊分类 — 最终提交物

**夏思远 · 2023110537 · G02组 · 智能财务前沿**

---

## 交付物清单

| 文件 | 说明 |
|------|------|
| `STK_labeled_夏思远_2023110537_G02.xlsx` | 全量标注结果（3724行，6个标注字段） |
| `validation_夏思远_2023110537_G02.xlsx` | 5%样本三模型F1 Score对比 |
| `report_夏思远_2023110537_G02.docx` | 工作报告（工作流、Prompt、模型选型、误差分析） |
| `附件/` | Python脚本、提示词文件与可视化工具 |

## 标注概况

- **数据来源**：STK_Violation_Main_第2组.xlsx（3721条证监会违规处罚公告）
- **标注模型**：Mimo V2.5 Pro
- **处理方式**：5线程并行，约35分钟完成
- **质量检查**：三层递进检查（结构规则 → 抽样语义 → 边界案例）

| 字段 | 数量 | 比例 |
|------|------|------|
| ann_related=1（影响年报） | 482 | 12.9% |
| ann_fin_flag=1（影响财务信息） | 319 | 8.6% |
| third_party_flag=1（第三方配合） | 77 | 2.1% |

## 模型对比（F1 Score）

| 模型 | ann_related | ann_fin_flag | third_party | 平均 |
|------|-------------|--------------|-------------|------|
| DeepSeek V4 Pro | 0.7925 | 0.9057 | 0.5714 | 0.7565 |
| MiniMax M2.7 | 0.7234 | 0.8750 | 0.5000 | 0.6995 |
| **Mimo V2.5 Pro** | **0.8817** | **0.7391** | **0.6667** | **0.7625** |

选定 Mimo V2.5 Pro 用于全量标注（平均F1最高）。

## 附件说明

```
附件/
├── py_ver/                     # Python 脚本
│   ├── config.py               # API 配置
│   ├── annotate_api.py         # 全量标注（并行版）
│   ├── quality_check.py        # 三层质量检查
│   ├── calc_f1.py              # F1 Score 计算
│   └── test_api.py             # API 连通性测试
├── prompts/
│   └── system_prompt.md        # 标注用 System Prompt（外部文件）
├── dashboard.html              # 数据可视化仪表盘（6页翻页，含图表）
├── demo.html                   # 标注界面 Demo（浏览器直接打开）
└── README.md                   # 附件详细说明
```

### 运行方式

```bash
cd 附件/py_ver
# 1. 在 config.py 中填入 API Key
# 2. 测试连通性
python test_api.py
# 3. 全量标注
python annotate_api.py
# 4. 质量检查
python quality_check.py --phase all
```

### HTML 工具

- **dashboard.html**：双击打开，展示标注数据的可视化分析（翻页动画、Chart.js 图表）
- **demo.html**：双击打开，无需 Python 环境即可体验 API 标注流程
