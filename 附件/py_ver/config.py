# -*- coding: utf-8 -*-
"""
API 配置文件
"""

# Mimo 模型配置
BASE_URL = "https://token-plan-cn.xiaomimimo.com/v1"
API_KEY = "your-api-key-here"  # 请填入你的 API Key
MODEL = "mimo-v2.5-pro"

# 并发数
MAX_WORKERS = 5

# 保存间隔
SAVE_INTERVAL = 50

# 文件路径
INPUT_FILE = "STK_Violation_Main_第2组.xlsx"
OUTPUT_FILE = "STK_labeled_夏思远_2023110537_G02.xlsx"
PROMPT_FILE = "prompts/system_prompt.md"
