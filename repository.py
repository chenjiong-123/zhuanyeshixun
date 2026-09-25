"""测试数据读取模块。"""

from __future__ import annotations

import json
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent
DATA_DIR = PROJECT_ROOT / "data"


def load_samples() -> list[dict[str, object]]:
    path = DATA_DIR / "sample_essays.json"
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, list):
        raise ValueError("测试数据格式错误")
    required = {"id", "title", "prompt", "essay", "reference_score"}
    for item in data:
        if not isinstance(item, dict) or not required.issubset(item):
            raise ValueError("测试数据缺少必要字段")
    return data


def public_samples() -> list[dict[str, object]]:
    return [
        {"id": item["id"], "title": item["title"], "prompt": item["prompt"], "essay": item["essay"]}
        for item in load_samples()
    ]

