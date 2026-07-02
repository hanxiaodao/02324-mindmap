"""
02324 离散数学 · 思维导图 生成器

用法:
  python generate.py

从 data/content.json（如果存在）或 data/content.yaml 读取结构化数据，
生成交互式 HTML 思维导图到 index.html。

如果数据文件不存在，则检查当前目录是否有 index.html，
有则跳过生成（保留已有文件），无则创建一个空白骨架。

后续接入:
  - 替换 data/ 下的数据源为从语雀/Notion API 拉取的内容
  - generate.py 本身不需要改，只换数据
"""

import json
import os
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))

# ── 数据文件路径（后续可改为从语雀/Notion API 获取） ──
DATA_FILE = os.path.join(ROOT, "data", "content.json")
ALTS = [
    os.path.join(ROOT, "data", "content.yaml"),
    os.path.join(ROOT, "data", "content.toml"),
]

OUTPUT = os.path.join(ROOT, "index.html")


def load_data():
    """加载结构化数据"""
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    for alt in ALTS:
        if os.path.exists(alt):
            print(f"Found alternative data file: {alt}")
            # 后续可支持 yaml/toml
            return None
    return None


def generate(data):
    """从结构化数据生成 HTML"""
    # TODO: 接入结构化数据后实现
    pass


def main():
    data = load_data()
    if data:
        generate(data)
        print(f"✅ Generated {OUTPUT}")
    elif os.path.exists(OUTPUT):
        print(f"ℹ️  No data file found, keeping existing {OUTPUT}")
    else:
        print("⚠️  No data file and no existing index.html, nothing to do")


if __name__ == "__main__":
    main()
