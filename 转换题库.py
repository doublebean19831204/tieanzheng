#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# SIG:BEGIN:header
# ============================================================
#  署名信息 / ATTRIBUTION  —  请勿删除 (DO NOT REMOVE)
#  制作人：徐徐 任欣
#  微信：doublebean_vip
#  Copyright (c) 2026 徐徐 任欣. All Rights Reserved.
# ============================================================
# SIG:END:header
"""
铁安证题库转换器
================
读取 `铁安证考试题库（参考）.xlsx`（单工作表，含「类型」列区分单选/多选/判断），
按类型拆分为 单选题 / 多选题 / 判断题 三个分类，生成与「桂水安」题库完全一致的
数据结构：

    questions.json      —— 纯数据（供 fetch 加载）
    questions-data.js   —— window.QUIZ_DATA = {...}（双击 index.html 即可离线使用）

每题结构：
    {
      "question": "题干",
      "options": ["A. ...", "B. ...", ...],
      "correct": 2 或 [0, 1, 2],     # 多选为索引数组
      "explanation": ""              # 源表无解析列，统一留空
    }

源表说明：
  - 列：类型 / 题目 / 选项A / 选项B / 选项C / 选项D / 正确答案
  - 判断题答案有两种写法：字母（A=正确, B=错误）或直接写「正确 / 错误」，
    部分判断题选项为空——统一归一化为「A. 正确 / B. 错误」。

用法：
    python 转换题库.py                       # 转换并写出两个数据文件
    python 转换题库.py --dry-run             # 只统计，不写文件
"""

import json
import os
import re
import sys

try:
    import openpyxl
except ImportError:  # pragma: no cover
    sys.exit("缺少依赖 openpyxl，请先执行：pip install openpyxl")

ROOT = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(ROOT, "铁安证考试题库（参考）.xlsx")

# 列索引（0 基）
COL_TYPE = 0      # 类型
COL_STEM = 1      # 题目
COL_OPT_A = 2     # 选项A
COL_OPT_D = 5     # 选项D
COL_ANS = 6       # 正确答案

LETTERS = "ABCDEF"
CJK = re.compile(r"[\u3000-\u303f\u4e00-\u9fff\uff00-\uffef]")

# 类型列 -> 输出分类名（与桂水安保持一致）
CATEGORY_ORDER = ["单选题", "多选题", "判断题"]


def _is_cjk(ch):
    return bool(ch) and bool(CJK.match(ch))


def clean(value):
    """单元格文本清洗：去 BOM/首尾空白，按中英文规则合并被折行的换行。"""
    if value is None:
        return ""
    s = str(value).replace("\ufeff", "")
    s = s.replace("\r\n", "\n").replace("\r", "\n")

    def join(m):
        before = m.string[: m.start()]
        after = m.string[m.end():]
        a = before[-1] if before else ""
        b = after[0] if after else ""
        return "" if (_is_cjk(a) and _is_cjk(b)) else " "

    s = re.sub(r"[ \t]*\n[ \t]*", join, s)
    s = re.sub(r"[ \t]{2,}", " ", s)
    return s.strip()


def norm_answer(raw):
    """答案归一化：'abcd' / 'A、B' / 'ＡＢ' -> 'ABCD'；'正确'/'错误' 原样返回。"""
    s = clean(raw)
    if not s:
        return ""
    out = []
    for ch in s:
        code = ord(ch)
        if 0xFF21 <= code <= 0xFF3A:
            out.append(chr(code - 0xFF21 + ord("A")))
        elif 0xFF41 <= code <= 0xFF5A:
            out.append(chr(code - 0xFF41 + ord("a")))
        else:
            out.append(ch)
    s = "".join(out)
    upper = s.upper()
    if re.fullmatch(r"[A-F\s、,，/|]+", upper):
        letters = sorted({c for c in upper if c in LETTERS})
        return "".join(letters)
    return s


def letters_to_indices(ans, n_options):
    idx = [LETTERS.index(c) for c in ans if c in LETTERS]
    idx = [i for i in idx if 0 <= i < n_options]
    return sorted(set(idx))


def parse_rows(rows):
    """把源表所有数据行按类型拆分并转换 -> (data, problems, stats)"""
    buckets = {c: [] for c in CATEGORY_ORDER}
    problems = []
    type_counter = {}

    for lineno, row in enumerate(rows[1:], start=2):
        row = list(row) + [None] * (COL_ANS + 1 - len(row))
        raw_type = clean(row[COL_TYPE])
        stem = clean(row[COL_STEM])
        if not stem:
            continue

        # 类型归一化
        if "多选" in raw_type:
            category = "多选题"
        elif "判断" in raw_type:
            category = "判断题"
        elif "单选" in raw_type:
            category = "单选题"
        else:
            problems.append("第%d行 未知题型：%r" % (lineno, raw_type))
            continue
        type_counter[category] = type_counter.get(category, 0) + 1

        ans_raw = norm_answer(row[COL_ANS])

        options = []
        for i in range(COL_OPT_A, COL_OPT_D + 1):
            text = clean(row[i]) if i < len(row) else ""
            if text:
                options.append("%s. %s" % (LETTERS[len(options)], text))

        if category == "判断题":
            if ans_raw in ("正确", "对", "√", "T", "TRUE"):
                correct = 0
            elif ans_raw in ("错误", "错", "×", "F", "FALSE"):
                correct = 1
            else:
                idx = letters_to_indices(ans_raw, 2)
                if len(idx) != 1:
                    problems.append("第%d行 判断题答案无法识别：%r" % (lineno, row[COL_ANS]))
                    continue
                correct = idx[0]
            options = ["A. 正确", "B. 错误"]
        else:
            if len(options) < 2:
                problems.append("第%d行 选项不足（%d 个）：%s" % (lineno, len(options), stem[:30]))
                continue
            idx = letters_to_indices(ans_raw, len(options))
            if not idx:
                problems.append("第%d行 答案无法识别：%r" % (lineno, row[COL_ANS]))
                continue
            if category == "多选题":
                if len(idx) < 2:
                    problems.append("第%d行 多选题答案只有 1 项：%r" % (lineno, ans_raw))
                correct = idx
            else:
                if len(idx) > 1:
                    problems.append("第%d行 单选题答案有多项：%r（取第一项）" % (lineno, ans_raw))
                correct = idx[0]

        buckets[category].append({
            "question": stem,
            "options": options,
            "correct": correct,
            "explanation": "",   # 源表无解析列
        })

    data = {c: buckets[c] for c in CATEGORY_ORDER if buckets[c]}
    stats = [(c, len(buckets[c])) for c in CATEGORY_ORDER]
    return data, problems, stats


def convert():
    if not os.path.isfile(XLSX):
        sys.exit("找不到题库源文件：%s" % XLSX)
    wb = openpyxl.load_workbook(XLSX, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        rows = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()
    return parse_rows(rows)


def write_outputs(data):
    json_path = os.path.join(ROOT, "questions.json")
    js_path = os.path.join(ROOT, "questions-data.js")

    with open(json_path, "w", encoding="utf-8", newline="\n") as f:
        f.write(json.dumps(data, ensure_ascii=False, indent=2) + "\n")

    with open(js_path, "w", encoding="utf-8", newline="\n") as f:
        f.write("window.QUIZ_DATA = "
                + json.dumps(data, ensure_ascii=False, indent=2)
                + ";\n")
    return json_path, js_path


def main():
    dry = "--dry-run" in sys.argv[1:]

    print("=" * 62)
    print("📚 铁安证题库转换器")
    print("=" * 62)
    print("📖 源文件：%s" % os.path.basename(XLSX))

    data, problems, stats = convert()

    print("\n📊 转换统计：")
    total = 0
    for category, count in stats:
        print("   %-6s %4d 题" % (category, count))
        total += count
    print("   合计：%d 题" % total)

    if problems:
        print("\n⚠️  需要人工确认的 %d 处：" % len(problems))
        for p in problems[:40]:
            print("   - %s" % p)
        if len(problems) > 40:
            print("   ...（其余 %d 处省略）" % (len(problems) - 40))
    else:
        print("\n✅ 全部题目解析正常，无异常。")

    if dry:
        print("\n（--dry-run：未写入文件）")
        return 1 if problems else 0

    jp, jsp = write_outputs(data)
    print("\n✅ 已写入：%s" % os.path.relpath(jp, ROOT))
    print("✅ 已写入：%s" % os.path.relpath(jsp, ROOT))
    print("\n📝 下一步：python 验证.py 校验格式，然后双击 index.html 做题。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
