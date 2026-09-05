#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""数模模型字典检索工具（优先用 31_模型字典_全审标注版.xlsx，带审核结论）

用法：
  python scripts/query_model_dict.py 熵权法               # 完整条目+审核结论
  python scripts/query_model_dict.py "GM(1,1)" -n 3       # 全角/半角括号自动归一
  python scripts/query_model_dict.py 灰狼 --field 禁忌点   # 只看某字段
  python scripts/query_model_dict.py 熵权法 --only-ok      # 只返回「可信」条目
  python scripts/query_model_dict.py --stats              # 审核总览
  python scripts/query_model_dict.py --groups             # 列出全部分组
  python scripts/query_model_dict.py --group 灰色系统预测  # 列出某分组模型

依赖：openpyxl。检索同时匹配 名称/场景/原理，名称命中排前。
审核结论五级：可信OK / 存疑WARN / 错误BAD / 注水PAD / 未核验UNK（详见 references/30 号说明）。
"""
import argparse
import re
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("需要 openpyxl：pip install openpyxl")

DICT_DIR = Path(__file__).resolve().parent.parent / "references"
AUDIT_FILE = DICT_DIR / "31_模型字典_全审标注版.xlsx"
RAW_FILE = DICT_DIR / "30_数模模型字典.xlsx"

FIELDS = ["序号", "模型名称", "模型大类", "具体分组", "模型类别", "适用场景", "数据要求",
          "原理讲解", "模型输入", "模型输出", "关键假设", "禁忌点", "模型缺陷", "检验方法"]
VERDICT_COL = "审核结论(2026-09-03全审)"
NOTE_COL = "审核备注"
VLA = {"OK": "可信", "WARN": "存疑", "BAD": "错误", "PAD": "注水", "UNK": "未核验"}
VLV = {v: k for k, v in VLA.items()}


def norm(s: str) -> str:
    """全角括号/逗号/冒号归一为半角并去空白，避免 GM（1,1） 搜不到 GM(1,1)。"""
    if not s:
        return ""
    s = str(s)
    table = str.maketrans({"（": "(", "）": ")", "，": ",", "：": ":", "　": " "})
    return re.sub(r"\s+", "", s.translate(table)).lower()


def load_rows():
    path = AUDIT_FILE if AUDIT_FILE.exists() else RAW_FILE
    if not path.exists():
        sys.exit(f"字典文件不存在：{DICT_DIR}")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["模型数据库"]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    header = list(rows[0])
    has_audit = VERDICT_COL in header
    data = []
    for r in rows[1:]:
        if not r[1]:
            continue
        rec = dict(zip(FIELDS, r))
        if has_audit:
            rec["_v"] = VLV.get(str(r[header.index(VERDICT_COL)]).strip(), "UNK")
            rec["_note"] = str(r[header.index(NOTE_COL)] or "").strip()
        else:
            rec["_v"], rec["_note"] = "UNK", "该文件无审核列"
        rec["_norm_name"] = norm(rec["模型名称"])
        rec["_norm_all"] = norm(" ".join(str(x) for x in r[:14]))
        data.append(rec)
    return data


VORDER = {"OK": 0, "WARN": 1, "BAD": 2, "PAD": 3, "UNK": 4}


def pretty(v) -> str:
    if v is None:
        return ""
    return str(v).replace("<br/>", "\n").replace("<br>", "\n")


def show(rec: dict, only_field: str | None = None):
    print("=" * 60)
    tag = {"OK": "✅", "WARN": "🟡", "BAD": "🔴", "PAD": "⬜", "UNK": "🔵"}.get(rec["_v"], "")
    print(f"{tag}【审核】{VLA.get(rec['_v'], rec['_v'])}" + (f"｜{rec['_note']}" if rec["_note"] else ""))
    keys = [only_field] if only_field in FIELDS else FIELDS
    for k in keys:
        val = pretty(rec.get(k, ""))
        if only_field:
            print(val)
        else:
            print(f"【{k}】{val}")
    print()


def main():
    ap = argparse.ArgumentParser(description="数模模型字典检索（带审核结论）")
    ap.add_argument("keyword", nargs="?", help="检索关键词")
    ap.add_argument("-n", type=int, default=3, help="最多显示条数（默认3）")
    ap.add_argument("--groups", action="store_true", help="列出全部分组")
    ap.add_argument("--group", help="列出某分组下所有模型名")
    ap.add_argument("--field", help="只显示指定字段")
    ap.add_argument("--only-ok", action="store_true", help="只显示「可信」条目")
    ap.add_argument("--verdict", help="按审核等级过滤：可信/存疑/错误/注水/未核验")
    ap.add_argument("--list-all", action="store_true", help="命中时只列名称不展开")
    ap.add_argument("--stats", action="store_true", help="审核总览统计")
    args = ap.parse_args()

    data = load_rows()

    if args.stats:
        from collections import Counter
        c = Counter(r["_v"] for r in data)
        total = len(data)
        print(f"模型字典全审总览（2026-09-03，共{total}条）")
        for v in ["OK", "WARN", "BAD", "PAD", "UNK"]:
            n = c.get(v, 0)
            print(f"  {VLA[v]}: {n}（{n/total:.1%}）")
        print("用法规则：选型只用可信；存疑按备注核对；错误禁引结论；注水=换皮/组合凑数；未核验须对照教材。详见 references/30 号说明。")
        return

    if args.groups:
        from collections import Counter
        c = Counter(f"{r['模型大类']} / {r['具体分组']}" for r in data)
        for k, v in c.most_common():
            print(f"{k}: {v}")
        return

    if args.group:
        hits = [r for r in data if norm(r["具体分组"]) == norm(args.group)]
        print(f"分组「{args.group}」共 {len(hits)} 条：")
        for r in hits:
            flag = {"OK": "✅", "WARN": "🟡", "BAD": "🔴", "PAD": "⬜", "UNK": "🔵"}.get(r["_v"], "")
            print(f"  #{r['序号']} {flag}{r['模型名称']}")
        return

    if not args.keyword:
        ap.print_help()
        return

    kw = norm(args.keyword)
    named = [r for r in data if kw in r["_norm_name"]]
    other = [r for r in data if kw not in r["_norm_name"] and kw in r["_norm_all"]]
    hits = named + other

    if args.only_ok:
        hits = [r for r in hits if r["_v"] == "OK"]
    if args.verdict:
        want = VLV.get(args.verdict)
        if not want:
            sys.exit(f"未知等级「{args.verdict}」，可选：{'/'.join(VLA.values())}")
        hits = [r for r in hits if r["_v"] == want]
    # 可信优先，注水垫底
    hits.sort(key=lambda r: (VORDER.get(r["_v"], 9), r["_norm_name"].find(kw) if kw in r["_norm_name"] else 999))

    if not hits:
        hint = "（若开了 --only-ok/--verdict 过滤可放宽试试）" if (args.only_ok or args.verdict) else ""
        print(f"未命中「{args.keyword}」{hint}。可先 --groups 看分组，或换关键词（如英文缩写/中文名）。")
        return

    print(f"命中 {len(hits)} 条（名称命中 {len(named)} 条）" + (f"，显示前 {min(args.n, len(hits))} 条（可信优先）：" if not args.list_all else "："))
    for r in hits if args.list_all else hits[: args.n]:
        if args.list_all:
            flag = {"OK": "✅", "WARN": "🟡", "BAD": "🔴", "PAD": "⬜", "UNK": "🔵"}.get(r["_v"], "")
            print(f"  #{r['序号']} {flag}{r['模型名称']}  [{r['模型大类']}/{r['具体分组']}]")
        else:
            show(r, args.field)
    if len(hits) > args.n and not args.list_all:
        print(f"（还有 {len(hits) - args.n} 条未展开，用 --list-all 看名单或加大 -n）")


if __name__ == "__main__":
    main()
