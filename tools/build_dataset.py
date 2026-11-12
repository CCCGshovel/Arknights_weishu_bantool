#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从源数据生成工具内嵌数据集。

输入：
  卫戍协议_干员数据库.csv        —— 特质正文（基础 / 精锐）、原型干员、阶数、特质类型
  卫戍协议_盟约交叉表.xlsx「按阶明细」—— 核心盟约 / 附加盟约 拆分
输出：
  干员卡池.html 中 <!--WS-DATA-START--> ... <!--WS-DATA-END--> 之间的 JSON

用法： python tools/build_dataset.py
"""
from __future__ import annotations

import csv
import datetime
import json
import re
import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
CSV_PATH = ROOT / "卫戍协议_干员数据库.csv"
XLSX_PATH = ROOT / "卫戍协议_盟约交叉表.xlsx"
LAYERS_XLSX = ROOT / "卫戍协议_盟约层数表.xlsx"
SKLAND_DIR = ROOT / "lineup" / "data" / "skland"
HTML_PATH = ROOT / "干员卡池.html"
INDEX_PATH = ROOT / "index.html"

# 十大核心盟约（有序）
CORE = ["炎", "萨尔贡", "维多利亚", "谢拉格", "阿戈尔",
        "叙拉古", "卡西米尔", "乌萨斯", "伊比利亚", "莱塔尼亚"]

# 副盟约全序（与「按阶明细」表一致，去掉 0 干员的「绝技」）
SUB_ALL = ["精准", "迅捷", "灵巧", "坚守", "助力", "远见", "奇迹",
           "投资人", "突袭", "不屈", "调和", "协防干员", "独行", "聚合", "奥术"]
# 副盟约全部 15 个都允许被选作禁用（不做「隐性不可禁」的限制）
SUB_BANNABLE = list(SUB_ALL)
# 不再有恒可出的副盟约
SUB_FREE = []

TIER_ROMAN = {1: "Ⅰ", 2: "Ⅱ", 3: "Ⅲ", 4: "Ⅳ", 5: "Ⅴ", 6: "Ⅵ"}

DATA_START = "<!--WS-DATA-START-->"
DATA_END = "<!--WS-DATA-END-->"
SCRIPT_OPEN = '<script id="dataset" type="application/json">'
SCRIPT_CLOSE = "</script>"


def load_csv() -> dict[str, dict]:
    with CSV_PATH.open("r", encoding="utf-8-sig", newline="") as fh:
        rows = list(csv.DictReader(fh))
    out = {}
    for r in rows:
        name = (r["干员"] or "").strip()
        if not name:
            continue
        out[name] = {
            "tier": (r["阶数"] or "").strip(),
            "covenant_raw": (r["盟约"] or "").strip(),
            "type": (r["特质类型"] or "").strip(),
            "base": (r["特质（基础）"] or "").strip(),
            "elite": (r["特质（精锐）"] or "").strip(),
            "proto": (r["原型干员"] or "").strip(),
        }
    return out


def load_cross() -> dict[str, dict]:
    wb = openpyxl.load_workbook(XLSX_PATH, data_only=True)
    ws = wb["按阶明细"]
    out = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        name = (row[0] or "").strip() if isinstance(row[0], str) else ""
        if not name:
            continue
        core_raw = (row[2] or "").strip()
        add_raw = (row[3] or "").strip()
        out[name] = {
            "tier": (row[1] or "").strip(),
            "core": [x.strip() for x in core_raw.split("、") if x.strip()],
            "add": [x.strip() for x in add_raw.split("、") if x.strip()],
            "type": (row[4] or "").strip(),
        }
    return out


def load_skland_meta() -> tuple[dict[str, str], dict[str, str]]:
    """从 lineup/data/skland/ 快照读 charId 与盟约 ID→中文名映射（可选输入）。"""
    cids: dict[str, str] = {}
    bond_names: dict[str, str] = {}
    ops_csv = SKLAND_DIR / "干员.csv"
    if ops_csv.exists():
        with ops_csv.open("r", encoding="utf-8-sig", newline="") as fh:
            for r in csv.DictReader(fh):
                name = (r.get("干员") or "").strip()
                cid = (r.get("charId") or "").strip()
                if name and cid:
                    cids[name] = cid
    bonds_csv = SKLAND_DIR / "盟约.csv"
    if bonds_csv.exists():
        with bonds_csv.open("r", encoding="utf-8-sig", newline="") as fh:
            rd = csv.DictReader(fh)
            keys = rd.fieldnames or []
            id_key = keys[0] if keys else ""
            name_key = "名称" if "名称" in keys else (keys[1] if len(keys) > 1 else "")
            for r in rd:
                bid = (r.get(id_key) or "").strip()
                nm = (r.get(name_key) or "").strip()
                if bid and nm:
                    bond_names[bid] = nm
    return cids, bond_names


def parse_stack_cell(cell: str, tier: int) -> list[dict]:
    """解析「干员<触发时机>+N/条件（每回合至多M）」形式的一格（可含多行多卡）。"""
    out = []
    for raw_line in cell.replace("\r", "").split("\n"):
        line = raw_line.strip()
        if not line:
            continue
        if "<" in line and ">" in line:
            name, _, tail = line.partition("<")
            trigger, _, rest = tail.partition(">")
            name, trigger, rest = name.strip(), trigger.strip(), rest.strip()
        else:
            m = re.match(r"^([^+＋（(]+?)\s*(.*)$", line)
            if not m:
                continue
            name, trigger, rest = m.group(1).strip(), "", m.group(2).strip()
        if not name:
            continue
        gain = re.search(r"\+\s*(\d+)", rest)
        cap = re.search(r"(每回合|每场)至多\s*(\d+)", rest)
        out.append({
            "n": name,
            "t": tier,
            "g": int(gain.group(1)) if gain else 0,
            "c": int(cap.group(2)) if cap else 0,
            "p": cap.group(1) if cap else "",
            "tr": trigger,
            "r": line,
        })
    return out


def load_stacks() -> dict[str, list[dict]]:
    """读「主盟约层数增加」「附加盟约层数增加」两张表，得到 盟约 → 加层卡清单。"""
    if not LAYERS_XLSX.exists():
        return {}
    wb = openpyxl.load_workbook(LAYERS_XLSX, data_only=True)
    stacks: dict[str, list[dict]] = {}
    for sheet_name in ("主盟约层数增加", "附加盟约层数增加"):
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        header_row = None
        for r in range(1, min(ws.max_row, 15) + 1):
            v = ws.cell(r, 1).value
            if isinstance(v, str) and v.strip().startswith("阶数") and "配色" not in v:
                header_row = r
                break
        if header_row is None:
            continue
        cols: dict[int, str] = {}
        for c in range(2, ws.max_column + 1):
            v = ws.cell(header_row, c).value
            if isinstance(v, str) and v.strip() and "条目数" not in v:
                cols[c] = v.strip()
        for r in range(header_row + 1, ws.max_row + 1):
            tv = ws.cell(r, 1).value
            if not isinstance(tv, str):
                continue
            tv = tv.strip()
            if not tv.endswith("阶"):
                continue
            roman = tv[:-1].strip()
            if roman not in TIER_ROMAN.values():
                continue
            tier = next(k for k, v in TIER_ROMAN.items() if v == roman)
            for c, cov in cols.items():
                cell = ws.cell(r, c).value
                if not isinstance(cell, str) or not cell.strip():
                    continue
                stacks.setdefault(cov, []).extend(parse_stack_cell(cell, tier))
    return stacks


def main() -> int:
    csv_map = load_csv()
    cross_map = load_cross()
    cids, bond_names = load_skland_meta()
    stacks = load_stacks()

    problems: list[str] = []
    only_csv = sorted(set(csv_map) - set(cross_map))
    only_xlsx = sorted(set(cross_map) - set(csv_map))
    if only_csv:
        problems.append(f"仅在 csv 中存在的干员：{only_csv}")
    if only_xlsx:
        problems.append(f"仅在交叉表中存在的干员：{only_xlsx}")

    records = []
    for name, cross in cross_map.items():
        src = csv_map.get(name)
        if not src:
            continue
        if src["tier"] != cross["tier"]:
            problems.append(f"{name}: 阶数不一致 csv={src['tier']} xlsx={cross['tier']}")
        if src["type"] != cross["type"]:
            problems.append(f"{name}: 特质类型不一致 csv={src['type']} xlsx={cross['type']}")
        for c in cross["core"]:
            if c not in CORE:
                problems.append(f"{name}: 未知核心盟约 {c!r}")
        for a in cross["add"]:
            if a not in SUB_ALL:
                problems.append(f"{name}: 未知副盟约 {a!r}")
        # 交叉一致性：联盟约文本里出现的盟约必须被拆分覆盖
        covered = set(cross["core"]) | set(cross["add"])
        for part in (x.strip() for x in src["covenant_raw"].split("/")):
            if part and part in SUB_ALL + CORE and part not in covered:
                problems.append(f"{name}: 盟约 {part!r} 未出现在拆分结果中")

        roman = src["tier"]
        if roman not in TIER_ROMAN.values():
            problems.append(f"{name}: 非法阶数 {roman!r}")
            continue
        tier = next(k for k, v in TIER_ROMAN.items() if v == roman)

        records.append({
            "name": name,
            "tier": tier,
            "core": cross["core"],
            "add": cross["add"],
            "type": src["type"],
            "base": src["base"],
            "elite": src["elite"],
            "proto": src["proto"],
            "cid": cids.get(name, ""),
        })

    if problems:
        print("数据校验失败：", file=sys.stderr)
        for p in problems:
            print("  -", p, file=sys.stderr)
        return 1

    payload = {
        "built": datetime.date.today().isoformat(),
        "count": len(records),
        "tiers": [TIER_ROMAN[i] for i in range(1, 7)],
        "core": CORE,
        "subAll": SUB_ALL,
        "subBannable": SUB_BANNABLE,
        "subFree": SUB_FREE,
        "bondNames": bond_names,
        "stacks": stacks,
        "operators": records,
    }
    blob = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    blob = blob.replace("<", "\\u003c")  # 防止意外闭合 script 标签

    if not HTML_PATH.exists():
        print(f"未找到 {HTML_PATH}，跳过注入；仅写出 _tmp/ops.json", file=sys.stderr)
        (ROOT / "_tmp").mkdir(exist_ok=True)
        (ROOT / "_tmp" / "ops.json").write_text(blob, encoding="utf-8")
        return 0

    html = HTML_PATH.read_text(encoding="utf-8")
    i = html.find(DATA_START)
    j = html.find(DATA_END)
    if i < 0 or j < 0 or j < i:
        print(f"{HTML_PATH.name} 缺少 {DATA_START} / {DATA_END} 标记", file=sys.stderr)
        return 1
    new_html = (html[:i]
                + DATA_START + SCRIPT_OPEN + blob + SCRIPT_CLOSE + DATA_END
                + html[j + len(DATA_END):])
    HTML_PATH.write_text(new_html, encoding="utf-8")
    # 注入后同步 Pages 入口，避免 index.html 落后于本体
    INDEX_PATH.write_bytes(HTML_PATH.read_bytes())

    # 统计自检
    by_tier = {}
    for r in records:
        by_tier[r["tier"]] = by_tier.get(r["tier"], 0) + 1
    stacked = sum(1 for r in records if r["type"] == "持续叠加")
    no_cid = [r["name"] for r in records if not r["cid"]]
    stack_total = sum(len(v) for v in stacks.values())
    print(f"注入完成：{len(records)} 名干员 | 持续叠加 {stacked} | 分阶 "
          + " ".join(f"{TIER_ROMAN[t]}:{n}" for t, n in sorted(by_tier.items())))
    print(f"  加层表：{len(stacks)} 个盟约 / {stack_total} 条 | charId 缺失 {len(no_cid)}"
          + (f"（{no_cid[:6]}）" if no_cid else "")
          + f" | 盟约 ID 映射 {len(bond_names)} 条")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
