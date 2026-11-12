"""抓取 143 名干员的 PRTS 头像，压成 WebP 44px，产出 base64 缓存。

用法（office-runtime python 或 .py310 均可，需 Pillow）：
    py -3 tools/fetch_avatars.py             # 全量，已有缓存则跳过
    py -3 tools/fetch_avatars.py --force     # 忽略缓存重抓

产物：_tmp/avatars_webp44.json  —— {干员名: "data:image/webp;base64,..."}
（_tmp/ 已被 .gitignore 忽略；真正的头像会由 tools/build_dataset.py 内嵌进 HTML）
数据源：PRTS 文件页「文件:头像 <干员名>.png」，经 api.php imageinfo 取 64px 缩略图，
        再本地重压到 44px WebP（q80）——实测 143 张合计约 325 KB base64。
"""
from __future__ import annotations

import argparse
import csv
import io
import json
import os
import ssl
import sys
import time
import urllib.parse
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CSV_PATH = ROOT / "卫戍协议_干员数据库.csv"
OUT = ROOT / "_tmp" / "avatars_webp44.json"

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/120.0 Safari/537.36 dsh-scraper/0.1")
HDR = {"User-Agent": UA, "Accept-Language": "zh-CN,zh;q=0.9", "Referer": "https://prts.wiki/"}
CTX = ssl.create_default_context()

TILE = 44
QUALITY = 80
BATCH = 45


def http_get(url: str) -> bytes:
    req = urllib.request.Request(url, headers=HDR)
    with urllib.request.urlopen(req, timeout=45, context=CTX) as resp:
        return resp.read()


def names() -> list[str]:
    with CSV_PATH.open(encoding="utf-8-sig", newline="") as fh:
        return [row["干员"].strip() for row in csv.DictReader(fh) if row.get("干员")]


def thumbs(batch: list[str]) -> dict[str, str]:
    """一次问一批文件页，返回 {干员名: 64px缩略图URL}。"""
    titles = "|".join("文件:头像 %s.png" % n for n in batch)
    url = ("https://prts.wiki/api.php?action=query&format=json&prop=imageinfo"
           "&iiprop=url|size&iiurlwidth=64&titles="
           + urllib.parse.quote(titles, safe="|:"))
    data = json.loads(http_get(url).decode("utf-8"))
    out: dict[str, str] = {}
    for page in data.get("query", {}).get("pages", {}).values():
        title = page.get("title") or ""
        name = title.replace("文件:头像 ", "").replace(".png", "")
        info = (page.get("imageinfo") or [{}])[0]
        if info.get("thumburl"):
            out[name] = info["thumburl"]
    return out


def to_webp(raw: bytes) -> str:
    from PIL import Image  # 延迟导入，--check 模式不需要

    im = Image.open(io.BytesIO(raw))
    if im.mode not in ("RGB", "RGBA"):
        im = im.convert("RGBA")
    im.thumbnail((TILE, TILE), Image.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, "WEBP", quality=QUALITY, method=6)
    return "data:image/webp;base64," + __import__("base64").b64encode(buf.getvalue()).decode("ascii")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--force", action="store_true", help="忽略已有缓存，全部重抓")
    args = ap.parse_args()

    cache: dict[str, str] = {}
    if OUT.exists() and not args.force:
        cache = json.loads(OUT.read_text(encoding="utf-8"))
        print("缓存：%d 张" % len(cache))

    todo = [n for n in names() if n not in cache]
    if not todo:
        print("全部命中缓存，无需抓取")
        return 0

    print("需要抓取：%d 张" % len(todo))
    failed: list[str] = []
    got = 0
    for i in range(0, len(todo), BATCH):
        batch = todo[i:i + BATCH]
        try:
            url_of = thumbs(batch)
        except Exception as exc:  # 单批失败不影响其他批
            print("  批次失败 %d-%d: %s" % (i, i + len(batch), exc))
            failed.extend(batch)
            continue
        for name in batch:
            url = url_of.get(name)
            if not url:
                failed.append(name)
                continue
            try:
                cache[name] = to_webp(http_get(url))
                got += 1
            except Exception as exc:
                print("  转换失败 %s: %s" % (name, exc))
                failed.append(name)
            time.sleep(0.05)
        print("  进度 %d/%d" % (min(i + BATCH, len(todo)), len(todo)))
        time.sleep(0.4)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(cache, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    total = sum(len(v) for v in cache.values())
    print("成功 %d / 失败 %d | 缓存 %d 张 | base64 合计 %.0f KB"
          % (got, len(failed), len(cache), total / 1024))
    if failed:
        print("失败名单：" + "、".join(failed))
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
