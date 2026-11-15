#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
指定影视站点爬虫（电影 / 电视剧 / 综艺 / 动漫 分类清单）
=========================================
两种模式，写在 site_rules.json 里就行，不用改代码：

  1) kind = "maccms"  —— 绝大多数“苹果CMS”影视站的通用接口
       只需要填 api 根地址，例如 https://xxx.com/api.php/provide/vod/
       脚本自动：取分类表 → 挑 电影/电视剧/综艺/动漫 → 每类按页抓条目

  2) kind = "html"    —— 普通网页列表页，给分类页地址模板 + 正则
       item_regex 里第 1 个括号 = 详情页链接，第 2 个括号 = 标题
       其他字段可以用 year_regex / remark_regex 从同一条目里再抠

用法：
    python site_crawler.py                 # 跑 site_rules.json 里所有启用的规则
    python site_crawler.py --rule 2        # 只跑第 2 条规则
    python site_crawler.py --check         # 只测接口通不通，不抓条目

输出：抓取结果/自定义站点_条目.json（make_excel.py 会读取它并生成“自定义站点条目”子表）
"""
from __future__ import annotations

import argparse
import json
import re
import ssl
import time
import urllib.parse
import urllib.request
from pathlib import Path

BASE = Path(__file__).resolve().parent
OUT = BASE / "抓取结果"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36")
SSL_CTX = ssl.create_default_context()
SSL_CTX.check_hostname = False
SSL_CTX.verify_mode = ssl.CERT_NONE

TARGET_CATES = {
    "电影": ("电影", "片"),
    "电视剧": ("连续剧", "电视剧", "剧集", "国产剧", "港台剧", "美剧", "韩剧", "日剧"),
    "综艺": ("综艺", "娱乐", "真人秀"),
    "动漫": ("动漫", "动画", "番剧", "国漫", "日漫"),
}


def get_json(url: str, timeout: int = 15, retries: int = 2):
    err = ""
    for _ in range(retries + 1):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "application/json,*/*"})
            with urllib.request.urlopen(req, timeout=timeout, context=SSL_CTX) as resp:
                body = resp.read(6_000_000)
            for enc in ("utf-8-sig", "utf-8", "gb18030"):
                try:
                    return json.loads(body.decode(enc))
                except Exception:
                    continue
            return None
        except Exception as exc:  # noqa: BLE001
            err = f"{type(exc).__name__}: {exc}"
            time.sleep(0.5)
    print(f"  [warn] 请求失败 {url[:60]} → {err[:70]}")
    return None


def get_html(url: str, timeout: int = 15) -> str:
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "text/html,*/*"})
    with urllib.request.urlopen(req, timeout=timeout, context=SSL_CTX) as resp:
        body = resp.read(6_000_000)
    for enc in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            return body.decode(enc)
        except Exception:
            continue
    return body.decode("utf-8", "ignore")


# ------------------------------------------------------------ 苹果CMS 模式
def maccms_classes(api: str) -> list:
    sep = "&" if "?" in api else "?"
    data = get_json(f"{api}{sep}ac=class")
    if not data:
        return []
    cls = data.get("class") or []
    out = []
    for item in cls:
        out.append({"id": item.get("type_id") or item.get("id"),
                    "name": str(item.get("type_name") or item.get("name") or "")})
    return [c for c in out if c["id"] is not None]


def match_cate(type_name: str) -> str:
    low = (type_name or "").strip()
    for cate, kws in TARGET_CATES.items():
        if any(kw in low for kw in kws):
            return cate
    return ""


def maccms_page(api: str, page: int, site_name: str) -> list:
    """抓一页最新条目（不带 t 过滤，按条目自己的 type_name 归类）。
    苹果CMS 的父分类 t 只返回挂在父类上的少量条目，子分类（动作片/喜剧片…）才是主库，
    所以这里整页抓、再按 type_name 归类，覆盖面比按 t 抓更全。"""
    sep = "&" if "?" in api else "?"
    data = get_json(f"{api}{sep}ac=detail&pg={page}")
    if not data:
        return []
    out = []
    for it in (data.get("list") or []):
        type_name = str(it.get("type_name") or "").strip()
        cate = match_cate(type_name)
        if not cate:
            continue
        out.append({
            "site": site_name,
            "cate": cate,
            "name": str(it.get("vod_name") or "").strip(),
            "year": str(it.get("vod_year") or "").strip(),
            "remarks": str(it.get("vod_remarks") or type_name or "").strip()[:30],
            "type_name": type_name,
        })
    return out


def run_maccms(rule: dict, limit_total: int) -> list:
    api = rule["api"].strip()
    name = rule.get("name", api)
    pages = int(rule.get("pages", 6))
    per_cate = int(rule.get("limit_per_category", 30))
    print(f"[maccms] {name} → {api}")
    classes = maccms_classes(api)
    if not classes:
        print("  分类接口没返回内容，跳过")
        return []
    print("  站点分类：", [c["name"] for c in classes][:10])
    rows, counts = [], {}
    for page in range(1, pages + 1):
        try:
            items = maccms_page(api, page, name)
        except Exception as exc:  # noqa: BLE001
            print(f"  [warn] 第 {page} 页抓取失败：{type(exc).__name__}")
            continue
        print(f"  第 {page} 页 → {len(items)} 条有效")
        for it in items:
            if counts.get(it["cate"], 0) >= per_cate:
                continue
            rows.append(it)
            counts[it["cate"]] = counts.get(it["cate"], 0) + 1
            if len(rows) >= limit_total:
                return rows
        if all(counts.get(c, 0) >= per_cate for c in TARGET_CATES):
            break
        time.sleep(0.15)
    print("  分类分布：", counts)
    return rows


# ------------------------------------------------------------ HTML 规则模式
def run_html(rule: dict, limit_total: int) -> list:
    name = rule.get("name", "")
    pages = int(rule.get("pages", 1))
    item_re = re.compile(rule["item_regex"], re.S | re.I)
    year_re = re.compile(rule.get("year_regex", r"(\d{4})"), re.S)
    remark_re = re.compile(rule["remark_regex"], re.S) if rule.get("remark_regex") else None
    rows = []
    for cate, tpl in (rule.get("list_urls") or {}).items():
        for page in range(1, pages + 1):
            url = tpl.replace("{page}", str(page))
            try:
                html = get_html(url)
            except Exception as exc:  # noqa: BLE001
                print(f"  [warn] {url} 抓取失败：{exc}")
                continue
            hits = item_re.findall(html)
            print(f"[html] {url} → {len(hits)} 条")
            for hit in hits:
                link, title = hit[0], hit[1]
                title = re.sub(r"<[^>]+>|&nbsp;|&amp;", " ", title).strip()
                if not title:
                    continue
                full = "".join(hit)
                rows.append({
                    "site": name, "cate": cate, "name": title,
                    "year": (year_re.search(full).group(1) if year_re.search(full) else ""),
                    "remarks": (re.sub(r"<[^>]+>", "", remark_re.search(full).group(1)).strip()[:30]
                                if remark_re else ""),
                    "link": urllib.parse.urljoin(url, link),
                })
                if len(rows) >= limit_total:
                    return rows
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--rule", type=int, default=0, help="只跑第 N 条规则（从 1 开始）")
    ap.add_argument("--check", action="store_true", help="只测接口")
    ap.add_argument("--limit", type=int, default=120)
    args = ap.parse_args()

    OUT.mkdir(parents=True, exist_ok=True)
    rules = json.loads((BASE / "site_rules.json").read_text(encoding="utf-8"))["rules"]
    rules = [r for r in rules if r.get("enabled", True)]
    if args.rule:
        rules = [rules[args.rule - 1]]

    rows = []
    for rule in rules:
        kind = rule.get("kind")
        try:
            if kind == "maccms":
                if args.check:
                    cls = maccms_classes(rule["api"])
                    print(f"{rule.get('name')} 分类 {len(cls)} 个：{[c['name'] for c in cls][:8]}")
                    continue
                rows.extend(run_maccms(rule, args.limit))
            elif kind == "html":
                rows.extend(run_html(rule, args.limit))
            else:
                print("未知规则类型", kind)
        except Exception as exc:  # noqa: BLE001
            print(f"[error] {rule.get('name')} 抓取失败：{type(exc).__name__}: {exc}")

    if rows:
        (OUT / "自定义站点_条目.json").write_text(
            json.dumps(rows, ensure_ascii=False, indent=1), encoding="utf-8")
        from collections import Counter
        print("共抓到", len(rows), "条，分类分布：", dict(Counter(r["cate"] for r in rows)))
        print("已写入", OUT / "自定义站点_条目.json")


if __name__ == "__main__":
    main()
