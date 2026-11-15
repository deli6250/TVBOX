#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
TVBox / IPTV 源抓取校验器
=========================================
零第三方依赖（除生成 Excel 的那一步用 openpyxl），Windows / macOS / Linux 均可直接跑。

做四件事：
  1) 抓「种子源」：TVBox 接口 JSON、m3u / txt 直播源、聚合导航页
  2) 从聚合页里继续挖出新地址（深度 1 层），形成候选池
  3) 逐个探活：站点调它的分类接口（能拿到 电影/电视剧/综艺/动漫 分类才算活）；
     频道切片头，能拿到 m3u8 清单或 ts 数据才算活
  4) 输出：可用直播源 m3u/txt、合并后的 TVBox 配置、results.json（给 make_excel.py 出清单）

用法：
    python tvbox_crawler.py                # 全量抓取 + 探活
    python tvbox_crawler.py --no-probe     # 只抓不探活，几秒出结果
    python tvbox_crawler.py --only tvbox   # 只跑 TVBox 接口 / live 只跑直播源
    python tvbox_crawler.py --stream-limit 400
"""
from __future__ import annotations

import argparse
import concurrent.futures as cf
import gzip
import json
import re
import ssl
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
OUT_DIR = BASE_DIR / "抓取结果"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36")

SSL_CTX = ssl.create_default_context()
SSL_CTX.check_hostname = False
SSL_CTX.verify_mode = ssl.CERT_NONE

CATE_RULES = [
    ("电影", ("电影", "片", "影院", "高清影视", "影视")),
    ("电视剧", ("连续剧", "电视剧", "剧集", "美剧", "韩剧", "日剧", "国产剧", "港台剧", "泰剧")),
    ("综艺", ("综艺", "娱乐", "真人秀", "脱口秀", "晚会")),
    ("动漫", ("动漫", "动画", "番剧", "国漫", "日漫", "少儿动画")),
]
CATE_KEYS = [c[0] for c in CATE_RULES]

GROUP_RULES = [
    ("央视频道", ("cctv", "央视", "中央", "cgtn", "cetv", "中国教育")),
    ("卫视频道", ("卫视", "satellite")),
    ("影视电影", ("电影", "影院", "影视", "剧场", "movie", "cinema", "star", "golden")),
    ("电视剧", ("电视剧", "剧场", "剧集", "series", "tvb", "drama")),
    ("综艺", ("综艺", "娱乐", "variety", "show")),
    ("动漫少儿", ("动漫", "动画", "少儿", "卡通", "anime", "cartoon", "kids", "优漫", "金鹰卡通")),
    ("体育", ("体育", "足球", "篮球", "赛事", "sport", "nba", "cba")),
    ("港澳台", ("香港", "澳门", "台湾", "翡翠", "明珠", "tvb", "hktv", "澳视", "台视", "中视", "华视", "民视")),
    ("地方频道", ("频道", "新闻综合", "都市", "公共", "经济", "生活", "影视文艺", "县", "市")),
    ("音乐", ("音乐", "music", "ktv")),
    ("国际频道", ("国际", "海外", "usa", "uk", "japan", "korea", "news")),
]


# ---------------------------------------------------------------- 基础工具
def log(msg: str) -> None:
    line = f"[{time.strftime('%H:%M:%S')}] {msg}"
    print(line)
    try:
        with open(OUT_DIR / "抓取日志.txt", "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def http_get(url: str, timeout: int = 15, retries: int = 1, max_bytes: int = 4_000_000) -> dict:
    last_err = ""
    for _ in range(retries + 1):
        try:
            # 地址里可能带中文/空格，urllib 直接请求会抛 UnicodeEncodeError，先做百分号转义
            safe = urllib.parse.quote(url, safe=":/?&=%#[]@!$&'()*+,;~-._")
            req = urllib.request.Request(safe, headers={
                "User-Agent": UA,
                "Accept": "*/*",
                "Accept-Encoding": "gzip",
            })
            with urllib.request.urlopen(req, timeout=timeout, context=SSL_CTX) as resp:
                raw = resp.read(max_bytes)
                if resp.headers.get("Content-Encoding") == "gzip":
                    try:
                        raw = gzip.decompress(raw)
                    except Exception:
                        pass
                return {"ok": True, "status": getattr(resp, "status", 200), "data": raw,
                        "ctype": resp.headers.get("Content-Type", "")}
        except Exception as exc:  # noqa: BLE001
            last_err = f"{type(exc).__name__}: {exc}"[:160]
            time.sleep(0.25)
    return {"ok": False, "status": 0, "data": b"", "ctype": "", "err": last_err}


def to_text(raw: bytes) -> str:
    for enc in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            return raw.decode(enc)
        except Exception:
            continue
    return raw.decode("utf-8", "ignore")


# ---------------------------------------------------------------- JSON 宽容解析
def _strip_js_comments(text: str) -> str:
    out, i, n = [], 0, len(text)
    in_str = esc = False
    while i < n:
        ch = text[i]
        if in_str:
            out.append(ch)
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_str = False
            i += 1
            continue
        if ch == '"':
            in_str = True
            out.append(ch)
            i += 1
            continue
        if ch == "/" and i + 1 < n and text[i + 1] == "/":
            while i < n and text[i] not in "\r\n":
                i += 1
            continue
        if ch == "/" and i + 1 < n and text[i + 1] == "*":
            i += 2
            while i + 1 < n and not (text[i] == "*" and text[i + 1] == "/"):
                i += 1
            i += 2
            continue
        out.append(ch)
        i += 1
    return "".join(out)


def load_config_json(text: str):
    t = text.lstrip("\ufeff").strip()
    t = _strip_js_comments(t)
    t = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", t)
    try:
        return json.loads(t)
    except Exception:
        pass
    start, end = t.find("{"), t.rfind("}")
    if start >= 0 and end > start:
        seg = t[start:end + 1]
        for attempt in (seg, seg.replace(",}", "}").replace(",]", "]")):
            try:
                return json.loads(attempt)
            except Exception:
                continue
    return None


# ---------------------------------------------------------------- 内容识别
def sniff_kind(text: str) -> str:
    stripped = text.lstrip()
    if stripped.startswith("{") or stripped.startswith("["):
        data = load_config_json(text)
        if isinstance(data, (dict, list)):
            return "tvbox"
    if text.count("#EXTINF") >= 2:
        return "live"
    if "#genre#" in text:
        return "live"
    head = stripped[:200].lower()
    if head.startswith("<!doctype") or head.startswith("<html") or "<body" in head:
        return "html"
    if re.search(r"\n[^,\r\n]{1,40},\s*(http|rtp|rtsp|udp)", text):
        return "live"
    return "unknown"


# 成人/赌博类源，一律不抓、不落表
ADULT_RE = re.compile(
    r"18\+|18禁|十八禁|成人|情色|色情|福利|裸[播聊]?|激情|三级|午夜|偷拍|自慰|porn|adult|xxx|"
    r"av\s*片|约炮|楼凤|嫖|裸聊|赌博|博彩|彩票网|六合彩|真人视讯|荷官", re.I)
NOISE_RE = re.compile(r"(高清|超清|蓝光|标清|4K|8K|备用|测试|线路\d*|移动|联通|电信|"
                      r"ipv4|ipv6|咪咕|备份|直播|频道|服务|高清频道)", re.I)
GROUP_PRIORITY = {"央视频道": 0, "卫视频道": 1, "影视电影": 2, "电视剧": 3, "综艺": 4,
                  "动漫少儿": 5, "港澳台": 6, "地方频道": 7, "体育": 8, "音乐": 9,
                  "国际频道": 10, "其他频道": 11}
# 探活名额配额（会按 --stream-limit 等比缩放）；保证 影视/电视剧/综艺/动漫 都有份额
CATEGORY_QUOTA = {"央视频道": 260, "卫视频道": 200, "影视电影": 300, "电视剧": 300,
                  "综艺": 240, "动漫少儿": 300, "港澳台": 120, "地方频道": 140,
                  "体育": 60, "音乐": 40, "国际频道": 30, "其他频道": 40}


def clean_channel_name(name: str) -> str:
    text = re.sub(r"[「」【】\[\]（）()<>《》\u200b\u3000]", "", name)
    text = re.sub(r"[^\w\u4e00-\u9fff\+\-#.· ]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    text = re.sub(r"[-_·.]{2,}", "", text).strip(" -_·.")
    return text


def group_key(name: str) -> str:
    base = NOISE_RE.sub("", name).strip(" -_·")
    return base or name


def parse_m3u(text: str, source: str) -> list:
    chans, attrs, name = [], {}, ""
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if line.upper().startswith("#EXTINF"):
            attrs = dict(re.findall(r'([A-Za-z0-9_-]+)="([^"]*)"', line))
            name = attrs.get("tvg-name") or (line.split(",", 1)[1].strip() if "," in line else "")
            name = re.sub(r"\s+", " ", name).strip()
        elif line.startswith("#"):
            continue
        else:
            if line.lower().startswith(("http://", "https://", "rtp://", "rtsp://", "rtmp://", "udp://")):
                group = attrs.get("group-title") or ""
                chans.append({"name": name or "未命名", "url": line,
                              "group": group.strip(), "source": source,
                              "logo": attrs.get("tvg-logo", "")})
                attrs, name = {}, ""
    return chans


def parse_txt_live(text: str, source: str) -> list:
    chans, group = [], ""
    for raw_line in text.splitlines():
        line = raw_line.strip().lstrip("\ufeff")
        if not line or line.startswith("#"):
            continue
        if "#genre#" in line:
            group = line.split(",")[0].strip().strip("#").strip()
            continue
        if "," not in line:
            continue
        name, url = line.split(",", 1)
        url = url.strip()
        if url.lower().startswith(("http://", "https://", "rtp://", "rtsp://", "rtmp://", "udp://")):
            chans.append({"name": name.strip() or "未命名", "url": url,
                          "group": group, "source": source, "logo": ""})
    return chans


URL_RE = re.compile(r"https?://[^\s\"'<>)\]\\，,；;（）]+", re.I)


def extract_urls(html: str) -> list:
    found = []
    for u in URL_RE.findall(html):
        u = u.rstrip(".,;:)]}'\"")
        low = u.lower()
        if low.endswith((".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".ico", ".css", ".js",
                         ".apk", ".jar", ".zip", ".rar", ".woff", ".ttf")):
            continue
        if ".json" in low or ".m3u" in low or ".m3u8" in low or ".txt" in low:
            found.append(u)
    return list(dict.fromkeys(found))


def classify_category(names: list) -> dict:
    covered = {k: False for k in CATE_KEYS}
    for nm in names:
        low = str(nm).lower()
        for cate, kws in CATE_RULES:
            if any(kw in low for kw in kws):
                covered[cate] = True
    return covered


def classify_group(name: str, group: str) -> str:
    text = f"{group} {name}".lower()
    for grp, kws in GROUP_RULES:
        if any(kw in text for kw in kws):
            return grp
    return "其他频道"


# ---------------------------------------------------------------- 探活
def probe_site(site: dict, timeout: int = 10) -> dict:
    api = (site.get("api") or "").strip()
    stype = site.get("type")
    result = dict(site)
    result.update({"alive": None, "ms": None, "classes": [], "covered": {k: False for k in CATE_KEYS},
                   "error": ""})
    if not api.lower().startswith(("http://", "https://")):
        result["alive"] = None
        result["error"] = "接口为相对路径/内置脚本，需客户端才能跑"
        return result
    if stype not in (0, 1, "0", "1", 4):
        result["error"] = f"type={stype}，非标准采集接口，跳过探活"
        return result

    probe_url = api + ("&" if "?" in api else "?") + "ac=class"
    t0 = time.time()
    resp = http_get(probe_url, timeout=timeout, retries=0, max_bytes=300_000)
    result["ms"] = int((time.time() - t0) * 1000)
    if not resp["ok"]:
        result["alive"] = False
        result["error"] = resp.get("err", "请求失败")
        return result

    text = to_text(resp["data"])
    names = []
    try:
        data = json.loads(text)
        for item in (data.get("class") or []):
            names.append(item.get("type_name") or item.get("name") or "")
    except Exception:
        names = re.findall(r"<ty[^>]*>([^<]{1,40})</ty>", text)
        if not names:
            names = re.findall(r'name="([^"]{1,40})"', text)
    names = [n.strip() for n in names if n and n.strip()]
    covered = classify_category(names)
    if names and any(covered.values()):
        result["alive"] = True
        result["classes"] = names[:40]
        result["covered"] = covered
    elif names:
        # 有返回但分类里没有 电影/电视剧/综艺/动漫，可能是假接口或非影视站
        result["alive"] = False
        result["classes"] = names[:40]
        result["error"] = "分类里没有电影/电视剧/综艺/动漫，判为无效"
    else:
        result["alive"] = False
        result["error"] = f"没有分类列表（HTTP {resp['status']}）"
    return result


def probe_stream(url: str, timeout: int = 7) -> dict:
    t0 = time.time()
    resp = http_get(url, timeout=timeout, retries=0, max_bytes=8192)
    ms = int((time.time() - t0) * 1000)
    if not resp["ok"]:
        return {"alive": False, "ms": ms, "kind": "", "error": resp.get("err", "请求失败")}
    data = resp["data"]
    ctype = (resp.get("ctype") or "").lower()
    head = data[:4096]
    if b"#EXTM3U" in head or b"#EXTINF" in head:
        return {"alive": True, "ms": ms, "kind": "m3u8", "error": ""}
    if "mpegurl" in ctype:
        return {"alive": True, "ms": ms, "kind": "m3u8", "error": ""}
    if head[:1] == b"\x47" or "mp2t" in ctype or "video/" in ctype:
        return {"alive": True, "ms": ms, "kind": "ts", "error": ""}
    if resp["status"] == 200 and len(data) > 0:
        return {"alive": True, "ms": ms, "kind": "other", "error": ""}
    return {"alive": False, "ms": ms, "kind": "", "error": "空响应"}


# ---------------------------------------------------------------- 抓取主流程
class Crawler:
    def __init__(self, limits: dict):
        self.limits = limits
        self.fetched = {}          # url -> 状态
        self.sites = {}            # api|name -> site dict
        self.channels = []         # 原始频道
        self.live_urls = []        # (url, tag, source_page)
        self.stats = {"seeds": 0, "fetched_ok": 0, "fetched_fail": 0, "discovered": 0,
                      "configs": 0, "channels_raw": 0, "blocked_sources": 0}
        self.source_meta = {}      # 直播源文件 -> {channels, tag}

    # -- 抓一个地址并按内容分流
    def fetch_one(self, url: str, tag: str, depth: int, kind_hint: str, discovered_from: str = ""):
        if url in self.fetched:
            return
        # 成人 / 赌博类源整体跳过，不抓也不落表
        if ADULT_RE.search(urllib.parse.unquote(url)):
            self.fetched[url] = {"ok": False, "kind": "blocked", "err": "成人/赌博类源，已跳过"}
            self.stats["blocked_sources"] += 1
            return
        self.fetched[url] = "queued"
        resp = http_get(url, timeout=self.limits["config_timeout"], retries=2)
        if not resp["ok"]:
            self.fetched[url] = {"ok": False, "err": resp.get("err")}
            self.stats["fetched_fail"] += 1
            return
        text = to_text(resp["data"])
        kind = "html" if kind_hint == "html" else sniff_kind(text)
        if kind_hint == "live" and kind == "unknown":
            kind = "live"
        if kind_hint == "tvbox" and kind == "unknown":
            kind = "tvbox"
        self.fetched[url] = {"ok": True, "kind": kind, "size": len(resp["data"])}
        self.stats["fetched_ok"] += 1

        if kind == "live":
            chans = parse_m3u(text, tag) if "#EXTINF" in text else parse_txt_live(text, tag)
            if not chans:
                chans = parse_txt_live(text, tag)
            chans = [c for c in chans
                     if not (ADULT_RE.search(c["name"]) or ADULT_RE.search(c["group"])
                             or ADULT_RE.search(urllib.parse.unquote(c["url"])))]
            self.channels.extend(chans)
            self.source_meta[url] = {"channels": len(chans), "tag": tag}
            self.stats["channels_raw"] += len(chans)
            log(f"直播源 {url[:70]} → {len(chans)} 个频道")
        elif kind == "tvbox":
            data = load_config_json(text)
            if isinstance(data, dict):
                self.stats["configs"] += 1
                self._absorb_config(data, url, tag, depth)
        elif kind == "html":
            urls = extract_urls(text)
            self.stats["discovered"] += len(urls)
            if depth < self.limits["max_depth"]:
                for u in urls[:400]:
                    self.live_urls.append((u, tag, url))
            log(f"聚合页 {url[:70]} → 挖到 {len(urls)} 个候选地址")

    def _absorb_config(self, data: dict, url: str, tag: str, depth: int):
        for s in (data.get("sites") or []):
            if not isinstance(s, dict):
                continue
            api = str(s.get("api") or "").strip()
            if not api:
                continue
            key = f"{api}||{s.get('name')}"
            if key not in self.sites:
                self.sites[key] = {
                    "key": s.get("key") or "",
                    "name": str(s.get("name") or "").strip() or "未命名",
                    "type": s.get("type"),
                    "api": api,
                    "searchable": s.get("searchable", 0),
                    "filterable": s.get("filterable", 0),
                    "from": tag,
                    "from_url": url,
                }
        for lv in (data.get("lives") or []):
            if not isinstance(lv, dict):
                continue
            lurl = str(lv.get("url") or "").strip()
            if lurl.lower().startswith(("http://", "https://")) and depth < self.limits["max_depth"]:
                self.live_urls.append((lurl, tag, url))
        # 多仓配置：urls 里还有下一层配置
        for u in (data.get("urls") or []):
            if isinstance(u, dict):
                uu = str(u.get("url") or "").strip()
                if uu.lower().startswith(("http://", "https://")) and depth < self.limits["max_depth"]:
                    self.live_urls.append((uu, tag, url))

    # -- 抓种子
    def run(self, seeds: list, only: str):
        log(f"开始抓取，种子 {len(seeds)} 个，深度上限 {self.limits['max_depth']}")
        want = [s for s in seeds if s.get("enabled", True)]
        if only != "all":
            want = [s for s in want if s.get("type") == only]
        self.stats["seeds"] = len(want)
        with cf.ThreadPoolExecutor(max_workers=self.limits["config_workers"]) as pool:
            list(pool.map(lambda s: self.fetch_one(s["url"], s.get("tag", ""), 0, s.get("type", "auto"),
                                                  discovered_from="seed"), want))

        # 深度 1：抓聚合页 / 配置里挖出来的地址
        uniq, seen = [], set()
        for u, tag, src in self.live_urls:
            if u in seen or u in self.fetched:
                continue
            seen.add(u)
            uniq.append((u, tag, src))
        uniq = uniq[: self.limits["max_urls_total"]]
        log(f"第二轮候选地址 {len(uniq)} 个（上限 {self.limits['max_urls_total']}）")
        if uniq:
            with cf.ThreadPoolExecutor(max_workers=self.limits["config_workers"]) as pool:
                list(pool.map(lambda t: self.fetch_one(t[0], t[1], 1, "auto", t[2]), uniq))
        log(f"抓取完成：成功 {self.stats['fetched_ok']}，失败 {self.stats['fetched_fail']}，"
            f"配置 {self.stats['configs']} 个，原始频道 {self.stats['channels_raw']} 个，"
            f"站点 {len(self.sites)} 个")

    # -- 频道筛选：按分类配额分配探活名额，同名频道保留多条不同来源线路
    @staticmethod
    def select_channels(chans: list, max_out: int, per_channel: int = 3) -> list:
        grouped = {}
        for c in chans:
            name = clean_channel_name(c["name"])
            if len(name) < 1:
                continue
            cate = classify_group(name, c["group"])
            key = group_key(name)
            c.update({"clean_name": key or name, "cate": cate})
            grouped.setdefault((cate, key or name), []).append(c)

        total_quota = sum(CATEGORY_QUOTA.values())
        scale = max_out / total_quota
        budget = {k: max(15, int(v * scale)) for k, v in CATEGORY_QUOTA.items()}
        ordered = sorted(grouped.items(),
                         key=lambda kv: (GROUP_PRIORITY.get(kv[0][0], 99), -len(kv[1]), kv[0][1]))
        out, used = [], {}
        for (cate, _key), items in ordered:
            if used.get(cate, 0) >= budget.get(cate, 15):
                continue
            room = min(per_channel, budget.get(cate, 15) - used.get(cate, 0))
            chosen, used_src, used_url = [], set(), set()
            # 第一轮：每个来源只出一条，保证线路不集中在同一个源
            for it in sorted(items, key=lambda x: (len(x["url"]), x["url"])):
                if len(chosen) >= room:
                    break
                if it["source"] in used_src or it["url"] in used_url:
                    continue
                chosen.append(it)
                used_src.add(it["source"])
                used_url.add(it["url"])
            # 第二轮：名额没填满就用同源的备选线路补上（单源频道也能有多线路）
            for it in items:
                if len(chosen) >= room:
                    break
                if it["url"] in used_url:
                    continue
                chosen.append(it)
                used_url.add(it["url"])
            out.extend(chosen)
            used[cate] = used.get(cate, 0) + len(chosen)
            if len(out) >= max_out:
                break
        return out[:max_out]

    # -- 探活
    def probe_all(self, site_limit: int, stream_limit: int):
        # 优先探活「纯 HTTP 采集接口」的站点：jar 蜘蛛类型只能靠 TVBox 客户端跑，HTTP 校验不了
        able = [s for s in self.sites.values()
                if str(s.get("api", "")).lower().startswith(("http://", "https://"))]
        able_ids = {id(s) for s in able}
        unable = [s for s in self.sites.values() if id(s) not in able_ids]
        cannot = [dict(s, alive=None, ms=None, covered={k: False for k in CATE_KEYS},
                       classes=[], error="jar/内置脚本类型，需 TVBox 客户端才能运行，HTTP 无法校验")
                  for s in unable]
        sites = able[:site_limit]
        log(f"站点总计 {len(self.sites)} 个，其中可 HTTP 探活 {len(able)} 个，实际探活 {len(sites)} 个 …")
        with cf.ThreadPoolExecutor(max_workers=self.limits["site_workers"]) as pool:
            self.site_results = list(pool.map(lambda s: probe_site(s, self.limits["site_timeout"]), sites))
        self.site_results.extend(cannot)

        total = len(self.channels)
        # 记录筛选前的分类分布（按 url 去重），用于报告对比
        raw_seen, raw_cate = set(), {}
        for c in self.channels:
            if c["url"] in raw_seen:
                continue
            raw_seen.add(c["url"])
            cn = classify_group(clean_channel_name(c["name"]), c["group"])
            raw_cate[cn] = raw_cate.get(cn, 0) + 1
        self.cate_raw = raw_cate
        todo = self.select_channels(self.channels, stream_limit)
        log(f"频道原始 {total} 个（已剔除成人源/重复），按频道挑线路后探活 {len(todo)} 个 …")
        with cf.ThreadPoolExecutor(max_workers=self.limits["stream_workers"]) as pool:
            probes = list(pool.map(lambda c: probe_stream(c["url"], self.limits["stream_timeout"]), todo))
        self.stream_results = []
        for c, p in zip(todo, probes):
            row = dict(c)
            row.update({"alive": p["alive"], "ms": p["ms"], "kind": p["kind"], "error": p["error"]})
            self.stream_results.append(row)
        alive = sum(1 for r in self.stream_results if r["alive"])
        log(f"直播源探活：{alive}/{len(self.stream_results)} 可用")


# ---------------------------------------------------------------- 输出
def write_live_files(rows: list):
    alive = [r for r in rows if r["alive"]]
    m3u_lines = ['#EXTM3U', '#EXT-X-VERSION:3']
    txt_lines = []
    last_group = None
    for r in sorted(alive, key=lambda x: (GROUP_PRIORITY.get(x["cate"], 99), x["name"])):
        grp = r["cate"]
        m3u_lines.append(f'#EXTINF:-1 tvg-name="{r["name"]}" group-title="{grp}",{r["name"]}')
        m3u_lines.append(r["url"])
        if grp != last_group:
            txt_lines.append(f"{grp},#genre#")
            last_group = grp
        txt_lines.append(f'{r["name"]},{r["url"]}')
    (OUT_DIR / "直播源_可用.m3u").write_text("\n".join(m3u_lines) + "\n", encoding="utf-8")
    (OUT_DIR / "直播源_可用.txt").write_text("\n".join(txt_lines) + "\n", encoding="utf-8")


def write_tvbox_config(site_rows: list):
    good = [s for s in site_rows if s.get("alive") and any(s["covered"].values())]
    sites = []
    for i, s in enumerate(good, 1):
        entry = {
            "key": f"src{i:03d}",
            "name": s["name"][:24],
            "type": s["type"],
            "api": s["api"],
            "searchable": 1 if s.get("searchable") else 0,
            "quickSearch": 1 if s.get("searchable") else 0,
            "filterable": 1 if s.get("filterable") else 0,
        }
        sites.append(entry)
    cfg = {
        "spider": "",
        "wallpaper": "",
        "sites": sites,
        "lives": [
            {
                "name": "校验通过的直播源",
                "type": 0,
                "url": "http://<把 直播源_可用.txt 放到局域网/网盘后填这里>/直播源_可用.txt",
                "epg": "http://epg.112114.xyz/?ch={name}&date={date}",
                "playerType": 1,
            }
        ],
        "parses": [],
        "flags": ["youku", "qq", "iqiyi", "qiyi", "letv", "sohu", "tudou", "pptv", "mgtv", "wasu"],
        "_说明": "由 tvbox_crawler.py 生成：只保留探活通过且至少覆盖一个点播分类的站点。lives 里的 url 需要换成你实际可访问的地址。",
    }
    (OUT_DIR / "tvbox_config.json").write_text(
        json.dumps(cfg, ensure_ascii=False, indent=1), encoding="utf-8")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-probe", action="store_true", help="只抓不探活")
    ap.add_argument("--only", default="all", choices=["all", "tvbox", "live", "html"])
    ap.add_argument("--stream-limit", type=int, default=0)
    ap.add_argument("--site-limit", type=int, default=0)
    args = ap.parse_args()

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    seed_file = BASE_DIR / "sources_seed.json"
    cfg = json.loads(seed_file.read_text(encoding="utf-8"))
    limits = cfg["limits"]
    if args.stream_limit:
        limits["max_stream_probe"] = args.stream_limit
    if args.site_limit:
        limits["max_sites_probe"] = args.site_limit

    for f in OUT_DIR.glob("*.txt"):
        if f.name.startswith("抓取日志"):
            f.unlink()

    cr = Crawler(limits)
    t0 = time.time()
    cr.run(cfg["seeds"], args.only)

    if args.no_probe:
        cr.site_results = []
        picked = cr.select_channels(cr.channels, limits["max_stream_probe"])
        cr.stream_results = [dict(c, alive=None, ms=None, kind="", error="未探活") for c in picked]
    else:
        cr.probe_all(limits["max_sites_probe"], limits["max_stream_probe"])

    write_live_files(cr.stream_results)
    write_tvbox_config(cr.site_results or [])

    payload = {
        "generated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "elapsed_sec": round(time.time() - t0, 1),
        "limits": limits,
        "stats": cr.stats,
        "fetch": {k: v for k, v in cr.fetched.items()},
        "source_meta": cr.source_meta,
        "cate_raw": getattr(cr, "cate_raw", {}),
        "sites": cr.site_results,
        "streams": cr.stream_results,
        "cate_counts": {},
    }
    counts = {}
    for r in cr.stream_results:
        if r["alive"]:
            counts[r["cate"]] = counts.get(r["cate"], 0) + 1
    payload["cate_counts"] = counts
    (OUT_DIR / "results.json").write_text(json.dumps(payload, ensure_ascii=False, indent=1), encoding="utf-8")

    alive_stream = sum(1 for r in cr.stream_results if r["alive"])
    alive_site = sum(1 for s in (cr.site_results or []) if s.get("alive"))
    log(f"全部完成，耗时 {payload['elapsed_sec']}s：可用频道 {alive_stream}，可用站点 {alive_site}")
    print("\n分类统计：", json.dumps(counts, ensure_ascii=False))


if __name__ == "__main__":
    main()
