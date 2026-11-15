# -*- coding: utf-8 -*-
"""为 TVBOX 私有仓准备精简上传文件：每频道只留最快的一条可用线路。"""
import json, pathlib, collections

BASE = pathlib.Path(r"C:\Users\liulei\Doubao\chats\2026-11-15\new-chat-4")
SRC = BASE / "抓取结果"
OUT = BASE / "TVBOX_上传"
OUT.mkdir(exist_ok=True)

# ---- 读探活结果，每频道取延迟最低的一条；剔除带时效签名、很快失效的线路 ----
d = json.loads((SRC / "results.json").read_text(encoding="utf-8"))
alive = [s for s in d["streams"] if s["alive"]]

# 签名/会话令牌参数：这类链接几小时内就会失效，精简版直接不收录
TOKEN_MARKS = ["msisdn=", "SecurityKey=", "txSecret=", "txTime=", "wsSecret=", "wsTime=",
               "auth_key=", "sig=", "exp=", "assertID=", "Channel_ID=", "ProgramID=",
               "ParentNodeID=", "promotionId=", "mvid=", "mcid=", "playurlVersion=",
               "encrypt=", "client_ip=", "spid=", "netType=", "timestamp=", "mdspid="]

def is_stable(url: str) -> bool:
    u = url.lower()
    return not any(m.lower() in u for m in TOKEN_MARKS)

best = {}
dropped_token = 0
for s in alive:
    if not is_stable(s["url"]):
        dropped_token += 1
        continue
    key = s["clean_name"] or s["name"]
    ms = s.get("ms") or 99999
    if key not in best or ms < (best[key].get("ms") or 99999):
        best[key] = s

cate_order = ["央视频道", "卫视频道", "影视电影", "电视剧", "综艺", "动漫少儿",
              "港澳台", "地方频道", "体育", "音乐", "国际频道", "其他频道"]
by_cate = collections.OrderedDict((c, []) for c in cate_order)
for s in best.values():
    by_cate.setdefault(s["cate"], []).append(s)

# ---- live.txt ----
lines = []
for cate, items in by_cate.items():
    if not items:
        continue
    lines.append(f"{cate},#genre#")
    for s in sorted(items, key=lambda x: (x["clean_name"], x["url"])):
        lines.append(f"{s['clean_name']},{s['url']}")
(OUT / "live.txt").write_text("\n".join(lines) + "\n", encoding="utf-8")

# ---- live.m3u ----
m3u = ["#EXTM3U", "#EXT-X-VERSION:3"]
for cate, items in by_cate.items():
    if not items:
        continue
    for s in sorted(items, key=lambda x: (x["clean_name"], x["url"])):
        logo = s.get("logo") or ""
        m3u.append(f'#EXTINF:-1 tvg-name="{s["clean_name"]}" tvg-logo="{logo}" group-title="{cate}",{s["clean_name"]}')
        m3u.append(s["url"])
(OUT / "live.m3u").write_text("\n".join(m3u) + "\n", encoding="utf-8")

# ---- config.json：直播源地址指向仓内 raw 路径，去掉说明字段 ----
cfg = json.loads((SRC / "tvbox_config.json").read_text(encoding="utf-8"))
if cfg.get("lives"):
    cfg["lives"][0]["url"] = "https://raw.githubusercontent.com/deli6250/TVBOX/main/live.txt"
cfg.pop("_说明", None)
(OUT / "config.json").write_text(json.dumps(cfg, ensure_ascii=False, indent=1), encoding="utf-8")

# ---- README ----
readme = """# TVBOX

TVBox 接口配置与可用直播源（精简版：每个频道只保留探活通过、延迟最低的一条线路；已剔除带时效签名、几小时内失效的线路）。

## 文件
- `config.json`：TVBox 点播接口配置（83 个站点），导入 TVBox 即可使用
- `live.txt`：直播源（TXT 格式，TVBox 直播导入用）
- `live.m3u`：直播源（M3U 格式，VLC / PotPlayer 等播放器用）

## 说明
- 本仓库为私有：`raw.githubusercontent.com` 直链在仓库外需要登录凭证，TVBox 无法直接拉取。
  要在线导入，把仓库改为公开，或把 `live.txt` 放到局域网 / 网盘后，替换 `config.json` 里 `lives` 的 `url`。
- 直播源会随时间失效，失效后重跑本地爬虫刷新即可。
- 仅用于个人观看，请勿分发或商用。
"""
(OUT / "README.md").write_text(readme, encoding="utf-8")

for p in sorted(OUT.iterdir()):
    print(p.name, p.stat().st_size, "bytes")
print("频道数(精简后)", len(best), " 站点数", len(cfg.get("sites", [])), " 剔除签名线路", dropped_token)
