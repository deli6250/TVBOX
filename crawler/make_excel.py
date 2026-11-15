#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""把 抓取结果/results.json 转成 源清单.xlsx（多子表）。"""
from __future__ import annotations

import json
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
OUT = BASE / "抓取结果"
RESULT = OUT / "results.json"

TITLE_FILL = PatternFill("solid", fgColor="3370FF")
HEAD_FILL = PatternFill("solid", fgColor="E1EAFF")
NOTE_FILL = PatternFill("solid", fgColor="F5F7FA")
WHITE_BOLD = Font(name="微软雅黑", bold=True, color="FFFFFF", size=12)
HEAD_FONT = Font(name="微软雅黑", bold=True, size=10, color="133C9A")
BODY_FONT = Font(name="微软雅黑", size=10)
THIN = Side(style="thin", color="B8C7D9")

CATE_ORDER = ["央视频道", "卫视频道", "影视电影", "电视剧", "综艺", "动漫少儿",
              "港澳台", "地方频道", "体育", "音乐", "国际频道", "其他频道"]


def head_row(ws, row, titles):
    for col, t in enumerate(titles, 1):
        cell = ws.cell(row, col, t)
        cell.fill = HEAD_FILL
        cell.font = HEAD_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=THIN)


def title_row(ws, text, span):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=span)
    cell = ws.cell(1, 1, text)
    cell.fill = TITLE_FILL
    cell.font = WHITE_BOLD
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26


def widths(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[get_column_letter(col)].width = w


def main():
    data = json.loads(RESULT.read_text(encoding="utf-8"))
    streams = data.get("streams", [])
    sites = data.get("sites", [])
    alive_streams = [s for s in streams if s.get("alive")]
    alive_sites = [s for s in sites if s.get("alive") and any((s.get("covered") or {}).values())]
    stats = data.get("stats", {})

    wb = Workbook()

    # ---------------------------------------------------------- 概览
    ws = wb.active
    ws.title = "概览"
    title_row(ws, "TVBox / IPTV 源抓取报告", 4)
    rows = [
        ("生成时间", data.get("generated_at", "")),
        ("抓取耗时", f"{data.get('elapsed_sec', 0)} 秒"),
        ("抓取源数量（种子）", stats.get("seeds", 0)),
        ("抓取成功 / 失败", f"{stats.get('fetched_ok', 0)} / {stats.get('fetched_fail', 0)}"),
        ("其中跳过成人/赌博源", stats.get("blocked_sources", 0)),
        ("TVBox 接口配置（含多仓）", stats.get("configs", 0)),
        ("原始频道条数（已剔成人）", stats.get("channels_raw", 0)),
        ("参与探活的频道线路", len(streams)),
        ("可用频道线路", len(alive_streams)),
        ("可用频道数（按名称去重）", len({s["clean_name"] for s in alive_streams})),
        ("探测的影视站点", len(sites)),
        ("可用影视站点（含分类覆盖）", len(alive_sites)),
    ]
    r = 3
    for k, v in rows:
        ws.cell(r, 1, k).font = Font(name="微软雅黑", size=10, bold=True)
        ws.cell(r, 1).alignment = Alignment(vertical="center")
        c = ws.cell(r, 2, v)
        c.font = BODY_FONT
        c.alignment = Alignment(horizontal="left", vertical="center")
        r += 1

    r += 1
    ws.cell(r, 1, "分类可用情况").font = HEAD_FONT
    r += 1
    cnt = defaultdict(lambda: [0, 0])
    for s in streams:
        key = s.get("cate", "其他频道")
        cnt[key][1] += 1
        if s.get("alive"):
            cnt[key][0] += 1
    for cate in CATE_ORDER:
        if cate not in cnt:
            continue
        ws.cell(r, 1, cate).font = BODY_FONT
        ws.cell(r, 2, f"{cnt[cate][0]} 可用 / {cnt[cate][1]} 探测").font = BODY_FONT
        r += 1

    r += 1
    notes = [
        "说明：",
        "1) 所有源均为网络上公开分享的第三方地址，本工具只做抓取、去重和可用性校验，不存储、不上传任何视频内容。",
        "2) 站点“可用”的判断标准是：能调通它的分类接口并返回 电影/电视剧/综艺/动漫 中的至少一类。",
        "3) 频道“可用”的判断标准是：能取到 m3u8 清单或视频切片数据；部分频道要求特定 UA/Referer，可能被误判为不可用。",
        "4) 成人/赌博类源已整体排除，不会出现在清单里。",
        "5) 请只用于个人合法观看；付费内容请通过正规渠道订阅，勿用于传播或商业用途。",
        "6) 直播源时效性强，过几天部分地址会失效，重新运行 tvbox_crawler.py 即可刷新。",
    ]
    for note in notes:
        cell = ws.cell(r, 1, note)
        cell.font = Font(name="微软雅黑", size=9, color="555555")
        cell.alignment = Alignment(vertical="center")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        if note == "说明：":
            cell.font = Font(name="微软雅黑", size=10, bold=True, color="133C9A")
        r += 1
    widths(ws, {1: 26, 2: 30, 3: 16, 4: 16})

    # ---------------------------------------------------------- 直播源清单
    ws2 = wb.create_sheet("直播源清单")
    titles = ["序号", "分类", "频道名", "线路", "来源文件", "类型", "响应ms", "播放地址"]
    head_row(ws2, 1, titles)
    ordered = sorted(alive_streams, key=lambda s: (
        CATE_ORDER.index(s["cate"]) if s["cate"] in CATE_ORDER else 99,
        s["clean_name"], s.get("ms") or 99999))
    line_seq = defaultdict(int)
    for i, s in enumerate(ordered, 1):
        line_seq[s["clean_name"]] += 1
        vals = [i, s["cate"], s["clean_name"], f"线路{line_seq[s['clean_name']]}",
                s.get("source", "")[-46:], s.get("kind", ""), s.get("ms", ""), s["url"]]
        for col, v in enumerate(vals, 1):
            cell = ws2.cell(i + 1, col, v)
            cell.font = BODY_FONT
            cell.alignment = Alignment(vertical="center",
                                       horizontal="center" if col in (1, 2, 4, 6, 7) else "left",
                                       wrap_text=col == 8)
    widths(ws2, {1: 6, 2: 12, 3: 20, 4: 8, 5: 34, 6: 8, 7: 9, 8: 62})
    ws2.freeze_panes = "A2"

    # ---------------------------------------------------------- 影视站点
    ws3 = wb.create_sheet("影视站点")
    titles = ["序号", "站点名", "类型", "电影", "电视剧", "综艺", "动漫", "响应ms", "接口地址", "来源配置"]
    head_row(ws3, 1, titles)
    ordered_sites = sorted(sites, key=lambda s: (0 if s.get("alive") else 1, s.get("ms") or 99999))
    shown = ordered_sites[:1200]
    for i, s in enumerate(shown, 1):
        cov = s.get("covered") or {}
        vals = [i, s.get("name", "")[:24], s.get("type", ""),
                "可用" if cov.get("电影") else "", "可用" if cov.get("电视剧") else "",
                "可用" if cov.get("综艺") else "", "可用" if cov.get("动漫") else "",
                s.get("ms", ""), s.get("api", ""), s.get("from", "")]
        for col, v in enumerate(vals, 1):
            cell = ws3.cell(i + 1, col, v)
            cell.font = BODY_FONT
            cell.alignment = Alignment(vertical="center",
                                       horizontal="left" if col in (9, 10) else "center",
                                       wrap_text=col in (9, 10))
        if not s.get("alive"):
            ws3.cell(i + 1, 2).font = Font(name="微软雅黑", size=10, color="999999")
    widths(ws3, {1: 6, 2: 22, 3: 6, 4: 7, 5: 9, 6: 7, 7: 7, 8: 9, 9: 50, 10: 14})
    ws3.freeze_panes = "A2"

    # ---------------------------------------------------------- 直播源文件
    ws4 = wb.create_sheet("直播源文件")
    head_row(ws4, 1, ["序号", "来源地址", "频道条数", "抓取状态"])
    meta = data.get("source_meta", {})
    fetch = data.get("fetch", {})
    items = sorted(meta.items(), key=lambda kv: -kv[1].get("channels", 0))
    for i, (url, m) in enumerate(items, 1):
        st = fetch.get(url, {})
        ok = "成功" if isinstance(st, dict) and st.get("ok") else "失败"
        for col, v in enumerate([i, url, m.get("channels", 0), ok], 1):
            cell = ws4.cell(i + 1, col, v)
            cell.font = BODY_FONT
            cell.alignment = Alignment(vertical="center",
                                       horizontal="left" if col == 2 else "center", wrap_text=col == 2)
    widths(ws4, {1: 6, 2: 78, 3: 10, 4: 10})
    ws4.freeze_panes = "A2"

    # ---------------------------------------------------------- 失败记录
    ws5 = wb.create_sheet("失败记录")
    head_row(ws5, 1, ["序号", "类别", "名称", "地址", "失败原因"])
    bad = []
    for s in streams:
        if not s.get("alive"):
            bad.append(("频道线路", s.get("clean_name", ""), s["url"], s.get("error", "")))
    for s in sites:
        if not s.get("alive"):
            bad.append(("影视站点", s.get("name", ""), s.get("api", ""), s.get("error", "")))
    for url, st in fetch.items():
        if isinstance(st, dict) and not st.get("ok") and st.get("kind") != "blocked":
            bad.append(("源地址", "", url, st.get("err", "")))
    for i, (kind, name, addr, err) in enumerate(bad, 1):
        for col, v in enumerate([i, kind, name, addr, err], 1):
            cell = ws5.cell(i + 1, col, v)
            cell.font = BODY_FONT
            cell.alignment = Alignment(vertical="center",
                                       horizontal="left" if col in (3, 4, 5) else "center",
                                       wrap_text=col in (3, 4, 5))
    widths(ws5, {1: 6, 2: 10, 3: 20, 4: 52, 5: 34})
    ws5.freeze_panes = "A2"

    # ---------------------------------------------------------- 自定义站点条目
    custom = OUT / "自定义站点_条目.json"
    if custom.exists():
        items = json.loads(custom.read_text(encoding="utf-8"))
        ws6 = wb.create_sheet("自定义站点条目")
        head_row(ws6, 1, ["序号", "分类", "片名", "年份", "备注", "来源站点"])
        for i, it in enumerate(items, 1):
            vals = [i, it.get("cate", ""), it.get("name", ""), it.get("year", ""),
                    it.get("remarks", ""), it.get("site", "")]
            for col, v in enumerate(vals, 1):
                cell = ws6.cell(i + 1, col, v)
                cell.font = BODY_FONT
                cell.alignment = Alignment(vertical="center",
                                           horizontal="left" if col in (3, 5, 6) else "center")
        widths(ws6, {1: 6, 2: 10, 3: 30, 4: 8, 5: 22, 6: 16})
        ws6.freeze_panes = "A2"

    out_file = OUT / "源清单.xlsx"
    wb.save(out_file)
    print("已生成", out_file)
    print(f"直播源 {len(alive_streams)} 条线路 / 频道 {len({s['clean_name'] for s in alive_streams})} 个，"
          f"站点 {len(alive_sites)} 个可用")


if __name__ == "__main__":
    main()
