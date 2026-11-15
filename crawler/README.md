# 爬虫代码（crawler/）

本目录是抓取 / 探活 / 生成清单的本地脚本，用于以后重跑刷新源。抓取产生的大数据（`results.json`、xlsx、原始 m3u/txt 明细）不入库。

## 文件
- `tvbox_crawler.py`：主爬虫。抓种子源 → 挖候选地址 → 探活站点与频道 → 输出可用直播源 m3u/txt、TVBox 配置、`results.json`
- `site_crawler.py`：指定影视站点爬虫，按 `site_rules.json` 抓 电影/电视剧/综艺/动漫 条目
- `make_excel.py`：把 `results.json` 生成多子表 Excel 清单（依赖 openpyxl）
- `prep_upload.py`：把探活结果精简成仓库用的 `config.json` / `live.txt`（一次性打包脚本，路径按本机写死，换机器需改 `BASE`）
- `sources_seed.json`：种子源与探活名额配置
- `site_rules.json`：指定站点的抓取规则（苹果CMS 接口 / HTML 正则两种）

## 运行

```powershell
python tvbox_crawler.py              # 全量抓取 + 探活
python tvbox_crawler.py --no-probe   # 只抓不测
python tvbox_crawler.py --only live  # 只跑直播源
python tvbox_crawler.py --stream-limit 4000

python site_crawler.py --check       # 先测站点接口通不通
python site_crawler.py               # 抓指定站点条目

python make_excel.py                 # 生成 Excel 清单
```

## 配置
- `sources_seed.json` 的 `limits`：探活并发、超时、名额上限；`seeds` 里增删种子源。
- `site_rules.json`：苹果CMS 站填接口根地址；HTML 站填分类页模板（`{page}` 为页码）+ 正则（第 1 个括号=详情链接，第 2 个=标题）。

## 说明
- 仅处理公开分享的第三方地址，不破解、不碰付费会员；成人/赌博类源在代码内直接屏蔽。
- 所有结果只代表「当轮能否连上」，不保证长期有效；直播源时效差，建议隔几天重跑。
- 仅个人观看，请勿分发或商用。
