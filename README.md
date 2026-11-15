# TVBOX

TVBox 接口配置与可用直播源（精简版：每个频道只保留探活通过、延迟最低的一条线路）。

## 文件
- `config.json`：TVBox 点播接口配置（83 个站点），导入 TVBox 即可使用
- `live.txt`：直播源（TXT 格式，TVBox 直播导入用）
- `live.m3u`：直播源（M3U 格式，VLC / PotPlayer 等播放器用）

## 说明
- 本仓库为私有：`raw.githubusercontent.com` 直链在仓库外需要登录凭证，TVBox 无法直接拉取。
  要在线导入，把仓库改为公开，或把 `live.txt` 放到局域网 / 网盘后，替换 `config.json` 里 `lives` 的 `url`。
- 直播源会随时间失效，失效后重跑本地爬虫刷新即可。
- 仅用于个人观看，请勿分发或商用。
