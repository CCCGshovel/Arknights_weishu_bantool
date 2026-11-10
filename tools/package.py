# -*- coding: utf-8 -*-
"""
打包分发版：把单文件工具 + 使用说明.txt 打成一个 zip，输出到 dist/。

用法：
    py -3 tools/package.py

产物（全部在 dist/ 下）：
    卫戍协议卡池_v<版本>.html   —— 干员卡池.html 的版本化副本
    使用说明.txt                —— 面向接收方的纯文本说明（UTF-8 BOM + CRLF，记事本直接可读）
    卫戍协议卡池_v<版本>.zip    —— 上面两个文件的打包

副产物（仓库根目录）：
    index.html                  —— GitHub Pages 入口，每次打包都与 干员卡池.html 同步
"""
import re
import sys
import zipfile
import pathlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
HTML = ROOT / "干员卡池.html"
INDEX = ROOT / "index.html"
OUT = ROOT / "dist"


def sync_index():
    """把 干员卡池.html 原样同步为 index.html（GitHub Pages 只认 index.html）。"""
    data = HTML.read_bytes()
    if INDEX.exists() and INDEX.read_bytes() == data:
        print("入口: index.html 已是最新，无需同步")
        return
    INDEX.write_bytes(data)
    print("入口: index.html 已同步 (%d 字节)" % len(data))


def read_var(src, name):
    m = re.search(r'var\s+' + name + r'\s*=\s*"([^"]*)"', src)
    if not m:
        sys.exit("在 %s 里找不到 %s 的定义" % (HTML.name, name))
    return m.group(1)


MANUAL = """\
============================================================
  卫戍协议 · 干员卡池筛选工具    {APP_VER}
  作者：{AUTHOR}
  适用：明日方舟 {GAME_VER} 更新的《{GAME_MODE}》
============================================================

【这是什么】
一个单独的文件，不需要安装、不需要联网、不需要任何运行环境。
它把 143 名干员的盟约与特质整理成五张表：你输入本局被禁的词条，
它就把本局还能出的卡按盟约和阶数摆给你看。

【怎么打开】
1. 把解压出来的「卫戍协议卡池_{APP_VER}.html」放到任意文件夹（路径里可以有中文）。
2. 双击它。如果双击后打开的是记事本 / 一堆代码，请：
   右键 -> 打开方式 -> 选择 Chrome / Edge / 火狐 / 360 / QQ 浏览器。
3. 需要 2020 年底以后内核的浏览器（Chrome 87+ / Edge 88+ / 火狐 85+）。
   手机打开本地网页很折腾，建议在电脑上看。

【怎么用】
第 1 步  在页面上方的「主盟约」里点 5 个本局被禁的主盟约
第 2 步  在「副盟约」里点 5 个本局被禁的副盟约
第 3 步  点「生成卡池」
点错了再点一次就能取消；点「重置」清空全部选择。
页面顶部会实时显示「可用 N 张 · 持续叠加 M 张」。

【五张表怎么看】
表1  可用干员，横轴 = 没被禁的主盟约，纵轴 = Ⅰ~Ⅵ 阶
表2  可用干员，横轴 = 没被禁的副盟约，纵轴 = Ⅰ~Ⅵ 阶
表3  本局出不了的干员，按「被禁主盟约 × 阶数」摆开
表4  本局出不了的干员，横轴 = 全部 15 个副盟约（被禁的列标「禁」）
     没标「禁」的列里还有名字，说明那人副盟约没禁、但主盟约被全禁，照样出不了
表5  可用干员的特质详情（默认折叠，点一下标题展开）
行底色代表阶数，Ⅰ 浅灰 -> Ⅵ 杏橙，跟官方盟约交叉表的配色一致；
表3 / 表4 用同色的暖灰版，一眼能分出这是「出不了」的表。

【格子里怎么读】
    野鬃（助力）★
  · 野鬃     = 干员名
  · （助力） = 它除本列盟约以外的其他盟约
  · ★        = 特质类型是「持续叠加」（名字加粗 + 星标）
一个格子里有多名干员时，每个干员单独占一行。

【为什么同一名干员会出现在好几列】
判定分两步走，先主盟约，再副盟约：
  第一步 主盟约：身上没有主盟约的，直接进第二步；有主盟约的，
          只要有一个没被禁就能出卡，这时候不再看副盟约。
          主盟约被全禁的，直接出不了，副盟约再干净也没用。
  第二步 副盟约：只有「没有主盟约」的干员才走到这一步，
          副盟约被全禁才出不了。
所以带多个主盟约的干员会在每一列都露一次面（比如复奏会同时出现在
叙拉古和莱塔尼亚两列）。
主盟约被全禁的进表3；身上带副盟约却又出不了的，都进表4
（含副盟约本身没被禁、只被主盟约连带卡住的）；两边都占的，表3 表4 里都有他。

【想查某张卡的特质】
在表1 或表2 里点干员名字，会弹出他的特质（基础 / 精锐）、原型干员等。
按 Esc 或点右上角的 × 关闭。

【记不住就点「说明」】
右上角的「说明」按钮里有同样内容，第一次打开会自动弹一次。

【数据与版本】
数据来源：PRTS 明日方舟中文维基（https://prts.wiki/），
          整理为 卫戍协议_干员数据库.csv 与 卫戍协议_盟约交叉表.xlsx「按阶明细」。
工具版本：{APP_VER}
数据版本：见页面顶部与页脚（生成日期 + 干员总数）。
游戏本体更新后数据可能过期，以游戏内实际为准。

【免责】
个人制作的第三方辅助工具，非官方，与鹰角网络、PRTS 均无关。
数据整理难免有错漏，欢迎反馈。

【常见问题】
Q 双击后文件变成一串代码？
A 说明被记事本关联了。右键 -> 打开方式 -> 选浏览器即可。
Q 换了浏览器，我点的选择题全没了？
A 选择存在浏览器本地，换浏览器或清缓存会丢，重点一遍就行。
Q 下载后文件名后面多了个 (1)？
A 不影响使用，随便改。
Q 能离线用吗？
A 能。整个工具就是这一个文件，断网照常用；只有你想点开数据来源链接时才需要网络。
"""


def main():
    src = HTML.read_text(encoding="utf-8")
    app_ver = read_var(src, "APP_VER")
    author = read_var(src, "APP_AUTHOR")
    game_ver = read_var(src, "GAME_VER")
    game_mode = read_var(src, "GAME_MODE")

    sync_index()

    OUT.mkdir(exist_ok=True)
    base = "卫戍协议卡池_%s" % app_ver
    html_out = OUT / (base + ".html")
    txt_out = OUT / "使用说明.txt"
    zip_out = OUT / (base + ".zip")

    html_out.write_bytes(HTML.read_bytes())

    text = MANUAL.format(
        APP_VER=app_ver, AUTHOR=author, GAME_VER=game_ver, GAME_MODE=game_mode
    )
    txt_out.write_bytes(text.replace("\n", "\r\n").encode("utf-8-sig"))

    for p in (zip_out,):
        if p.exists():
            p.unlink()

    with zipfile.ZipFile(zip_out, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.write(txt_out, arcname="使用说明.txt")
        zf.write(html_out, arcname=html_out.name)

    print("版本: %s" % app_ver)
    print("说明: %s (%d 字节)" % (txt_out.relative_to(ROOT), txt_out.stat().st_size))
    print("页面: %s (%d 字节)" % (html_out.relative_to(ROOT), html_out.stat().st_size))
    print("压缩: %s (%d 字节)" % (zip_out.relative_to(ROOT), zip_out.stat().st_size))
    with zipfile.ZipFile(zip_out) as zf:
        for info in zf.infolist():
            print("  - %s (%d 字节)" % (info.filename, info.file_size))


if __name__ == "__main__":
    main()
