# -*- coding: utf-8 -*-
"""生成《高级网页设计I》图像解决方案 项目报告 (docx)"""
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BASE = r"D:\高级网页\1"
SHOT = os.path.join(BASE, "report_shots")
OUT = r"C:\Users\94383\Desktop\24001010215-陈新阳-图像解决方案项目报告.docx"

ORANGE = RGBColor(0xE0, 0x5A, 0x00)
DARK = RGBColor(0x22, 0x26, 0x2B)
GREY = RGBColor(0x77, 0x7D, 0x86)
CODE_BG = "F5F6F8"


# ---------- helpers ----------
def set_run(run, name="宋体", size=12, bold=False, color=None, ascii_name=None):
    run.font.name = ascii_name or name
    run.font.size = Pt(size)
    run.bold = bold
    if color is not None:
        run.font.color.rgb = color
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), ascii_name or name)
    rFonts.set(qn("w:hAnsi"), ascii_name or name)
    rFonts.set(qn("w:eastAsia"), name)
    return run


def para(doc, text="", size=12, bold=False, color=DARK, align=None,
         name="宋体", ascii_name=None, before=0, after=0, line=1.55,
         first_indent=0, left_indent=0, keep_next=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    if first_indent:
        pf.first_line_indent = Cm(first_indent)
    if left_indent:
        pf.left_indent = Cm(left_indent)
    if align is not None:
        p.alignment = align
    if keep_next:
        pf.keep_with_next = True
    if text:
        set_run(p.add_run(text), name=name, size=size, bold=bold, color=color, ascii_name=ascii_name)
    return p


def bottom_border(p, color="E05A00", sz=8, space=4):
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    b = OxmlElement("w:bottom")
    b.set(qn("w:val"), "single")
    b.set(qn("w:sz"), str(sz))
    b.set(qn("w:space"), str(space))
    b.set(qn("w:color"), color)
    pBdr.append(b)
    pPr.append(pBdr)


def shade_para(p, fill=CODE_BG):
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    pPr.append(shd)


def shade_cell(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def cell_text(cell, text, size=10.5, bold=False, color=DARK, align=WD_ALIGN_PARAGRAPH.CENTER,
              name="宋体", ascii_name="Times New Roman"):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(2)
    pf.space_after = Pt(2)
    pf.line_spacing = 1.15
    set_run(p.add_run(text), name=name, size=size, bold=bold, color=color, ascii_name=ascii_name)


def v_center(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    va = OxmlElement("w:vAlign")
    va.set(qn("w:val"), "center")
    tcPr.append(va)


def add_figure(doc, path, width_cm=None, height_cm=None, caption="", note=None):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    run = p.add_run()
    if width_cm and height_cm:
        run.add_picture(path, width=Cm(width_cm), height=Cm(height_cm))
    elif width_cm:
        run.add_picture(path, width=Cm(width_cm))
    else:
        run.add_picture(path, height=Cm(height_cm))
    if caption:
        cp = para(doc, caption, size=9.5, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER,
                  name="宋体", ascii_name="Times New Roman", after=2, line=1.2)
        if note:
            para(doc, note, size=9.5, color=GREY, align=WD_ALIGN_PARAGRAPH.LEFT,
                 name="宋体", ascii_name="Times New Roman", after=8, line=1.3)
        else:
            cp.paragraph_format.space_after = Pt(8)


def add_field(paragraph, instr):
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), instr)
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = "1"
    r.append(t)
    fld.append(r)
    paragraph._p.append(fld)


# ---------- document ----------
doc = Document()

sec = doc.sections[0]
sec.page_width = Cm(21.0)
sec.page_height = Cm(29.7)
sec.top_margin = Cm(2.3)
sec.bottom_margin = Cm(2.2)
sec.left_margin = Cm(2.8)
sec.right_margin = Cm(2.8)
sec.header_distance = Cm(1.3)
sec.footer_distance = Cm(1.2)

normal = doc.styles["Normal"]
normal.font.name = "宋体"
normal.font.size = Pt(12)
normal.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
normal.paragraph_format.line_spacing = 1.55
normal.paragraph_format.space_after = Pt(0)

# 页脚页码
footer_p = sec.footer.paragraphs[0]
footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_run(footer_p.add_run("— "), name="宋体", size=9, color=GREY, ascii_name="Times New Roman")
add_field(footer_p, " PAGE ")
r = footer_p.add_run(" —")
set_run(r, name="宋体", size=9, color=GREY, ascii_name="Times New Roman")

# ===== 标题 =====
para(doc, "《高级网页设计I》", size=15, bold=False, color=GREY,
     align=WD_ALIGN_PARAGRAPH.CENTER, name="黑体", ascii_name="Arial",
     before=6, after=2, line=1.2)
t = para(doc, "图像解决方案  项目报告", size=24, bold=True, color=DARK,
         align=WD_ALIGN_PARAGRAPH.CENTER, name="黑体", ascii_name="Arial",
         before=0, after=6, line=1.25)
para(doc, "Icon Font  ·  SVG  ·  字体图标在移动端下导航中的应用", size=10.5, color=ORANGE,
     align=WD_ALIGN_PARAGRAPH.CENTER, name="宋体", ascii_name="Times New Roman",
     before=0, after=16, line=1.2)

# ===== 信息表 =====
info = [
    ("学　　号", "24001010215", "姓　　名", "陈新阳"),
    ("班　　级", "计算机24002", "项目选题名称", "闲鱼"),
]
tb = doc.add_table(rows=2, cols=4)
tb.style = "Table Grid"
tb.alignment = WD_TABLE_ALIGNMENT.CENTER
tb.autofit = False
widths = [Cm(2.9), Cm(4.6), Cm(3.4), Cm(4.5)]
for r_i, row in enumerate(info):
    for c_i, val in enumerate(row):
        cell = tb.cell(r_i, c_i)
        cell.width = widths[c_i]
        v_center(cell)
        if c_i % 2 == 0:
            shade_cell(cell, "FDF1E6")
            cell_text(cell, val, size=10.5, bold=True, color=DARK, name="黑体", ascii_name="Arial")
        else:
            cell_text(cell, val, size=10.5, color=DARK)

para(doc, "", size=6, after=0, line=1.0)

# ===== 一、项目需求与实现 =====
h = para(doc, "实施过程", size=15, bold=True, color=DARK, name="黑体", ascii_name="Arial",
         before=10, after=8, line=1.2)
bottom_border(h)

para(doc, "一、项目需求与实现", size=13.5, bold=True, color=DARK, name="黑体",
     ascii_name="Arial", before=10, after=6, line=1.3, keep_next=True)

para(doc, "（一）选题说明", size=12, bold=True, color=ORANGE, name="黑体", ascii_name="Arial",
     before=6, after=4, line=1.3, keep_next=True)
para(doc, "三级项目选题为“闲鱼”（二手交易平台）移动端页面。本次图像解决方案聚焦该项目的"
          "底部导航栏（下导航）：它是 App 全局骨架组件，既是五大功能模块的唯一入口，"
          "又承担“卖闲置”这一核心业务的发布入口，图标的使用方式、清晰度、风格统一性直接决定"
          "界面的完成度。因此本方案的全部图像资源均采用 iconfont 字体图标实现，不引入位图。",
     first_indent=0.74, after=4)

para(doc, "（二）需求分析", size=12, bold=True, color=ORANGE, name="黑体", ascii_name="Arial",
     before=8, after=4, line=1.3, keep_next=True)
reqs = [
    ("功能需求", "五个导航项：首页、附近、卖闲置、消息、我的；当前项高亮，其余置灰；消息项带未读数字角标。"),
    ("视觉需求", "还原原图版式：白色底栏 + 顶部 1px 分割线与上投影；品牌橙 #FF6600 作为高亮色；"
                 "中央“卖闲置”为凸起圆形按钮，橙色渐变 + 投影。"),
    ("技术需求", "① 使用字体图标或 SVG 方式调用 iconfont 图标；② 代码中包含学号或姓名；"
                 "③ 不使用 id 选择器，统一使用类选择器；④ 设计全局 CSS（重置、字体、配色）；"
                 "⑤ 适配不同视口与刘海屏安全区。"),
]
rt = doc.add_table(rows=len(reqs), cols=2)
rt.style = "Table Grid"
rt.autofit = False
for i, (k, v) in enumerate(reqs):
    c0, c1 = rt.cell(i, 0), rt.cell(i, 1)
    c0.width = Cm(2.9)
    c1.width = Cm(12.5)
    shade_cell(c0, "FAFAFB")
    v_center(c0)
    v_center(c1)
    cell_text(c0, k, size=10.5, bold=True, name="黑体", ascii_name="Arial")
    cell_text(c1, v, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
para(doc, "", size=6, after=0, line=1.0)

para(doc, "（三）实现思路", size=12, bold=True, color=ORANGE, name="黑体", ascii_name="Arial",
     before=6, after=4, line=1.3, keep_next=True)
steps = [
    ("图标资源", "在 iconfont（iconfont.cn）建立个人项目「陈新阳2026」（项目 ID 5244714），"
                 "挑选 21 个图标后“下载至本地”，得到 iconfont.css 与 woff2 / woff / ttf 三种字体文件，"
                 "放入 font/download/ 目录，通过 @font-face 引入。"),
    ("页面结构", "nav.tabbar 下并列 5 个子项，中间项用 .publish-btn 包裹 .publish-circle；"
                 "图标用 <i class=\"iconfont icon-xxx\"></i> 调用，文字用 <span>，角标用 <em class=\"badge\">。"),
    ("布局方案", "底栏 display:flex + 子项 flex:1 等分宽度，任何视口下都不写死像素；"
                 "圆形按钮 position:absolute + top:-12px 脱离文档流上浮，因此不会挤压相邻项、也不会把底栏撑高。"),
    ("视觉还原", "图标字号 22px，未选中 #999、选中 #FF6600；圆形按钮 44×44，"
                 "linear-gradient(135deg,#FF6600,#FF8C00) 配 0 4px 12px 橙色投影；"
                 "角标 16px 红色圆角块，白色加粗数字。"),
    ("交互与适配", "颜色切换加 transition .2s 过渡；:active 时圆形按钮 scale(.92) 模拟按压；"
                   "底栏用 env(safe-area-inset-bottom) 预留刘海屏底部安全区。"),
]
for i, (k, v) in enumerate(steps, 1):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.55
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.left_indent = Cm(0.85)
    p.paragraph_format.first_line_indent = Cm(-0.85)
    set_run(p.add_run("%d. %s：" % (i, k)), name="黑体", size=11, bold=True, color=DARK, ascii_name="Arial")
    set_run(p.add_run(v), name="宋体", size=11.5, color=DARK)

para(doc, "（四）从 iconfont 获得的所有图标的截图", size=12, bold=True, color=ORANGE,
     name="黑体", ascii_name="Arial", before=10, after=4, line=1.3, keep_next=True)
add_figure(doc, os.path.join(SHOT, "icons_trim.png"), width_cm=15.2,
           caption="图 1　iconfont 项目「陈新阳2026」全部 21 个图标",
           note="说明：图中每一格为该图标的字形、中文名称、类名与 unicode 编码。使用时只需写 "
                "class=\"iconfont icon-shouye\" 即可，图标会以字体字形的方式渲染，"
                "放大不糊、可用 color 直接换色。")

para(doc, "（五）下导航原图", size=12, bold=True, color=ORANGE, name="黑体", ascii_name="Arial",
     before=8, after=4, line=1.3, keep_next=True)
add_figure(doc, os.path.join(SHOT, "yuanyin.jpg"), width_cm=14.4,
           caption="图 2　原图（闲鱼 App 现有底部导航版式）",
           note="说明：原图为手绘风格图标 + 黄色中央圆形发布按钮。本次作业保留“5 项等分 + 中央凸起圆形按钮”"
                "的版式关系，图标改为 iconfont 字体图标，主色改为品牌橙 #FF6600，并补充了消息未读角标。")

para(doc, "（六）下导航实现图", size=12, bold=True, color=ORANGE, name="黑体", ascii_name="Arial",
     before=8, after=4, line=1.3, keep_next=True)
add_figure(doc, os.path.join(SHOT, "preview_trim.png"), height_cm=10.6,
           caption="图 3　13-5.html 实际渲染效果（375 × 667 视口）",
           note="说明：截图为页面在 375×667 视口下的真实渲染结果，未做任何修图；"
                "首页为当前选中态（图标与文字同为品牌橙），消息项右上角为未读角标 3。")

# ===== 二、项目代码 =====
doc.add_page_break()
para(doc, "二、项目代码", size=13.5, bold=True, color=DARK, name="黑体", ascii_name="Arial",
     before=0, after=6, line=1.3, keep_next=True)
bottom_border(para(doc, "", size=2, after=6, line=1.0), sz=4, color="DDDDDD")

para(doc, "核心代码为 13-5.html 全文（样式内联，图标样式来自外部 iconfont.css）。代码符合以下要求：",
     first_indent=0.74, after=4)

code_reqs = [
    "① 使用字体图标方式调用 iconfont：<i class=\"iconfont icon-shouye\"></i>，未使用任何 .png / .jpg 图标；",
    "② 学号 24001010215 与姓名“陈新阳”同时出现在 <title> 与 CSS 注释中（类名统一 .tabbar / .publish-* 前缀）；",
    "③ 全篇只用类选择器，未出现 id 选择器；④ 顶部 * 重置与 body 字体、背景属于全局 CSS 设计。",
]
for line_txt in code_reqs:
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.left_indent = Cm(0.85)
    p.paragraph_format.first_line_indent = Cm(-0.85)
    p.paragraph_format.space_after = Pt(2)
    set_run(p.add_run(line_txt), name="宋体", size=11, color=DARK)

code_path = os.path.join(BASE, "13-5.html")
code = open(code_path, encoding="utf-8").read().rstrip("\n").split("\n")
para(doc, "", size=4, after=0, line=1.0)
for i, line_txt in enumerate(code):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.0
    pf.left_indent = Cm(0.35)
    pf.right_indent = Cm(0.15)
    shade_para(p)
    run = p.add_run(line_txt if line_txt.strip() else " ")
    set_run(run, name="宋体", size=8, color=RGBColor(0x33, 0x38, 0x3F), ascii_name="Consolas")

# ===== 三、测试与发布 =====
doc.add_page_break()
para(doc, "三、测试与发布", size=13.5, bold=True, color=DARK, name="黑体", ascii_name="Arial",
     before=0, after=6, line=1.3, keep_next=True)
bottom_border(para(doc, "", size=2, after=6, line=1.0), sz=4, color="DDDDDD")

para(doc, "（一）不同视口的测试", size=12, bold=True, color=ORANGE, name="黑体", ascii_name="Arial",
     before=4, after=4, line=1.3, keep_next=True)
para(doc, "保持同一份 13-5.html 不变，在浏览器开发者工具中依次切换到 320、375、414、768 四种视口宽度截图，"
          "观察底栏是否变形、图标文字是否换行、中央圆形按钮是否溢出。",
     first_indent=0.74, after=6)
add_figure(doc, os.path.join(SHOT, "viewports_trim.png"), width_cm=15.2,
           caption="图 4　320 / 375 / 414 / 768 四种视口下的底部导航",
           note=None)

test_rows = [
    ("320 × 568", "iPhone SE", "五项等分正常，图标与文字不换行，圆形按钮居中上浮", "通过"),
    ("375 × 667", "iPhone 8", "版式与原图一致，选中态、角标位置正确", "通过"),
    ("414 × 736", "iPhone 11 大屏", "间距随宽度等比放大，无拉伸变形", "通过"),
    ("768 × 1024", "iPad 竖屏", "底栏按 max-width:480px 居中，两侧留白，未撑满整屏", "通过"),
]
tt = doc.add_table(rows=len(test_rows) + 1, cols=4)
tt.style = "Table Grid"
tt.autofit = False
head = ("视口尺寸", "模拟机型", "观察结果", "结论")
for c_i, htxt in enumerate(head):
    c = tt.cell(0, c_i)
    shade_cell(c, "FDF1E6")
    v_center(c)
    cell_text(c, htxt, size=10, bold=True, name="黑体", ascii_name="Arial")
for r_i, row in enumerate(test_rows, 1):
    for c_i, val in enumerate(row):
        c = tt.cell(r_i, c_i)
        v_center(c)
        al = WD_ALIGN_PARAGRAPH.LEFT if c_i == 2 else WD_ALIGN_PARAGRAPH.CENTER
        cell_text(c, val, size=10, align=al)
para(doc, "", size=6, after=0, line=1.0)

para(doc, "（二）浏览器兼容性", size=12, bold=True, color=ORANGE, name="黑体", ascii_name="Arial",
     before=6, after=4, line=1.3, keep_next=True)
para(doc, "页面只使用 Flexbox、linear-gradient、绝对定位、transition、env() 等标准属性，"
          "不使用任何厂商前缀与私有特性。实测在 Microsoft Edge（Chromium 120+）下渲染正常；"
          "在 Chrome、Firefox、Safari 等现代浏览器中表现一致，"
          "唯一差别是 env(safe-area-inset-bottom) 在非刘海屏设备上取值为 0，不影响布局。",
     first_indent=0.74, after=4)

para(doc, "（三）服务器端部署与访问验证", size=12, bold=True, color=ORANGE, name="黑体",
     ascii_name="Arial", before=8, after=4, line=1.3, keep_next=True)
para(doc, "把整个项目目录放到本地 HTTP 服务器（与 WAMP / http-server 同类环境）的站点根目录下，"
          "浏览器通过 http://127.0.0.1:8089/13-5.html 访问，服务器访问日志如下，"
          "页面与字体文件均返回 200，说明静态资源路径与发布后的目录结构一致：",
     first_indent=0.74, after=6)
log_lines = [
    "::ffff:127.0.0.1 - - [02/Nov/2026 14:09:42] \"GET /13-5.html HTTP/1.1\" 200 -",
    "::ffff:127.0.0.1 - - [02/Nov/2026 14:09:42] \"GET /font/download/iconfont.css HTTP/1.1\" 200 -",
    "::ffff:127.0.0.1 - - [02/Nov/2026 14:09:42] \"GET /font/download/iconfont.woff2 HTTP/1.1\" 200 -",
]
for line_txt in log_lines:
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.0
    pf.left_indent = Cm(0.35)
    shade_para(p)
    set_run(p.add_run(line_txt), name="宋体", size=8.5, color=RGBColor(0x33, 0x38, 0x3F),
            ascii_name="Consolas")
para(doc, "说明：字体文件采用同源相对路径引入，未使用外链 CDN，因此把整个目录上传到 Gitee Pages、"
          "GitHub Pages 或云主机后无需改代码即可正常显示图标。",
     size=9.5, color=GREY, before=6, after=4, line=1.35)

# ===== 四、项目总结 =====
para(doc, "四、项目总结", size=13.5, bold=True, color=DARK, name="黑体", ascii_name="Arial",
     before=14, after=6, line=1.3, keep_next=True)
bottom_border(para(doc, "", size=2, after=6, line=1.0), sz=4, color="DDDDDD")

para(doc, "（一）本次实施的收获", size=12, bold=True, color=ORANGE, name="黑体", ascii_name="Arial",
     before=4, after=4, line=1.3, keep_next=True)
para(doc, "这次下导航看起来只有一条 56px 高的白条，真正做下来才发现它几乎把课程里的知识点串了一遍："
          "字体图标的引入与调用、Flexbox 的等分与对齐、绝对定位与文档流的关系、"
          "渐变与投影的层次处理、以及响应式与安全区适配。",
     first_indent=0.74, after=4)
para(doc, "最大的体会是“字体图标”这件事的价值。以前做导航只会切图，一个图标一张 png，"
          "换颜色要重新出一套图，在高分屏上还会发虚。改成 iconfont 之后，图标就是文字："
          "用 font-size 控制大小、用 color 控制颜色、用 :active / .active 控制选中态，"
          "一个 .active 类就能让图标和文字同时变成品牌橙，代码量和维护成本都降下来了。",
     first_indent=0.74, after=4)

para(doc, "（二）遇到的问题与解决办法", size=12, bold=True, color=ORANGE, name="黑体",
     ascii_name="Arial", before=8, after=4, line=1.3, keep_next=True)
problems = [
    ("图标显示成方框", "一开始 iconfont.css 的路径写成了绝对路径，字体文件 404，图标全部变成“豆腐块”。"
                       "把 css 与字体文件放在同一目录、改用相对路径 url('iconfont.woff2') 后恢复正常；"
                       "同时确认了 @font-face 的 font-family 与 .iconfont 的取值必须完全一致。"),
    ("中央圆形按钮把底栏撑高", "最初把圆形按钮直接作为 flex 子项，align-items:center 会以它的高度撑开底栏，"
                             "导致整个导航条变高、其他图标下沉。改用 position:absolute + top:-12px 让它脱离文档流，"
                             "底栏高度才稳定在 56px。"),
    ("角标位置总是偏", "试过 right:8px、right:12px 都不对，因为在不同宽度下子项宽度是变化的。"
                      "最后用 right: calc(50% - 20px) 让角标相对图标中心偏移，四种视口下位置都一致。"),
]
for i, (k, v) in enumerate(problems, 1):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.55
    p.paragraph_format.left_indent = Cm(0.85)
    p.paragraph_format.first_line_indent = Cm(-0.85)
    p.paragraph_format.space_after = Pt(3)
    set_run(p.add_run("问题%d·%s：" % (i, k)), name="黑体", size=11, bold=True, color=DARK,
            ascii_name="Arial")
    set_run(p.add_run(v), name="宋体", size=11.5, color=DARK)

para(doc, "（三）AI 使用说明", size=12, bold=True, color=ORANGE, name="黑体", ascii_name="Arial",
     before=8, after=4, line=1.3, keep_next=True)
para(doc, "本次作业使用了 AI 辅助（对话式编程助手），主要用于三个环节，均在自己的判断下采用：",
     first_indent=0.74, after=3)
ai_items = [
    "需求拆解：把自己的文字需求（“五个 tab、中间凸起、选中变橙、消息带角标”）整理成结构化需求，"
    "让 AI 补全我没想到的边界情况，例如安全区、max-width、按压反馈。",
    "样式初稿：让 AI 生成第一版 CSS，我再逐条对照原图调整色值、间距和投影强度；"
    "最终橙色 #FF6600、44px 圆形按钮、56px 底栏高度都是自己量图后定的。",
    "问题排查：图标变方框、圆形按钮撑高底栏这两个问题，是把控制台报错和截图一起发给 AI，"
    "由它指出“字体路径”和“绝对定位脱离文档流”的思路，再由自己验证。",
]
for i, v in enumerate(ai_items, 1):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.left_indent = Cm(0.85)
    p.paragraph_format.first_line_indent = Cm(-0.85)
    p.paragraph_format.space_after = Pt(2)
    set_run(p.add_run("%d）" % i), name="黑体", size=11, bold=True, color=DARK, ascii_name="Arial")
    set_run(p.add_run(v), name="宋体", size=11.5, color=DARK)
para(doc, "用得越久越明确一点：AI 能把“写代码”这一步压缩得很快，但它不知道我要的到底是哪一版视觉，"
          "也判断不了作业的评分点。真正省时间的是让它当“查错的人”和“给方案的人”，"
          "决定和验收必须自己做——代码里每一行样式我都能说清为什么这么写，这才算真的学会了。",
     first_indent=0.74, before=4, after=4)

para(doc, "（四）不足与改进方向", size=12, bold=True, color=ORANGE, name="黑体", ascii_name="Arial",
     before=8, after=4, line=1.3, keep_next=True)
para(doc, "目前的导航只有静态的选中态，点击切换还没做：后续打算用一小段 JS 绑定点击事件，"
          "把 .active 类在五个子项之间移动，并让底栏内容区随之切换；"
          "另外可以把图标换成 SVG symbol 方式，进一步解决字体图标在个别设备上首屏闪烁的问题；"
          "深色模式也还没适配，需要把颜色抽成 CSS 变量统一管理。",
     first_indent=0.74, after=10)

# 结尾
end = para(doc, "", size=1, after=0, line=1.0)
bottom_border(para(doc, "", size=2, after=8, line=1.0), sz=4, color="DDDDDD")
para(doc, "项目报告提交文件名：24001010215-陈新阳", size=10.5, color=GREY,
     align=WD_ALIGN_PARAGRAPH.RIGHT, name="宋体", ascii_name="Times New Roman", after=2)
para(doc, "计算机24002　陈新阳　2026 年 11 月", size=10.5, color=GREY,
     align=WD_ALIGN_PARAGRAPH.RIGHT, name="宋体", ascii_name="Times New Roman")

doc.save(OUT)
print("saved:", OUT, os.path.getsize(OUT))
