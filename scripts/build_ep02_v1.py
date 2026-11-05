# -*- coding: utf-8 -*-
"""生成 EP02《初入部落》120 秒分镜脚本（8 单元 / 32 镜），重点：说话语气与说话微表情。"""
from docx import Document
from docx.shared import Pt, Mm, Inches, RGBColor
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = r"E:\日常对话\ai-short-drama\docs\EP02-初入部落\EP02_分镜脚本_V1_120秒版.docx"
C_INK, C_GOLD, C_LINE, C_SOFT, C_WHITE = "1C3B4A", "B08A3E", "C3CFD6", "F4F7F8", "FFFFFF"
C_PROMPT = "F2F6F7"


def set_run(run, name="宋体", size=9, bold=False, color=None):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for k in ("w:eastAsia", "w:ascii", "w:hAnsi"):
        rFonts.set(qn(k), name)


def style_font(style, name, size, bold=False):
    style.font.name = name
    style.font.size = Pt(size)
    style.font.bold = bold
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for k in ("w:eastAsia", "w:ascii", "w:hAnsi"):
        rFonts.set(qn(k), name)


def shade(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def para_shade(p, hex_color):
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    pPr.append(shd)


def para_bar(p, color=C_GOLD, sz=18):
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), str(sz))
    left.set(qn("w:space"), "6")
    left.set(qn("w:color"), color)
    pBdr.append(left)
    pPr.append(pBdr)


def table_borders(t, color=C_LINE, sz=4):
    tblPr = t._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for e in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{e}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        b.append(el)
    tblPr.append(b)


def cell_margins(t, top=32, bottom=32, left=56, right=56):
    tblPr = t._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for tag, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        el = OxmlElement(f"w:{tag}")
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.append(mar)


def fixed(t, widths):
    tblPr = t._tbl.tblPr
    lay = OxmlElement("w:tblLayout")
    lay.set(qn("w:type"), "fixed")
    tblPr.append(lay)
    for i, w in enumerate(widths):
        t.columns[i].width = Inches(w)
    for row in t.rows:
        for i, w in enumerate(widths):
            row.cells[i].width = Inches(w)


def no_split(row):
    row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))


def repeat_hdr(row):
    row._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))


def cell_text(cell, lines, name="宋体", size=8, bold=False, color=None, align=0):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    if isinstance(lines, str):
        lines = [lines]
    first = True
    for ln in lines:
        p = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        p.alignment = align
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = 1.1
        set_run(p.add_run(ln), name=name, size=size, bold=bold, color=color)


def para(doc, text="", size=9.5, bold=False, color="2B3B44", before=0, after=4, align=3):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = 1.3
    if text:
        for i, seg in enumerate(text.split("**")):
            if seg:
                set_run(p.add_run(seg), size=size, bold=(bold or i % 2 == 1), color=color)
    return p


def heading(doc, text, level=1, new_page=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.keep_with_next = True
    if new_page:
        pf.page_break_before = True
    if level == 1:
        pf.space_before = Pt(14)
        pf.space_after = Pt(7)
        set_run(p.add_run(text), name="微软雅黑", size=14.5, bold=True, color=C_INK)
        pPr = p._p.get_or_add_pPr()
        b = OxmlElement("w:pBdr")
        bo = OxmlElement("w:bottom")
        bo.set(qn("w:val"), "single")
        bo.set(qn("w:sz"), "8")
        bo.set(qn("w:space"), "4")
        bo.set(qn("w:color"), C_GOLD)
        b.append(bo)
        pPr.append(b)
    else:
        pf.space_before = Pt(7)
        pf.space_after = Pt(3)
        set_run(p.add_run(text), name="微软雅黑", size=11.5, bold=True, color="2F5A6B")
    return p


def prompt_block(doc, sid, dur, lens, subj, cam, light, style):
    head = doc.add_paragraph()
    pf = head.paragraph_format
    pf.space_before = Pt(3)
    pf.space_after = Pt(0)
    pf.left_indent = Pt(8)
    pf.keep_with_next = True
    para_shade(head, C_PROMPT)
    para_bar(head)
    set_run(head.add_run(f"{sid}　｜　{dur}s　｜　{lens}"), name="微软雅黑", size=7.2, bold=True, color=C_GOLD)
    items = (("主体动作", subj), ("镜头", cam), ("光色", light), ("风格", style))
    for i, (label, text) in enumerate(items):
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.left_indent = Pt(8)
        pf.line_spacing = 1.05
        pf.keep_with_next = i < len(items) - 1
        para_shade(p, C_PROMPT)
        para_bar(p)
        set_run(p.add_run(f"{label}："), name="微软雅黑", size=6.9, bold=True, color="2F5A6B")
        set_run(p.add_run(text), name="宋体", size=7.0, color="22323A")


# ============================ 数据：8 单元 / 32 镜 ============================
# 字段：镜号, 时长, 景别·焦段, 运镜, 画面·微表情, 台词·语气, 主体动作, 镜头, 光色, 风格

UNITS = [
    dict(no="P01", time="00:00–00:15", title="茶馆 · 闲话入口",
         func="接住第一章的钩子：一句村民闲话把「大祭司被救」送到他耳边；一句吩咐把下一次见面定下来。",
         emo="松弛 → 留耳 → 收网", hook="他为什么突然要请一个外族女子喝茶？",
         sound="集市底噪、茶馆门帘、瓷杯轻碰、古琴泛音", trans="斜阳街景 → 包房门帘（硬切）",
         shots=[
             ("P01-S01", 4, "大远 ｜ 24mm", "航拍低空横移（侯孝贤式市井留白）",
              "斜阳铺满长街，收摊前的热闹；两人穿行，墨蓝衣角扫过摊位，人群自动让开半步。",
              "无",
              "斜阳下的古朴集市长街，摊贩陆续收摊，暖金光斑落在地面；一名极高个子的墨蓝锦袍男子与玄色劲装护卫并肩穿行，人群自然让开。",
              "24mm 大远景，低空横移跟拍，前景有摊位与人群掠过，人物始终在右三分线。",
              "斜阳 3800K 侧逆光，暖金高光与长影，空气中可见细尘与炊烟。",
              "古装市井写实，胶片颗粒，参考《知否》市集调度；禁止现代元素。"),
             ("P01-S02", 4, "中近 ｜ 65mm", "跟拍 + 焦点转移（迎接）",
              "青云嫂子快步迎来，双手在围裙上擦了一下才欠身；眉眼恭敬，笑意只到唇角。",
              "青云嫂子（恭谨、音量压低、尾音收）：「少主，您往里坐。」",
              "四十岁上下、素色布衣、圆脸的茶馆老板娘快步迎来，先在围裙上擦了擦手，再欠身相迎；笑意只到唇角，眼神恭敬不敢直视。",
              "65mm 中近跟拍，焦点从她的手移到脸上，浅景深，背景茶馆门帘虚化。",
              "茶馆檐下暖光 3200K，门内透出的暖黄与门外斜阳形成冷暖交界。",
              "古装生活写实，衣料有洗旧痕迹；禁止夸张谄媚表情。"),
             ("P01-S03", 3, "中 ｜ 35mm", "固定 + 纵深构图（门框画中画）",
              "包房内他斜倚木椅，右手托白瓷杯，杯沿抵着薄唇；视线透过半开的雕花窗棂落在外面。",
              "无",
              "男子斜倚木椅，右手托白瓷茶杯，杯沿抵着薄唇；凤眸并未看茶，而是透过半开雕花窗棂落在外间，眼底无波。",
              "35mm 中景，固定机位，窗棂形成画中画，人物在右三分线。",
              "室内烛暖 3000K + 窗外斜阳入画，明暗交界压在他下颌线。",
              "清冷矜贵，参考《一代宗师》室内构图；禁止多余道具。"),
             ("P01-S04", 3, "近 ｜ 85mm", "浅景深 + 轻推（隔窗偷听）",
              "外间两村民凑近，一人压低声音；说到「差点没挺过来」时眉毛一挑，眼睛睁大。",
              "村民甲（气声、语速快、尾音下压）：「你们听说了吗，昨晚大祭司差点没挺过来！」",
              "外间木桌旁两名村民凑近说话，一人压着嗓子，眉毛一挑、眼睛睁大；另一人端着茶碗前倾身子。",
              "85mm 近景，浅景深，焦点在说话者的眉与嘴，隔着窗棂前景虚化。",
              "茶馆大堂暖光偏暗，桌面油光反一点高光。",
              "市井写实，表情要有「说闲话」的兴奋与小心；禁止整齐划一。"),
         ]),
    dict(no="P02", time="00:15–00:30", title="一句吩咐 · 夜叩窗",
         func="把「明天午时」这个约钉下来，并让陆驰露出唯一一次情绪（欣慰）。",
         emo="平淡 → 威压 → 笃定", hook="他邀她，是为了治病，还是为了她的来历？",
         sound="瓷杯落桌、古琴低音、夜风、三声叩窗", trans="包房 → 月上柳梢（延时）",
         shots=[
             ("P02-S01", 4, "特 ｜ 100mm", "微距 + 缓推（眼神静止）",
              "他轻抿一口，杯沿遮住半唇；喉结动了一下，睫毛垂落投下细影。",
              "无",
              "男子轻抿一口清茶，白瓷杯沿遮住半唇，喉结轻轻动了一下，睫毛垂落投下细影；指尖仍在杯壁上。",
              "100mm 微距，f/2.0，极缓推近，焦点在唇与杯沿。",
              "烛火 3000K 左下侧，杯壁反光柔和，面部一半沉暗。",
              "克制特写，参考《色戒》茶戏；禁止夸张吞咽。"),
             ("P02-S02", 4, "中近 ｜ 65mm", "缓推 + 视线引导（命令）",
              "他放下茶杯，视线始终没离开窗外；开口时下颌不动，只有唇在动。",
              "邶（音量低、语速慢、每句停半拍、尾音下沉）：「去，明日此时，邀林姑娘来此茶馆一叙。」",
              "男子放下茶杯，视线未离开窗外，缓缓开口；下颌几乎不动，只有唇形变化，右眼下泪痣在侧光里清楚。",
              "65mm 中近，缓推 10%，焦点在眼睛，浅景深。",
              "室内烛暖为底，窗外斜阳在他侧脸勾一道边光。",
              "上位者威压靠「不动」建立；禁止皱眉瞪眼。"),
             ("P02-S03", 4, "近 ｜ 85mm", "过肩正反打（垂首应答）",
              "陆驰立刻躬身，双臂贴紧，头低到颈线折起；应声短促。",
              "陆驰（沉声、短促、尾音收）：「是，主子。」",
              "玄色劲装护卫立刻躬身，双臂贴紧身侧，头低到颈线折成直角；应声短促有力，肩线绷紧。",
              "85mm 近景过肩，前景男子肩头虚化，焦点在护卫低垂的眉眼。",
              "逆光轮廓，面部只余一线暖光。",
              "武人礼数，参考黑泽明家臣戏；禁止多余动作。"),
             ("P02-S04", 3, "特 ｜ 100mm", "微距 + 焦点转移（情绪外泄）",
              "陆驰垂眸，眼底掠过一丝欣慰：眼角先松，唇角极轻上扬又压回；他的指尖在摩挲杯壁。",
              "无（画外音：主子终于愿意直面身上的顽疾）",
              "陆驰垂眸瞬间，眼角先松、唇角极轻上扬又立刻压回；画面右侧插入男子指尖摩挲杯壁的微距。",
              "100mm 微距双焦点，焦点从护卫眼角转到指尖的杯壁。",
              "侧光弱照，眼白一点高光。",
              "克制情绪，一眼即逝；禁止笑出来。"),
         ]),
    dict(no="P03", time="00:30–00:45", title="夜叩窗 · 转身即走",
         func="夜里传话，他连门都不进；她一句「我还没答应呢」立住性格。",
         emo="警惕 → 释然 → 无奈", hook="这个侍卫为什么连茶都不喝一口就走？",
         sound="夜风、瓦片轻响、三声叩窗、衣袂破风", trans="窗外黑影 → 屋内烛火（硬切）",
         shots=[
             ("P03-S01", 4, "大远 ｜ 35mm", "升格 60fps + 仰角（李安式屋顶轻功）",
              "月挂柳梢，一道玄色身影踏瓦而来，足尖点过屋脊，瓦片轻震，露珠弹起。",
              "无",
              "月挂柳梢，一道玄色身影踏瓦掠来，足尖点过屋脊；瓦片轻震，草叶上一颗露珠弹起，在月光里碎开。",
              "35mm 仰角，升格 60fps，镜头随身影横移，末尾落幅在院落。",
              "冷月 6500K，青黑基调，仅瓦面与露珠留高光。",
              "写实轻功，参考《卧虎藏龙》屋顶；禁止威亚感与慢动作滥用。"),
             ("P03-S02", 4, "中近 ｜ 65mm", "固定 + 焦点从窗到脸（开窗）",
              "她推窗，月光落在脸上；眼神先警惕（眉心微压、瞳孔收紧），看清来人后眉峰一松。",
              "林（轻声、带戒备、尾音上扬）：「夜已深，不知你来此处有何事？」",
              "她推开木窗，月光瞬间落在脸上；眼神先警惕——眉心微压、瞳孔收紧；认出是白日那位侍卫后，眉峰一松、唇微张。",
              "65mm 中近，固定，焦点从窗棂移到她的眼睛，浅景深。",
              "月光 6500K 侧逆，脸上一半冷光一半暖烛。",
              "写实少女，清透质感；禁止磨皮与过度美颜。"),
             ("P03-S03", 4, "中近 ｜ 50mm", "越肩 + 略俯（公事传话）",
              "他拱手躬身，视线落在地面；说话时音量压到最低，气息不散。",
              "陆驰（压低、公事、不带情绪、尾音收）：「叨扰林姑娘安眠了。我家主子邀姑娘明日午时，于集市茶馆一叙，有要事讨教。」",
              "玄色劲装护卫拱手躬身，视线落在地面三尺处；说话时音量压到极低，气息收而不散，肩背始终绷直。",
              "50mm 越肩，略俯角度，前景窗框虚化，焦点在他唇与手。",
              "月冷光为主，他半张脸在暗处。",
              "夜行传话的克制；禁止抬头直视。"),
             ("P03-S04", 3, "近 ｜ 85mm", "手持 + 甩向背影（转身即走）",
              "话音落，他不等回应便足尖一点腾空；她探身伸手，指尖抓空。",
              "林（急、尾音上扬、气声）：「等等！」",
              "他话音落下即转身，足尖一点腾空而起；她探身伸手，指尖抓空，衣袖被夜风带起。",
              "85mm 近景，手持轻晃，末尾快甩向屋顶背影。",
              "冷月光逆照，她脸上都是急出来的亮。",
              "动作干脆不拖泥带水；禁止回眸。"),
         ]),
    dict(no="P04", time="00:45–01:00", title="空夜喃喃 · 晨光急叩",
         func="她一句「我还没答应呢」定性格；天亮一场急叩把节奏从静推快。",
         emo="无奈 → 好奇 → 急", hook="大祭司到底怎么了？",
         sound="夜风、窗棂合上、晨鸟、急促拍门、少年变声的喊", trans="窗合 → 晨光（硬切）",
         shots=[
             ("P04-S01", 3, "大远 ｜ 35mm", "固定 · 留白（人走夜空）",
              "几个起落，身影消失在屋脊尽头；院落只剩月光与一只被惊起的夜鸟。",
              "无",
              "身影几个起落消失在屋脊尽头；院落只剩月光，一只夜鸟被惊起，掠过月面。",
              "35mm 大远景固定，画面上方留白给月与鸟。",
              "冷月青黑，屋檐下有一豆未熄的灯。",
              "空镜留白，参考侯孝贤；禁止配乐堆情绪。"),
             ("P04-S02", 4, "中近 ｜ 65mm", "固定 + 轻微后拉（喃喃）",
              "她扶着窗框，怔了半秒，然后无奈地笑了一下；说到「奇怪的人」时摇了摇头。",
              "林（先无奈后轻笑、尾音拖）：「我还没答应呢，当真是个奇怪的人。」",
              "她扶着窗框怔了半秒，随即无奈地笑了一下，说到「奇怪的人」时轻轻摇头，发梢扫过肩头。",
              "65mm 中近，固定，末帧轻微后拉 5%。",
              "月光柔化，面部补一点暖烛。",
              "写实少女神态；禁止夸张嘟嘴。"),
             ("P04-S03", 4, "近 ｜ 85mm", "轻推 + 焦点在眼睛（困惑）",
              "她歪头，眉心轻蹙，视线向右上方飘；嘴唇小小地抿了一下。",
              "无",
              "她歪了歪头，眉心轻蹙，视线向右上方飘，嘴唇小小地抿了一下；眼底是干净的困惑，没有戒备。",
              "85mm 近景，轻推，焦点在眼睛，浅景深。",
              "侧光，睫毛在颧骨投下细影。",
              "微表情特写；禁止眉毛上扬夸张。"),
             ("P04-S04", 4, "中近 ｜ 65mm", "手持 + 快速推门（急转）",
              "晨光打在门上，拍门声急促；她一把拉开门，头发还没梳齐。",
              "岜沅（画外，慌乱、破音、气促）：「林姑娘！林姑娘！」",
              "晨光斜打在木门上，门板被拍得震动；她一把拉开门，发髻未梳齐，几缕碎发贴在颊边，眼神瞬间清明。",
              "65mm 中近，手持轻晃，末帧切门板震动特写。",
              "晨光 4300K 逆照，门缝漏出一道金线。",
              "急转节奏，拍门要有实感；禁止慢门。"),
         ]),
    dict(no="P05", time="01:00–01:15", title="救人 · 施针逼毒",
         func="把她的专业立在动作上：看瞳、搭腕、施针、推掌，四步不含糊。",
         emo="急 → 稳 → 惊", hook="一口黑血喷出来，她为什么一点都不慌？",
         sound="急促脚步、银针出鞘、衣料摩擦、闷哼", trans="巷道 → 主屋（跟拍不切）",
         shots=[
             ("P05-S01", 4, "中全 ｜ 35mm", "跟拍 + 手持（跑动）",
              "两人穿巷疾行，少年几次回头看她；她提着裙摆，呼吸稳，眼神已经落到远处屋门。",
              "林（语速快、音量稳、尾音收）：「别慌，我随你去。」",
              "两人穿过晨雾巷道，少年几次回头看她，脚步虚浮；她提着裙摆疾行，呼吸节奏稳，视线已越过他落在远处屋门上。",
              "35mm 中全，跟拍手持，人物在画面中轴偏左。",
              "晨光侧逆，雾气在两人之间流动。",
              "写实赶路，脚步要有重量；禁止轻飘。"),
             ("P05-S02", 5, "近 ｜ 85mm", "微距 + 焦点在手指（诊脉）",
              "她俯身看瞳，眉心绷紧；两指搭腕，指腹缓缓加压，眼睛半阖听脉。",
              "林（低声、自言自语、气息匀）：「……脉在指下跳得很乱。」",
              "她俯身查看老者眼瞳，眉心绷紧；两指搭上腕脉，指腹缓缓加压，眼睛半阖，呼吸放慢；指尖随脉象轻移。",
              "85mm 近景，焦点在指腹与脉门，浅景深，镜头极轻推。",
              "屋内暗，窗边一束晨光切在手腕上。",
              "医学写实，手指用力要有细节；禁止多余手指。"),
             ("P05-S03", 4, "特 ｜ 100mm", "微距 + 俯拍（施针）",
              "银针出鞘，她手腕一沉，针尖入穴；针尾轻颤，她指尖稳得没有一丝抖。",
              "无（银针轻响）",
              "银针从布卷中抽出，她手腕一沉，针尖精准入穴；针尾轻颤，她指尖稳得没有一丝抖动，另三指扶着皮肤。",
              "100mm 微距俯拍，焦点在针尖与皮肤接触点，浅景深。",
              "窗光 4300K 冷白，针身一点金属高光。",
              "专业感来自稳，不来自快；禁止血与伤口特写。"),
             ("P05-S04", 2, "近 ｜ 65mm", "手持 + 快速下压（推掌）",
              "双掌贴合，掌根一上一下旋揉两次，稳稳按在两侧肩胛，轻轻一送。",
              "无",
              "她双掌贴合，掌根一上一下旋揉两次，稳稳按在老者两侧肩胛上，掌心透出一层极淡的气感，轻轻一送。",
              "65mm 近景手持，随掌根下压轻微下沉。",
              "背光，掌心边缘有一圈极淡光晕，克制。",
              "内力表现要「看得见但不炫」；禁止特效爆炸。"),
         ]),
    dict(no="P06", time="01:15–01:30", title="黑血 · 少年慌",
         func="她一句话稳住全屋；少年的结巴把惊惶演满。",
         emo="惊 → 慌 → 稳", hook="这一口黑血，是好转还是恶化？",
         sound="呕血声、咳嗽、少年变声、众人抽气", trans="主屋 → 院门（顺接）",
         shots=[
             ("P06-S01", 4, "特 ｜ 100mm", "固定 + 轻微手持（喷血）",
              "老者身子前倾，一口暗色血喷出，随即几声虚弱咳嗽，软倒下去。",
              "无（“噗——咳咳咳”）",
              "老者身子猛地前倾，一口暗色血喷出，落在布巾上；随即几声虚弱咳嗽，身子一软倒回榻上。",
              "100mm 特写，固定，轻微手持，焦点在嘴角与布巾。",
              "窗光偏冷，暗部保留层次。",
              "医学写实，避免血腥特写；如需过审可改为「一口浊气 + 黑帕」暗示。"),
             ("P06-S02", 4, "中近 ｜ 65mm", "手持 + 急推（少年慌）",
              "少年脸色瞬间煞白，往前冲半步又不敢碰；张嘴两次才发出声音，喉结上下滚。",
              "岜沅（颤抖、结巴、音量拔高）：「这、这！林姑娘，我阿父他、他怎么样了？」",
              "少年脸色瞬间煞白，往前冲半步又不敢碰；张嘴两次才发出声音，喉结上下滚动，手指在空中蜷了一下。",
              "65mm 中近，手持急推，焦点在少年的眼与嘴。",
              "逆光，脸上一半在阴影里。",
              "少年慌乱要真，参考《一一》少年戏；禁止夸张嚎哭。"),
             ("P06-S03", 4, "近 ｜ 85mm", "固定 + 焦点扫过众人（安抚）",
              "她收回手掌，目光缓缓扫过一屋子人，最后落在少年脸上；语速慢下来。",
              "林（沉稳、逐字、尾音落）：「无妨，如今已无大碍。今早吐出的黑血，正是淤毒排出的好转之兆。」",
              "她从容收回手掌，目光缓缓扫过屋内众人，最后落在少年慌乱的脸上；语速明显放慢，每个字都落稳。",
              "85mm 近景固定，焦点随视线从左到右扫一次，末尾落少年。",
              "窗光在她脸上形成柔和过渡，眼神高光一处。",
              "专业者的镇定；禁止微笑与夸张手势。"),
             ("P06-S04", 3, "中全 ｜ 35mm", "固定 + 缓缓后拉（气氛落地）",
              "满屋人齐齐长舒一口气，有人扶墙、有人抹脸；少年腿一软坐下。",
              "无",
              "满屋人齐齐长舒一口气：有人扶着墙，有人抹了把脸，少年腿一软坐回凳上，肩膀塌下去。",
              "35mm 中全固定，末帧缓缓后拉，交代屋内人数与位置。",
              "屋内光线回升半档，暗示气氛松动。",
              "群像反应要有先后层次；禁止整齐同步。"),
         ]),
    dict(no="P07", time="01:30–01:45", title="认姐 · 狂奔",
         func="甜点在「筱筱姐」；抬头见日头，节奏立刻抽紧。",
         emo="感激 → 亲昵 → 惊觉", hook="糟了——她忘了午时之约！",
         sound="少年笑声、揉头、蝉鸣、急促脚步、心跳", trans="院门 → 集市（快切）",
         shots=[
             ("P07-S01", 4, "中全 ｜ 35mm", "固定 + 略仰（躬身谢恩）",
              "少年在院门口深深躬身，头几乎碰到膝盖；起身时眼睛发亮。",
              "岜沅（激动、音量高、尾音上扬）：「真是太感谢你了林姑娘！你就是我们家的救命大恩人！」",
              "少年在院门口深深躬身，头几乎碰到膝盖；起身时眼睛发亮，双手还攥着衣角。",
              "35mm 中全固定，略仰角度，背景是院门与晨光。",
              "晨光 4300K 侧逆，少年轮廓一圈亮边。",
              "少年真挚，参考《海街日记》少年戏；禁止浮夸行礼。"),
             ("P07-S02", 4, "中近 ｜ 65mm", "固定 + 浅景深（柔和叮嘱）",
              "她眉眼柔和，唇角有浅浅笑意；说话时语气放软，尾音轻。",
              "林（轻柔、语速慢、尾音软）：「无需多礼。今后好好照料你阿父，遇到不懂的养护事宜，随时可来找我。」",
              "她眉眼柔和，唇角噙着浅浅笑意，说话时语气放软；抬手把一缕碎发别到耳后，动作很轻。",
              "65mm 中近，固定，浅景深，背景院墙虚化。",
              "顺光柔和，面部干净，眼中有柔和反光。",
              "温柔不腻；禁止过度甜笑。"),
             ("P07-S03", 4, "中近 ｜ 50mm", "越肩 + 轻推（认姐）",
              "少年一拍胸脯，下巴一抬，笑得露出一点牙；她被这模样逗得眼睛弯起来。",
              "岜沅（爽利、明亮、快）：「好！林姑娘，往后你就是我的筱筱姐，你直接唤我岜沅就好！」",
              "少年用力一拍胸脯，下巴一抬，笑得露出一点牙；她被这模样逗得眼睛弯成月牙，肩膀轻轻一松。",
              "50mm 中近越肩，轻推，焦点在少年胸脯到她的眼睛之间转移一次。",
              "晨光大亮，暖调上扬。",
              "少年感要足，参考《小森林》少年；禁止油腻。"),
             ("P07-S04", 3, "中 ｜ 35mm", "固定 + 急摇向上（惊觉）",
              "她抬手揉他头顶，动作慢；忽然抬眼看向日头，脸色一变。",
              "林（软、慢）：「好，小岜沅，好好照顾你阿父。」（顿）「糟了！」",
              "她抬手轻轻揉了揉少年的头顶，动作很慢；说到一半忽然抬眼望向天际，日头高悬，脸色一变，手停在半空。",
              "35mm 中景固定，末帧急摇向上带出正午烈日。",
              "正午 5600K 硬光，地面反光强，影子很短。",
              "反差节奏：软 → 急，靠摇镜完成；禁止切黑。"),
         ]),
    dict(no="P08", time="01:45–02:00", title="茶馆问答 · 钩子",
         func="两人第一次真正对话：他审问，她反问；「我信」两个字把关系翻面。",
         emo="审 → 慌 → 坦 → 惊", hook="她说「我是山里凭空生的」——他为什么当场就信？",
         sound="茶壶注水、瓷杯轻碰、呛水声、木门合上、低频闷响", trans="木门合上 → 硬切黑",
         shots=[
             ("P08-S01", 4, "中 ｜ 35mm", "固定 + 门框构图（闯入）",
              "她冲进茶馆，双手扶腰，胸口起伏；说话断成两截，脸颊跑得泛红。",
              "林（气喘、断句、歉疚）：「实在抱歉……被急事耽搁了片刻，劳你等候。」",
              "她冲进茶馆大堂，双手扶腰，胸口明显起伏；发丝凌乱，脸颊跑得泛红；说话断成两截，气还没匀。",
              "35mm 中景，固定，门框构成画中画，背景茶客虚化。",
              "茶馆内暖光 3200K，门口斜射一道强光。",
              "真实喘息，参考《饮食男女》奔跑后；禁止立刻恢复端庄。"),
             ("P08-S02", 4, "中全 ｜ 35mm", "缓推 + 俯仰（斟茶 · 单字）",
              "他垂眼斟茶，水流稳；抬眸锁定她，只说一个字。她大大方方落座，直视回去。",
              "邶（单字、低沉、不带温度）：「坐。」",
              "男子垂眼斟茶，水流稳而不断；放下茶壶后抬眸锁定她，睫毛抬起的瞬间眼神压过来；她毫不避让地坐下，直视回去。",
              "35mm 中全，缓推，略俯拍他、略仰拍她，末帧双人同框。",
              "窗棂光切在两人之间，桌上茶汤映一层暖光。",
              "「坐」字比任何台词都重；禁止他先笑。"),
             ("P08-S03", 5, "近 ｜ 85mm", "正反打 + 急推特写（审问 · 呛水）",
              "他忽然开口，语速快、气压低；她一口茶呛住，猛抬头，眼睛睁大。",
              "邶（语速快、气压低、尾音下压）：「说吧，你并非本部落之人，究竟是什么来历？」／林（呛、气声、错愕）：「咳咳！你说什么？」",
              "男子忽然开口，语速加快、气压骤低，下颌线绷起；她一口茶呛在喉间，猛地抬头，眼睛睁大、鼻尖发红、手忙捂住嘴。",
              "85mm 近景正反打，末尾急推到她呛住的脸上，抓拍水珠与睫毛。",
              "侧光，他半张脸在暗处；她脸上有呛出来的水光。",
              "审问的「快」与呛水的「乱」要对切；禁止他无表情。"),
             ("P08-S04", 3, "特 ｜ 100mm", "极缓推 + 焦点锁定（放缓追问）",
              "他停顿一拍，收敛气场；再开口时音量降下来，眼神却没松。她攥紧杯壁，指节泛粉。",
              "邶（音量降、语速慢、尾音轻）：「你是什么人？或者说，你从何处而来？」",
              "男子停顿一拍，收敛迫人气场，再开口时音量明显降低、语速放慢；眼神依旧牢牢锁在她脸上。她指尖攥住白瓷杯壁，指节泛出浅浅粉色。",
              "100mm 微距，极缓推，焦点在他眼睛，末帧焦点转到他唇形。",
              "烛光 + 窗光双光源，明暗交界正落在两人之间。",
              "钩子镜头：问题落地即切黑；禁止给答案。"),
         ]),
]


# ============================ 文档 ============================
doc = Document()
sec = doc.sections[0]
sec.orientation = WD_ORIENT.LANDSCAPE
sec.page_width = Mm(297)
sec.page_height = Mm(210)
sec.left_margin = Mm(15)
sec.right_margin = Mm(15)
sec.top_margin = Mm(12)
sec.bottom_margin = Mm(11)
sec.header_distance = Mm(7)
sec.footer_distance = Mm(7)
style_font(doc.styles["Normal"], "宋体", 10.5)

hp = sec.header.paragraphs[0]
hp.alignment = 2
set_run(hp.add_run("《第二章 · 初入部落》EP02 ｜ 120 秒 ｜ 8 单元 / 32 镜 ｜ 重点：语气与说话微表情"),
        name="微软雅黑", size=8, color="8C9AA3")

fp = sec.footer.paragraphs[0]
fp.alignment = 1
set_run(fp.add_run("— "), name="微软雅黑", size=8, color="8C9AA3")
_r = fp.add_run()
_f1 = OxmlElement("w:fldChar"); _f1.set(qn("w:fldCharType"), "begin")
_it = OxmlElement("w:instrText"); _it.set(qn("xml:space"), "preserve"); _it.text = "PAGE"
_f2 = OxmlElement("w:fldChar"); _f2.set(qn("w:fldCharType"), "end")
for _e in (_f1, _it, _f2):
    _r._r.append(_e)
set_run(_r, name="微软雅黑", size=8, color="8C9AA3")
set_run(fp.add_run(" —"), name="微软雅黑", size=8, color="8C9AA3")

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(22)
set_run(p.add_run("SHORT DRAMA · EP02 · 120s"), name="微软雅黑", size=8.5, bold=True, color=C_GOLD)
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(6)
set_run(p.add_run("第二章 · 初入部落"), name="微软雅黑", size=32, bold=True, color=C_INK)
p = doc.add_paragraph()
set_run(p.add_run("EP02 · 120 秒 ｜ 8 个 15 秒单元 ｜ 32 镜 ｜ 运镜克制、动作恰到好处、语气与微表情为骨"),
        name="微软雅黑", size=12, color="4A6675")
p = doc.add_paragraph()
set_run(p.add_run("续接第一章 V5 结尾（月圆夜 · 今夜十五）：次日午后开场，走「邀约 → 夜访 → 救治 → 茶馆问答 → 密信」主线。"),
        name="宋体", size=10.5, color="5C7A88")

rows = [
    ("版本", "EP02 V1（120 秒；沿用 EP01 V5 的单元结构与提示词规范）"),
    ("时长", "120 秒（2′00″）＝ 8 × 15 秒；32 镜，平均 3.75s/镜"),
    ("画幅", "9:16 竖屏 1080×1920；安全区上 1/6、下 1/5"),
    ("主线", "茶馆闲话 → 一句吩咐 → 夜叩窗 → 晨急叩 → 施针逼毒 → 认姐 → 茶馆审问 → 「我信」钩子"),
    ("本集重点", "说话语气（音量/语速/停顿/气声/尾音）与说话微表情（眉/眼/唇/下颌/喉结/视线落点）逐镜标注"),
    ("提示词", "四行；主体动作行内含微表情与语气；镜头行只写一个运动；光色写色温+方位+影调"),
    ("人物锚", "邶峙珩 21 岁 220cm 右眼下泪痣；林洛筱 22 岁 176cm 左耳银铃；陆驰 24 岁 219cm；新增青云嫂子、巫罗·岜沅（少年）"),
]
t = doc.add_table(rows=0, cols=2)
table_borders(t)
cell_margins(t, 46, 46, 86, 86)
for k, v in rows:
    row = t.add_row()
    no_split(row)
    cell_text(row.cells[0], k, name="微软雅黑", size=8.5, bold=True, color=C_INK, align=1)
    cell_text(row.cells[1], v, size=9, color="2B3B44")
    shade(row.cells[0], C_SOFT)
fixed(t, [1.45, 8.95])

# ===== 语气与微表情对照表 =====
heading(doc, "一、本集核心：语气与说话微表情对照", 1, new_page=True)
para(doc, "短剧的对话戏，信息在台词里，**情绪在语气和脸上**。下表把每句关键台词拆成「语气参数」与「微表情落点」，逐镜执行即可，不要靠演员自由发挥。", size=9, after=6)

tone = [
    ("谁 · 台词", "语气参数（音量 / 语速 / 停顿 / 气声 / 尾音）", "说话微表情（眉 / 眼 / 唇 / 下颌 / 喉结 / 视线）"),
    ("青云嫂子：「少主，您往里坐。」", "音量低、语速中、句尾收；气声多、不拖", "眉平、眼不敢直视、笑意只到唇角；欠身时视线落地面"),
    ("村民甲：「昨晚大祭司差点没挺过来！」", "气声、语速快、尾音下压；东张西望", "眉挑、眼睁大、嘴角带一点兴奋；说话时凑近"),
    ("邶：「去，明日此时，邀林姑娘来此茶馆一叙。」", "音量低、语速慢、每句停半拍、尾音下沉；不加重音", "眉不动、眼不眨、下颌几乎不动、只有唇动；视线不离开窗外"),
    ("陆驰：「是，主子。」", "沉声、短促、尾音断；不加语气词", "头低到颈线折起、肩线绷紧、视线落地面"),
    ("林：「夜已深，不知你来此处有何事？」", "音量中、语速中、尾音上扬；带一点戒备", "眉心微压、瞳孔收紧；认出后眉峰一松、唇微张"),
    ("陆驰：「我家主子邀姑娘明日午时……一叙。」", "压低到最低、字数清楚、尾音收；不寒暄", "视线落地面三尺、不抬头、拱手时手背绷直"),
    ("林：「等等！」", "拔高、气声、尾音上扬且短", "探身、伸手、指尖抓空；眼急"),
    ("林：「我还没答应呢，当真是个奇怪的人。」", "先无奈后轻笑、尾音拖长；自言自语", "摇头、无奈笑、发梢扫肩；眼尾微垂"),
    ("岜沅：「林姑娘！林姑娘！」", "破音、气促、音量高；喊到第二声哑", "拍门时手掌发红、脸白、眉头拧死"),
    ("林：「别慌，我随你去。」", "语速快、音量稳、尾音收；压住对方", "神色一敛、眼瞬间清明、下颌微收"),
    ("岜沅：「我阿父他、他怎么样了？」", "颤抖、结巴、音量拔高；换气两次", "脸煞白、喉结上下滚、手指在空中蜷"),
    ("林：「无妨，如今已无大碍。」", "沉稳、逐字、尾音落；语速放慢一半", "收回手掌、目光扫屋、眼神高光一处；不笑"),
    ("岜沅：「你就是我们家的救命大恩人！」", "激动、音量高、尾音上扬", "深躬、起身时眼发亮、双手攥衣角"),
    ("林：「好，小岜沅，好好照顾你阿父。」", "软、慢、尾音轻；带一点笑", "揉头顶时手腕放松；眼弯成月牙"),
    ("林：「实在抱歉……被急事耽搁了片刻。」", "气喘、断句、歉疚；尾音虚", "扶腰、胸口起伏、脸颊泛红、鼻尖一点汗"),
    ("邶：「坐。」", "单字、低沉、不带温度；只一个字，不加助词", "抬眸、睫毛抬起的速度比平时快半拍；唇只动一次"),
    ("邶：「你究竟是什么来历？」", "语速快、气压低、尾音下压；字与字之间不留缝", "下颌线绷起、眉略压、眼神锁死"),
    ("林：「咳咳！你说什么？」", "呛声、气声、错愕", "猛抬头、眼睁大、鼻尖红、手捂嘴；水光在睫毛"),
    ("邶：「你是什么人？你从何处而来？」", "音量降、语速慢、尾音轻；像在问自己也像在问她", "眼不松、唇线平；说「何处」时视线下移半寸"),
]
t = doc.add_table(rows=0, cols=3)
table_borders(t, sz=3)
cell_margins(t, 30, 30, 56, 56)
for i, r in enumerate(tone):
    row = t.add_row()
    no_split(row)
    if i == 0:
        for j, v in enumerate(r):
            cell_text(row.cells[j], v, name="微软雅黑", size=7.8, bold=True, color=C_WHITE, align=1)
            shade(row.cells[j], C_INK)
        repeat_hdr(row)
        continue
    cell_text(row.cells[0], r[0], size=7.2, color="22323A")
    cell_text(row.cells[1], r[1], size=7.2, color="4A6675")
    cell_text(row.cells[2], r[2], size=7.2, color="6B3A2E")
fixed(t, [3.5, 3.4, 3.5])

# ===== 逐单元 =====
heading(doc, "二、逐单元分镜与提示词（8 单元 / 32 镜）", 1, new_page=True)
para(doc, "每单元一页：单元卡 → 分镜表（台词栏含语气标注）→ 4 条提示词（主体动作行内含微表情与语气）。", size=8.8, after=4)

for idx, u in enumerate(UNITS):
    heading(doc, f"{u['no']} ｜ {u['time']} ｜ {u['title']}", 2, new_page=(idx > 0))

    def meta_line(items):
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = 1.18
        pf.left_indent = Pt(8)
        pf.keep_with_next = True
        para_shade(p, C_SOFT)
        para_bar(p)
        first = True
        for label, val in items:
            if not first:
                set_run(p.add_run("　｜　"), name="微软雅黑", size=7.0, color="9AA9B1")
            set_run(p.add_run(f"{label}："), name="微软雅黑", size=7.0, bold=True, color="2F5A6B")
            set_run(p.add_run(val), name="宋体", size=7.1, color="22323A")
            first = False

    meta_line([("功能", u["func"])])
    meta_line([("情绪", u["emo"]), ("钩子", u["hook"])])
    meta_line([("声音", u["sound"]), ("转场", u["trans"])])

    st = doc.add_table(rows=1, cols=6)
    table_borders(st, sz=3)
    cell_margins(st, 20, 20, 48, 48)
    for j, txt in enumerate(("镜号", "时长", "景别 · 焦段", "运镜", "画面 · 微表情", "台词 · 语气")):
        cell_text(st.rows[0].cells[j], txt, name="微软雅黑", size=7.0, bold=True, color=C_WHITE, align=1)
        shade(st.rows[0].cells[j], C_INK)
    repeat_hdr(st.rows[0])
    no_split(st.rows[0])
    for s in u["shots"]:
        row = st.add_row()
        no_split(row)
        cell_text(row.cells[0], s[0].split("-")[1], name="微软雅黑", size=7.2, bold=True, color=C_INK, align=1)
        cell_text(row.cells[1], f"{s[1]}s", name="微软雅黑", size=7.2, color="2B3B44", align=1)
        cell_text(row.cells[2], s[2], size=6.7, color="2B3B44")
        cell_text(row.cells[3], s[3], size=6.7, color="3A4A52")
        cell_text(row.cells[4], s[4], size=6.8, color="1F2E36")
        cell_text(row.cells[5], s[5], size=6.7, color="6B3A2E")
    fixed(st, [0.6, 0.4, 1.0, 1.45, 4.55, 2.45])

    ph = doc.add_paragraph()
    ph.paragraph_format.space_before = Pt(2)
    ph.paragraph_format.space_after = Pt(0)
    ph.paragraph_format.line_spacing = 1.0
    ph.paragraph_format.keep_with_next = True
    set_run(ph.add_run(f"▍{u['no']} 提示词（{len(u['shots'])} 条）"), name="微软雅黑", size=7.2, bold=True, color=C_GOLD)
    for s in u["shots"]:
        prompt_block(doc, s[0], s[1], s[2], s[6], s[7], s[8], s[9])

# ===== 时长核对 =====
heading(doc, "三、时长核对与节奏", 1, new_page=True)
total = sum(s[1] for u in UNITS for s in u["shots"])
rows2 = [("单元", "时间码", "时长", "累计", "本单元钩子", "语速/情绪曲线")]
acc = 0
pacing = {
    "P01": "慢 · 松 → 留耳",
    "P02": "最慢 · 只说不动",
    "P03": "静 → 快（转身即走）",
    "P04": "慢 → 急（拍门）",
    "P05": "快 → 稳（专业）",
    "P06": "惊 → 稳",
    "P07": "软 → 急（抬头看日）",
    "P08": "审 → 慌 → 停（钩子）",
}
for u in UNITS:
    d = sum(s[1] for s in u["shots"])
    acc += d
    rows2.append((u["no"], u["time"], f"{d}s", f"{acc}s", u["hook"], pacing[u["no"]]))

t = doc.add_table(rows=0, cols=6)
table_borders(t, sz=3)
cell_margins(t, 30, 30, 56, 56)
for i, r in enumerate(rows2):
    row = t.add_row()
    no_split(row)
    if i == 0:
        for j, v in enumerate(r):
            cell_text(row.cells[j], v, name="微软雅黑", size=7.6, bold=True, color=C_WHITE, align=1)
            shade(row.cells[j], C_INK)
        repeat_hdr(row)
        continue
    cell_text(row.cells[0], r[0], name="微软雅黑", size=7.4, bold=True, color=C_INK, align=1)
    cell_text(row.cells[1], r[1], name="微软雅黑", size=7.4, color="2B3B44", align=1)
    cell_text(row.cells[2], r[2], size=7.4, color="2B3B44", align=1)
    cell_text(row.cells[3], r[3], size=7.4, color="2B3B44", align=1)
    cell_text(row.cells[4], r[4], size=7.4, color="6B3A2E")
    cell_text(row.cells[5], r[5], size=7.4, color="4A6675", align=1)
fixed(t, [0.65, 1.1, 0.65, 0.75, 4.9, 2.35])

para(doc, f"合计 {total} 秒 = 2′00″，8 单元 × 15 秒；32 镜，平均 {total/32:.2f} 秒/镜。", size=9, bold=True, color=C_INK, before=6, after=6)

# ===== 导演备注 =====
heading(doc, "四、导演备注：对话戏的三条纪律", 2)
for s in [
    "**台词不喊**：本集没有一个需要「用力」的台词。男主的威压靠「不动」——下颌不动、只动唇；女主的急靠「气息」——断句、换气、尾音虚。",
    "**一句台词只给一个表情**：眉毛、眼睛、嘴角、下颌，四者一次只动一处；动两处就是夸张，AI 生成也会崩。",
    "**说话前那半拍最值钱**：所有关键台词前加 0.3–0.5 秒停顿，镜头不动、只给眼神；这半拍就是「话里有话」。",
    "**避坑**：少年喊门、喷血、被呛水三处最容易演过头，要求「一次只做一件事」；AI 生成时把「夸张表情」写进负面词。",
]:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.left_indent = Pt(12)
    for _i, _seg in enumerate(("· " + s).split("**")):
        if _seg:
            set_run(p.add_run(_seg), size=8.8, bold=(_i % 2 == 1), color="2B3B44")

heading(doc, "五、与第一章的衔接", 2)
para(doc, "第一章结尾是月圆夜、毒纹发作、「今夜十五」。本集开场用三个信息接住：① 光从冷月变斜阳（次日午后）；② 村民一句「昨晚大祭司差点没挺过来」回扣第一章的「救人」；"
          "③ 他主动邀约，暗示他开始为「十五」这天的病做打算。第三条是本集真正的主线，也是第三章的入口。", size=8.8, color="2B3B44")

doc.core_properties.title = "第二章 初入部落 · EP02 分镜脚本 V1（120 秒）"
doc.core_properties.author = "编剧组"
doc.save(OUT)
print("saved:", OUT)
print("units:", len(UNITS), "shots:", sum(len(u["shots"]) for u in UNITS), "total_sec:", total)
