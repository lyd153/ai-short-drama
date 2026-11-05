# -*- coding: utf-8 -*-
"""生成 EP01 120 秒精炼版分镜脚本（8 单元 / 32 镜）。"""
from docx import Document
from docx.shared import Pt, Mm, Inches, RGBColor
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = r"E:\日常对话\ai-short-drama\docs\EP01-初遇\EP01_分镜脚本_V4_120秒版.docx"
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


def cell_margins(t, top=36, bottom=36, left=60, right=60):
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


def para(doc, text="", size=9.5, bold=False, color="2B3B44", before=0, after=4, align=3, name="宋体"):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = 1.32
    if text:
        parts = text.split("**")
        for i, seg in enumerate(parts):
            if not seg:
                continue
            set_run(p.add_run(seg), name=name, size=size,
                    bold=(bold or i % 2 == 1), color=color)
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
# 字段：镜号, 时长, 景别·焦段, 运镜, 画面动作, 台词, 主体动作, 镜头, 光色, 风格

UNITS = [
    dict(no="N01", time="00:00–00:15", title="寒月 · 毒兆",
         func="冷开场立三件事：月圆＝旧疾日、男主有病、兔子不是宠物。",
         emo="静 → 冷 → 不安", hook="他腕上那道旧纹，为什么今夜发亮？",
         sound="虫鸣、瓦片轻响、低音提琴长音", trans="烛火将熄 → 硬切黑",
         shots=[
             ("N01-S01", 5, "大远 ｜ 14mm", "航拍缓降 + 轻横移（库布里克式纵深）",
              "月压山脊，薄云掠月；一道黑影自屋脊掠过。", "无",
              "夜色如墨，将满圆月低悬墨色山脊，薄云掠过月面；山谷零星灯火，屋脊上一道黑影掠过。",
              "无人机航拍，俯角 35°，14mm，缓慢下降并轻微横移，底部留 1/5 黑场。",
              "冷月光 6500K，山雾散射，青黑基调，仅月轮与屋瓦保留高光，欠曝一档。",
              "写实古装，胶片颗粒，参考《影武者》月夜层次；无星空、无现代建筑。"),
             ("N01-S02", 4, "特 ｜ 100mm macro", "微距缓推 + 焦点锁腕脉",
              "腕上淡青灰旧纹轻轻一颤，颜色略深又淡下；指节收紧再松开。", "无",
              "极度特写：苍白手腕上一道淡青灰旧年纹路，颜色略深又淡下去；指节先收紧、再缓缓松开。",
              "100mm 微距，f/2.0，焦点锁纹路，前臂虚化，极缓推近。",
              "单一烛火 3200K 左下侧，纹路只在暗部显色；不发光、不做特效。",
              "写实皮肤质感；不出现伤口、血迹、血管特写。"),
             ("N01-S03", 3, "近 ｜ 85mm", "固定 + 呼吸感手持（伦勃朗光）",
              "他斜倚榻上，半张脸沉在烛影，睫毛抬起时眼底掠冷光。", "无",
              "21 岁古装男子斜倚榻上，墨蓝暗纹锦袍，高束墨发，肤色偏白，右眼下泪痣；半脸沉暗，睫毛抬起，眼底冷光。",
              "85mm，f/1.8，固定微呼吸，人物占屏 2/3，眼睛压上三分线。",
              "伦勃朗光：烛火单光源，鼻梁一处高光，眼窝与右颊沉暗，暖橘冷青对撞。",
              "矜贵慵懒，参考《一代宗师》浅景深；禁止现代发型。"),
             ("N01-S04", 3, "特 ｜ 90mm macro", "侧拍微距 + 升格 60fps",
              "灵兔伏在膝上，左耳缺一小口，鼻尖逸出一缕极淡青光。", "无",
              "雄性灵兔伏在膝上：咖白短毛，左耳缺一小口，颈间旧铜铃；闭眼吐纳，鼻尖一缕极淡青色光晕。",
              "90mm 微距侧拍，f/2.2，升格 60fps，焦点在鼻尖与铜铃。",
              "烛火暖底 + 极淡青光（仅鼻尖一圈），细小微粒，铜铃一点高光。",
              "写实动物，毛发根根分明；禁止拟人化、穿衣服。"),
         ]),
    dict(no="N02", time="00:15–00:30", title="密报 · 试探",
         func="一次汇报立住权力关系：他听着闲话，每一句都在筛人。",
         emo="慵懒 → 威压 → 收网", hook="「盯着她」——她凭什么值得盯？",
         sound="衣袂、膝盖触地闷响、古琴泛音", trans="门帘落下 → 烛火暗",
         shots=[
             ("N02-S01", 4, "中全 ｜ 35mm", "低角度仰拍 + 广角微畸变（黑泽明式）",
              "护卫落地无声，单膝跪地抱拳贴额，背脊绷直，冷汗滑到下颌。", "无（仅呼吸）",
              "24 岁玄色劲装护卫单膝跪地，双手抱拳贴额，背脊绷直，眉骨一道旧疤，额角渗汗滑至下颌。",
              "35mm 低角度仰拍，广角轻畸变强化压迫，人物置右三分线，左侧留纵深。",
              "逆光剪影 + 侧后轮廓光，面部只余一线暖光。",
              "古装写实，硬朗武人质感，衣料磨旧。"),
             ("N02-S02", 3, "大特 ｜ 100mm", "极缓推至瞳孔（3 秒走屏宽 12%）",
              "眼睛占满竖屏，瞳孔映一豆烛火，眼尾极轻上挑。", "「哦，外族女子。」（低、慢、尾音下沉）",
              "男子眼睛大特写占满画面，瞳孔映一豆烛火，眼尾极轻上挑，下眼睑松弛。",
              "100mm 微距，f/1.8，极缓推近至瞳孔占屏 1/4 停住。",
              "单点烛光 3200K，眼白不过曝，睫毛投影落颧骨，右侧压暗。",
              "极简面部特写，参考《荒野猎人》眼部打光；禁止磨皮。"),
             ("N02-S03", 5, "中近 ｜ 50mm", "越肩正反打 + 末尾缓拉",
              "他手指插进兔毛摩挲，视线扫向护卫；护卫肩线在抖。", "邶：「陆驰，这外族女子跟我有什么关系？」（顿）「要说就说点有用的——否则回去种地去。」",
              "男子一手搭膝、一手插进兔毛摩挲，视线扫向跪地护卫；护卫头更低，肩膀微抖。",
              "50mm 越肩，前景肩头虚化，焦点在护卫；最后一句说完缓拉 20cm。",
              "烛火暖 + 窗缝冷月漏光，明暗交界压在护卫背上。",
              "权力构图，参考黑泽明室内戏；禁止对称平拍。"),
             ("N02-S04", 3, "全 ｜ 24mm", "固定大景深（门框画中画）",
              "护卫躬身退到门边；穿堂风把烛火压得几乎伏倒。", "陆驰：「是！」（沉声）",
              "护卫躬身退到门边，门帘被风掀起又落下；穿堂风把烛火压得几乎伏倒，屋内骤然变暗。",
              "24mm 全景固定，大景深，门框形成画中画，人物偏右。",
              "烛光呼吸式明暗（亮—暗—微亮），风带动火苗。",
              "古装室内写实，木构、苇帘、青砖地。"),
         ]),
    dict(no="N03", time="00:30–00:45", title="月下 · 开市（合并原月下+破晓+集市）",
         func="倒计时钉进观众脑子；用空碗说完「他熬了一夜」，叠化到天亮开市。",
         emo="沉静 → 隐忍 → 破晓", hook="「明日就是十五了」——这十五天，他一直是这么熬的？",
         sound="风声、窗棂轻响、铜铃、空碗摩擦声、竹笛起", trans="窗月 → 叠化晨雾（全片唯一叠化）",
         shots=[
             ("N03-S01", 4, "中 ｜ 40mm", "侧后 3/4 缓推（侯孝贤式留白）",
              "他推开半扇窗，夜风涌入，烛火一晃；窗外圆月将满。", "无",
              "男子立在花窗前，抬手推开半扇窗，夜风涌入，烛火轻晃；窗外将满圆月，云影缓移。",
              "40mm 侧后 3/4 机位，缓推，从肩头推向窗外的月。",
              "月光 5600K 冷蓝铺满窗棂，室内只剩一线烛暖勾边，人物半剪影。",
              "极简东方构图，窗棂作画框，参考《卧虎藏龙》夜窗。"),
             ("N03-S02", 4, "特 ｜ 100mm", "微距 + 焦点从碗底移到手指",
              "矮几上一只空碗，碗底冷透的汤；他的手把碗推开，烛光在碗沿一晃。", "无",
              "矮几上一只素色空碗，碗底剩着冷透的汤；男性的手从画右伸入，把碗轻轻推开，烛光在碗沿一晃。",
              "100mm 微距，浅景深，焦点从碗底移到手指，镜头轻微下压。",
              "单点烛光 3200K 左上侧，碗沿高光柔和，背景沉暗。",
              "写实静物，陶土质感真实，参考《饮食男女》器物镜头。"),
             ("N03-S03", 4, "中近 ｜ 65mm", "平视手持轻晃 + 焦点转移",
              "兔子跳上窗台蹭他袖口，尾巴一收；他低头看它，目光软了一瞬。", "陆驰（画外）：「主子，明日就是十五了。」／邶（低声）：「备着。」",
              "灵兔跳上窗台，用脑袋蹭他袖口，尾巴往身侧一收，双耳轻颤；男子低头看它，目光软了一瞬。",
              "65mm 中近，平视，轻微手持，焦点在兔耳与他的下颌之间切换一次。",
              "冷月光 + 极淡光晕，兔毛边缘受光，眼中一处高光。",
              "写实动物互动，参考《少年派》；禁止拟人夸张。"),
             ("N03-S04", 3, "大远 ｜ 16mm", "航拍上升 + 穿雾 + 叠化",
              "叠化：晨雾漫起，金线斜插雾中，长街已有人支摊，蒸笼白汽升起。", "无（竹笛明亮起）",
              "天光破晓，晨雾在山谷流动，金线般阳光斜插雾中；依山木屋层层显形，长街已有人支摊，蒸笼白汽升起。",
              "无人机 16mm 大远景，逆光上升 + 轻缓前推。",
              "晨光 4300K 逆光，丁达尔光束清晰，青绿雾气 + 暖金高光。",
              "写实山野晨光，胶片颗粒，参考《卧虎藏龙》。"),
         ]),
    dict(no="N04", time="00:45–01:00", title="抢草 · 冲突",
         func="全片第一个爆点：兔子闯祸，女主登场；先背影再转身，吊 1.5 秒。",
         emo="突发 → 混乱 → 热闹", hook="谁在喊？",
         sound="「嗖」、铜铃乱响、纸包哗啦、音乐一顿", trans="兔子撞怀 → 切女主",
         shots=[
             ("N04-S01", 3, "中 ｜ 35mm", "甩镜 whip pan + 手持（盖·里奇式）",
              "兔子蹬腿窜出，只剩咖白残影掠过草药堆。", "无",
              "灵兔猛地一蹬后腿窜出，只剩一团咖白残影掠过草药堆，草叶被带得飞散。",
              "35mm 甩镜自怀里甩向左侧摊位，运动模糊明显，末尾急停。",
              "环境光；运动段允许拉丝，静止帧清晰。",
              "动态抓拍，参考《疯狂的石头》；不要慢动作。"),
             ("N04-S02", 4, "近 ｜ 50mm", "跟拍 + 升格 90fps（吴宇森式）",
              "兔子叼着一束灵草折返，其中一株叶缘泛青荧。", "无（兔子急喘）",
              "灵兔嘴里叼着一束灵草（其中一株叶缘泛青荧），四肢慌乱蹬地折返，草叶在嘴边晃，铜铃乱响。",
              "50mm 近景跟拍，升格 90fps，轻微手持，焦点跟住草束。",
              "晨光暖调，青荧草拖出一道极淡青色残影。",
              "高速摄影，绒毛与草叶清晰；兔子四足发力真实。"),
             ("N04-S03", 4, "中 ｜ 35mm", "固定 + 轻微后缩",
              "兔子一头扎回他怀里，草束塞进臂弯，耳朵后压。", "无",
              "灵兔一头扎回男子怀里，把草束塞向他臂弯，双耳后压、身体轻抖，心虚又像献宝。",
              "35mm 中景固定，人物占屏 2/3，镜头轻微后缩表现冲击。",
              "晨光，怀抱处柔和阴影，草束在两人之间。",
              "生活化喜剧节奏；不要卡通化反应。"),
             ("N04-S04", 4, "近 ｜ 65mm", "先背影后转身 + 缓推（张艺谋式色块）",
              "浅蓝衣裙少女直起腰，先见背影与银铃；转身时脸颊微红。", "林洛筱：「啊——我的灵草！」（清亮、拔高）",
              "摊位后 22 岁苗疆少女直起腰，浅蓝对襟长裙，先见背影与左耳银铃；她转身，眉眼灵动、脸颊微红，指尖捏着半束被糟蹋的灵草。",
              "65mm 近景：先拍背影 1.5 秒，转身时缓推 2.5 秒，焦点落眼睛。",
              "逆光勾边，浅蓝衣料是全片唯一高明度色块，背景人流虚化。",
              "明亮清新，一眼记住；禁止与男主同一套打光。"),
         ]),
    dict(no="N05", time="01:00–01:15", title="对峙 · 试探",
         func="第一次正面对话：她凶、他笑，并接上「林姑娘」的认亲。",
         emo="嗔怒 → 审视 → 温和", hook="他为什么一点都不恼？",
         sound="环境声退远、兔子「咕」、鼓点一记、轻拨弦", trans="耳语 → 切他开口",
         shots=[
             ("N05-S01", 4, "特 ｜ 85mm", "略仰拍怼脸（竖屏满屏）",
              "她一手叉腰一手指兔，小脸鼓成包子，眼睛圆亮却无凶气。", "「坏小兔，快把我的灵草吐出来！」（气鼓鼓、软糯尾音）",
              "苗疆少女一手叉腰、一手指向兔子，白皙小脸鼓成包子，眉峰拧起，眼睛圆亮却毫无凶气，鼻尖微红，左耳银铃轻晃。",
              "85mm 特写，略低于视平线仰拍，占满竖屏，眼睛压上三分线。",
              "侧逆光 + 面部 40% 补光，留一半阴影。",
              "鲜活娇憨；不要演成撒泼。"),
             ("N05-S02", 5, "中 ｜ 35mm", "随肩线摇 30° + 末帧定格（《一代宗师》对峙）",
              "他停步转身，把她框进画面；两人一高一低，中间隔着半臂与乱草。", "无",
              "男子脚步停稳后缓缓转身，把少女框进画面；两人一高一低相对而立，中间隔着半臂距离与一束凌乱灵草。",
              "35mm 中景，越肩，随转身肩线摇过 30°，末帧定格两人同框。",
              "逆光，两人轮廓被晨光勾亮，中间浮尘可见。",
              "经典相遇构图；不要正面对称站。"),
             ("N05-S03", 6, "近 ｜ 85mm", "过肩 + 耳部特写插入（LCut）",
              "护卫贴耳低语；男子眼帘不动，只喉结动了一下，随后开口。", "陆驰（低声）：「就是那外族女子，林洛筱。」／邶：「想必你就是林姑娘。不知我这灵兔吃了姑娘什么，惹得姑娘气恼？」",
              "护卫上前半步侧身贴耳压嗓低语；男子眼帘不动，只喉结轻动，目光没离开少女，随即温和开口。",
              "85mm 近景过肩，切 100mm 耳部特写，再切回正脸。",
              "半侧光，压低对比，保留耳廓边缘光。",
              "情报私密感转温和试探；口型对齐耳语与开口节奏。"),
         ]),
    dict(no="N06", time="01:15–01:30", title="心动 · 抬眼",
         func="3 秒完成关系建立；封面与切片首选。",
         emo="慌 → 羞 → 倔", hook="心跳声，是谁的？",
         sound="心跳两记（60Hz）、环境声虚化、单音钢琴", trans="对视 → 她先移开眼",
         shots=[
             ("N06-S01", 4, "特 ｜ 100mm", "微距手部 + 上摇至下颌（《色戒》手部戏）",
              "指尖倏地收回，缩进掌心，又攥住裙摆，指节泛白。", "无",
              "她指尖倏地收回，缩进掌心，又攥住身前裙摆，指节泛白，布料被捏出细褶；指腹有淡淡药渍。",
              "100mm 微距手部，焦点在指节，缓慢上摇到下颌。",
              "侧光，手部质感清晰，裙料纤维可见。",
              "细节特写；不要美甲、不要现代饰品。"),
             ("N06-S02", 4, "近 ｜ 85mm", "主观镜头 POV + 缓推（莱昂内式）",
              "她的视角：他逆光站着，肩头金边，睫毛下的眼睛很静。", "无",
              "少女主观视角：男子逆光站立，晨光在肩头勾出金边，睫毛下的眼睛很静，身后人流虚化成色块。",
              "85mm 主观镜头，轻微呼吸手持，焦点在他眼睛，边缘柔焦。",
              "强逆光 + 轻眩光，高光柔散，整体偏暖。",
              "「一眼万年」柔光视角；保留真实肤质。"),
             ("N06-S03", 4, "特 ｜ 100mm", "慢推 5% + 焦点锁睫毛（封面帧）",
              "她抬眼撞进他眼睛：瞳孔一颤，睫毛乱两拍，耳根渐红。", "无",
              "少女抬眼正撞进他的眼睛：瞳孔先一颤，睫毛乱了两拍，耳根到脸颊一点点染上绯红，唇微张又抿住。",
              "100mm 特写，慢推 5%，焦点锁睫毛与眼白。",
              "侧逆光，睫毛与碎发透光，脸红靠妆与光。",
              "封面候选帧：抬眼后、低头前 12–18 帧。"),
             ("N06-S04", 3, "中近 ｜ 50mm", "固定 + 极缓收（留白）",
              "她垂下眼，声音软下来；他没动，只看着她。", "林（低声）：「……它吃了我很贵重的灵草。」",
              "她垂下眼，声音软下来，指尖仍攥着裙摆；男子没有动，只看着她。",
              "50mm 中近固定，末帧极缓后拉 5% 留白。",
              "侧光转柔，背景虚成暖色块。",
              "克制的收尾；留气口给下一单元报价。"),
         ]),
    dict(no="N07", time="01:30–01:45", title="报账 · 送别（合并原报账+递钱+送别）",
         func="立住「不是贪财是认死理」；36/40 做记忆点；双铃同框做声音伏笔。",
         emo="较真 → 好笑 → 怅然", hook="多给的四枚银币，算什么？",
         sound="竹笛俏皮、银币碰撞、人群笑、银铃、铜铃、竹帘落下", trans="竹帘落下 → 切他",
         shots=[
             ("N07-S01", 2, "近 ｜ 85mm", "正反打固定",
              "他低头看她，眉梢微挑。", "邶：「敢问姑娘，这灵草多少钱？」",
              "男子低头看她，眉梢微挑，像在问一件不相干的小事。",
              "85mm 近景，正反打固定。",
              "柔和侧光，肤色干净。",
              "轻松调侃，状态松弛。"),
             ("N07-S02", 3, "近 ｜ 65mm", "微俯 + 随手指下压",
              "三根手指一根根掰开，眉头认真，嘴唇随数字轻动。", "林：「一株十二银币。它吃了我三株——就是三十六银币。」",
              "少女伸出手，三根手指一根一根掰开，眉头认真，嘴唇随数字轻开合；另一只手护着被啃坏的草束。",
              "65mm 近景，微俯，随手指下压轻微跟随，末帧切手部。",
              "顺光，手部细节清晰，指腹药渍可见。",
              "生活化较真；数字口型清晰。"),
             ("N07-S03", 3, "中全 ｜ 35mm", "固定 + 前后错位站位",
              "他偏头吩咐，护卫取出绣云纹小钱袋，银币轻响。", "邶：「赔林姑娘四十银币。多出来的，算这顽兔的赔罪。」",
              "男子偏头吩咐，护卫点头取出绣云纹小钱袋，银币碰撞轻响；少女在前景伸手要接。",
              "35mm 中全固定，前后错位站位，前景少女手臂成框。",
              "明亮晨光，钱袋银线微闪。",
              "交易仪式感；银币质感要真。"),
             ("N07-S04", 3, "特 ｜ 100mm", "微距俯拍 + 末帧焦点上移",
              "两只手交接钱袋；他顺势俯身靠近，声音放轻。", "邶（低声）：「至于怎认得姑娘——昨日你治好了我族的大祭司。」",
              "两只手在画面正中交接钱袋：一只白净戴墨玉戒，一只带细伤与药渍；男子顺势俯身靠近，声音放轻。",
              "100mm 微距俯拍，末帧焦点从钱袋上移到她指尖。",
              "柔光，皮肤纹理与布料纤维清晰。",
              "手部关系镜头；不要出现多余手指。"),
             ("N07-S05", 4, "全 ｜ 24mm", "升格 60fps 侧移跟拍 + 长焦收尾",
              "她后退致歉又站直，转身离去；银铃与铜铃先后一响；他目送。", "林：「治病救人乃是医者本分。用得上的地方——就来大祭司处寻我。」／邶（低声）：「倒是有趣。」",
              "她后退两步轻撞簸箕，慌忙道歉随即理裙挺背站定，转身离去，浅蓝裙袂荡开、银铃一晃；怀里灵兔抬头，铜铃应了一声；男子立在原地目送，帘子落下。",
              "24mm 全景侧移跟拍，升格 60fps；末帧切 135mm 长焦拍背影被竹帘吞没。",
              "逆光，裙料透出浅蓝高光，是全画面唯一清亮色块。",
              "色彩记忆点，参考《英雄》色块调度；人群做灰调衬托。"),
         ]),
    dict(no="N08", time="01:45–02:00", title="灵草 · 钩子",
         func="全集唯一信息差：被她按十二银币卖掉的草，正是他的药；今晚就是十五。",
         emo="松弛 → 骤冷", hook="今夜十五——他撑得住吗？",
         sound="古琴一声 → 骤停 → 低频闷响 → 余音", trans="急拉失焦 → 第 12 帧切黑 → 字幕「今夜，十五」",
         shots=[
             ("N08-S01", 4, "特 ｜ 100mm", "微距浅景深",
              "兔子嘴边挂着半片草叶；他捻起叶子，叶脉透出极淡青色。", "无",
              "兔子嘴边挂着半片被咬过的草叶；男子伸手捻起，叶脉在光下透出极淡青色，叶缘有细密锯齿。",
              "100mm 微距，浅景深，焦点在叶脉，背景市集虚成光斑。",
              "侧逆光穿叶，叶脉半透明青绿，其余暖调。",
              "植物微距，真实叶脉结构；不要发光宝石。"),
             ("N08-S02", 4, "特 ｜ 100mm", "极缓推至叶缘 + 焦点转移",
              "他把叶子举向光，瞳孔一缩；叶缘浮出一线极淡青荧。", "无",
              "他把叶子举向光，瞳孔一缩；叶缘在光下浮出一线极淡青荧，与普通灵草截然不同。",
              "100mm 微距，极缓推至叶缘（屏宽 15%），末帧焦点从叶缘转到他瞳孔。",
              "叶缘极淡青荧，是画面唯一冷光源；背景曝掉半档。",
              "识别性特写；青荧克制，像薄霜反光。"),
             ("N08-S03", 4, "近 ｜ 85mm + 腕部插入", "固定 + 画中画插入",
              "他笑意一点点退下去：嘴角先松，再是眼睛；腕上旧纹一亮一暗。", "无",
              "男子脸上笑意一点点退下去：嘴角先松，再是眼睛；画面右下角插入手腕，旧年纹路随心跳一亮一暗。",
              "85mm 近景固定，右下角画中画插入腕部微距，构图留 1/5 空间。",
              "冷调渐起，暖光被压下去，纹路微光成唯一亮点。",
              "情绪反转特写；不要夸张惊讶，是「心往下沉」。"),
             ("N08-S04", 3, "全 ｜ 35mm", "急拉 + 失焦 + 硬切黑（库布里克式）",
              "镜头猛地拉开，市集在他身后失焦；山脊上一轮满月正升起。", "邶（画外，低声）：「青萤……能治旧疾？」／字幕：「今夜，十五。」",
              "镜头猛地拉开，喧闹市集在男子身后迅速失焦；远处山脊上一轮满月正升起，他的影子被拉得很长，立于画面正中。",
              "35mm 全景，急拉（0.8 秒）+ 快速失焦，第 12 帧硬切黑。",
              "15 秒内由晨光转冷：暖金褪去，冷蓝压上来。",
              "库布里克式单点透视收尾；黑场留 12 帧，画外音落第一帧。"),
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
set_run(hp.add_run("《第一章 · 初遇》V4 ｜ 120 秒精炼版 ｜ 8 单元 / 32 镜"), name="微软雅黑", size=8, color="8C9AA3")

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
p.paragraph_format.space_before = Pt(24)
set_run(p.add_run("SHORT DRAMA · STORYBOARD V4 · 120s"), name="微软雅黑", size=8.5, bold=True, color=C_GOLD)
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(6)
set_run(p.add_run("第一章 · 初遇"), name="微软雅黑", size=34, bold=True, color=C_INK)
p = doc.add_paragraph()
set_run(p.add_run("V4 · 120 秒精炼版 ｜ 8 个 15 秒单元 ｜ 32 镜 ｜ 精炼场景 + 精确提示词"), name="微软雅黑", size=12, color="4A6675")
p = doc.add_paragraph()
set_run(p.add_run("前 30 秒一秒不动；后 90 秒把 V3 的 8 个单元压成 6 个，镜数 40 → 32，场景描写砍到一句。"), name="宋体", size=11, color="5C7A88")

rows = [
    ("版本", "V4（V3 150 秒 → 120 秒；前 30 秒完全沿用 V3）"),
    ("时长", "120 秒（2′00″）＝ 前 30s（N01–N02）+ 后 90s（N03–N08）"),
    ("结构", "8 个 15 秒单元；每单元 3–5 镜，共 32 镜，平均 3.75s/镜"),
    ("画幅", "9:16 竖屏 1080×1920，安全区上 1/6、下 1/5"),
    ("本轮删减", "合并 2 组：月下+破晓+集市 → N03；报账+递钱+送别 → N07；场景描写一律压到一句"),
    ("提示词", "四行、每行 25–45 字；参数化（焦段/色温/光源方位/运动），可验证，不写抽象情绪"),
    ("合规", "不写毒、汗、伤、血、药；用「旧纹 / 旧疾 / 空碗冷汤」"),
]
t = doc.add_table(rows=0, cols=2)
table_borders(t)
cell_margins(t, 50, 50, 90, 90)
for k, v in rows:
    row = t.add_row()
    no_split(row)
    cell_text(row.cells[0], k, name="微软雅黑", size=8.5, bold=True, color=C_INK, align=1)
    cell_text(row.cells[1], v, size=9, color="2B3B44")
    shade(row.cells[0], C_SOFT)
fixed(t, [1.5, 8.9])

heading(doc, "一、删减对照：150s → 120s", 1, new_page=True)
para(doc, "原则没变：**不删剧情节点，只合并过渡**。前 30 秒一秒不动（红果前 3 秒钩子的本钱）；后段从 8 个单元 40 镜压到 6 个单元 24 镜。", size=9, after=6)

mapping = [
    ("V3 单元", "V4 单元", "处理", "镜数"),
    ("N01 寒月毒兆", "N01", "原样保留", "4 → 4"),
    ("N02 密报试探", "N02", "原样保留", "4 → 4"),
    ("N03 月下·破晓", "N03", "扩并开市：推窗 4 + 空碗 4 + 兔 4 + 叠化晨雾与长街 3", "4 → 4（并入原 N04）"),
    ("N04 集市烟火", "N03", "并入：只保留「叠化后长街开市」3 秒，独立单元取消", "4 → 0"),
    ("N05 抢草冲突", "N04", "四镜不变，删闲笔", "4 → 4"),
    ("N06 对峙嗔怒", "N05", "三镜：鼓腮 / 转身同框 / 耳语+认亲（原「他开口」并入末镜）", "4 → 3"),
    ("N07 心动抬眼", "N06", "四镜不变，末镜加「它吃了我很贵重的灵草」作过渡", "4 → 4"),
    ("N08 报账较真", "N07", "并入送别：报价 2 + 掰指 3 + 赔银 3 + 递钱俯身 3 + 转身双铃 4", "4 → 5（并入原 N09）"),
    ("N09 递钱·送别", "N07", "同上，双铃与「倒是有趣」压进末镜", "4 → 0"),
    ("N10 灵草钩子", "N08", "保留完整 15s", "4 → 4"),
    ("合计", "8 单元", "150s → 120s", "40 → 32"),
]
t = doc.add_table(rows=0, cols=4)
table_borders(t, sz=3)
cell_margins(t, 34, 34, 60, 60)
for i, r in enumerate(mapping):
    row = t.add_row()
    no_split(row)
    if i == 0:
        for j, v in enumerate(r):
            cell_text(row.cells[j], v, name="微软雅黑", size=7.8, bold=True, color=C_WHITE, align=1)
            shade(row.cells[j], C_INK)
        repeat_hdr(row)
        continue
    cell_text(row.cells[0], r[0], name="微软雅黑", size=7.6, bold=True, color=C_INK)
    cell_text(row.cells[1], r[1], name="微软雅黑", size=7.6, color="2F5A6B", align=1)
    cell_text(row.cells[2], r[2], size=7.4, color="22323A")
    cell_text(row.cells[3], r[3], size=7.4, color="6B3A2E", align=1)
fixed(t, [2.2, 1.4, 5.5, 1.3])

para(doc, "**被删掉的是「铺陈」，不是「信息」**：集市烟火独立单元（保留长街 3 秒）、露珠空镜、门帘特写、报账的第二轮追问、送别单独的目送镜头。"
          "仍然一个不落：倒计时、兔子护主、抢草、心动三秒、36/40 银币、双铃、「倒是有趣」、灵草钩子。", size=8.8, color="2B3B44", before=6)

heading(doc, "二、逐单元分镜与提示词（8 单元 / 32 镜）", 1, new_page=True)
para(doc, "每单元一页：单元卡 → 分镜表 → 3–5 条提示词。运镜栏标大师参照；提示词按「可验证」原则写，每行一个可核对的事实。", size=8.8, after=4)

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
    cell_margins(st, 20, 20, 50, 50)
    for j, txt in enumerate(("镜号", "时长", "景别 · 焦段", "运镜（大师参照）", "画面 · 动作（精炼）", "台词 / 声音")):
        cell_text(st.rows[0].cells[j], txt, name="微软雅黑", size=7.0, bold=True, color=C_WHITE, align=1)
        shade(st.rows[0].cells[j], C_INK)
    repeat_hdr(st.rows[0])
    no_split(st.rows[0])
    for s in u["shots"]:
        row = st.add_row()
        no_split(row)
        cell_text(row.cells[0], s[0].split("-")[1], name="微软雅黑", size=7.2, bold=True, color=C_INK, align=1)
        cell_text(row.cells[1], f"{s[1]}s", name="微软雅黑", size=7.2, color="2B3B44", align=1)
        cell_text(row.cells[2], s[2], size=6.8, color="2B3B44")
        cell_text(row.cells[3], s[3], size=6.8, color="3A4A52")
        cell_text(row.cells[4], s[4], size=6.9, color="1F2E36")
        cell_text(row.cells[5], s[5], size=6.8, color="6B3A2E")
    fixed(st, [0.62, 0.42, 1.05, 1.55, 4.85, 2.01])

    ph = doc.add_paragraph()
    ph.paragraph_format.space_before = Pt(2)
    ph.paragraph_format.space_after = Pt(0)
    ph.paragraph_format.line_spacing = 1.0
    ph.paragraph_format.keep_with_next = True
    set_run(ph.add_run(f"▍{u['no']} 提示词（{len(u['shots'])} 条）"), name="微软雅黑", size=7.2, bold=True, color=C_GOLD)
    for s in u["shots"]:
        prompt_block(doc, s[0], s[1], s[2], s[6], s[7], s[8], s[9])

heading(doc, "三、时长核对与节奏", 1, new_page=True)
total = sum(s[1] for u in UNITS for s in u["shots"])
rows2 = [("单元", "时间码", "时长", "累计", "本单元钩子")]
acc = 0
for u in UNITS:
    d = sum(s[1] for s in u["shots"])
    acc += d
    rows2.append((u["no"], u["time"], f"{d}s", f"{acc}s", u["hook"]))

t = doc.add_table(rows=0, cols=5)
table_borders(t, sz=3)
cell_margins(t, 34, 34, 60, 60)
for i, r in enumerate(rows2):
    row = t.add_row()
    no_split(row)
    if i == 0:
        for j, v in enumerate(r):
            cell_text(row.cells[j], v, name="微软雅黑", size=7.8, bold=True, color=C_WHITE, align=1)
            shade(row.cells[j], C_INK)
        repeat_hdr(row)
        continue
    cell_text(row.cells[0], r[0], name="微软雅黑", size=7.6, bold=True, color=C_INK, align=1)
    cell_text(row.cells[1], r[1], name="微软雅黑", size=7.6, color="2B3B44", align=1)
    cell_text(row.cells[2], r[2], size=7.6, color="2B3B44", align=1)
    cell_text(row.cells[3], r[3], size=7.6, color="2B3B44", align=1)
    cell_text(row.cells[4], r[4], size=7.6, color="6B3A2E")
fixed(t, [0.7, 1.15, 0.7, 0.8, 6.6])

para(doc, f"合计 {total} 秒 = 2′00″，8 单元 × 15 秒，与目标 120 秒一致；32 镜，平均 {total/32:.2f} 秒/镜。", size=9, bold=True, color=C_INK, before=6, after=6)

para(doc, "120 秒版的三条节奏调整", name="微软雅黑", size=9.5, bold=True, color="2F5A6B", before=4, after=3)
for s in [
    "**前 30 秒不变**：0–3 秒仍是黑影掠顶 + 月轮，第 2 单元仍在 30 秒前完成「盯着她」的收网。",
    "**中段提速**：N03 一个单元完成「夜→晨→开市」，N05 三镜完成对峙与认亲，信息密度比 150 秒版高约 25%。",
    "**后段合流**：N07 五镜把报价、掰手指、赔银、俯身试探、转身双铃一次做完；N08 独占 15 秒做钩子，不抢拍。",
    "**结尾不变**：硬切黑 + 字幕「今夜，十五」，黑场 12 帧，画外音落第一帧。",
]:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.left_indent = Pt(12)
    for _i, _seg in enumerate(("· " + s).split("**")):
        if _seg:
            set_run(p.add_run(_seg), size=8.8, bold=(_i % 2 == 1), color="2B3B44")

heading(doc, "四、与 V3 的差异（供对比）", 2)
para(doc, "V3：10 单元 / 40 镜 / 150 秒 ｜ V4：8 单元 / 32 镜 / 120 秒。V4 少 30 秒，靠三处压缩：① 集市独立单元并入开市；"
          "② 报账与送别合成一个 5 镜单元；③ 每个镜头的场景描写从 2 句压到 1 句，提示词从 40–60 字压到 25–45 字。"
          "人物设定、合规用词、视觉锚点三处与 V2/V3 完全一致，可直接混用素材。", size=8.8, color="2B3B44")

doc.core_properties.title = "第一章 初遇 · 分镜脚本 V4（120 秒精炼版）"
doc.core_properties.author = "编剧组"
doc.save(OUT)
print("saved:", OUT)
print("units:", len(UNITS), "shots:", sum(len(u["shots"]) for u in UNITS), "total_sec:", total)
