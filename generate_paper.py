# -*- coding: utf-8 -*-
"""
生成《世界贸易组织概论》期末论文 Word 文档（含图、表）。

题目：入世是"污染天堂"还是"技术外溢"？
      ——基于跨国面板数据对 WTO 成员身份与出口 CO2 强度关系的实证检验

输出：WTO期末论文_入世与出口碳强度.docx
依赖：python-docx；需先运行 generate_figures.py 生成 fig1-fig5.png。
正文（含图表说明）约 3500 字。
"""

import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

CN_FONT = "宋体"
CN_HEI = "黑体"
EN_FONT = "Times New Roman"


def set_run_font(run, cn=CN_FONT, en=EN_FONT, size=12, bold=False, color=None):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = en
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), en)
    rfonts.set(qn("w:hAnsi"), en)
    rfonts.set(qn("w:eastAsia"), cn)
    if color is not None:
        run.font.color.rgb = color


def add_body(doc, text, size=12, first_indent=True, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
             space_after=0, line=1.5, bold=False, cn=CN_FONT):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = line
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(0)
    if first_indent:
        pf.first_line_indent = Pt(size * 2)
    run = p.add_run(text)
    set_run_font(run, cn=cn, size=size, bold=bold)
    return p


def add_h1(doc, text):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(12); pf.space_after = Pt(6)
    pf.line_spacing = 1.5; pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    run = p.add_run(text)
    set_run_font(run, cn=CN_HEI, size=14, bold=True)
    return p


def add_h2(doc, text):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(6); pf.space_after = Pt(3)
    pf.line_spacing = 1.5; pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    run = p.add_run(text)
    set_run_font(run, cn=CN_HEI, size=12.5, bold=True)
    return p


def add_centered(doc, text, size=12, bold=False, cn=CN_FONT, space_after=0, space_before=0):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.space_after = Pt(space_after); pf.space_before = Pt(space_before)
    pf.line_spacing = 1.5; pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    run = p.add_run(text)
    set_run_font(run, cn=cn, size=size, bold=bold)
    return p


def add_caption(doc, text, before=4, after=2):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.space_before = Pt(before); pf.space_after = Pt(after); pf.line_spacing = 1.2
    run = p.add_run(text)
    set_run_font(run, cn=CN_HEI, size=10.5, bold=True)
    return p


def add_image(doc, path, width_cm=14.8):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(2)
    run = p.add_run()
    run.add_picture(path, width=Cm(width_cm))
    return p


def add_note(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf = p.paragraph_format
    pf.space_after = Pt(6); pf.line_spacing = 1.2
    run = p.add_run(text)
    set_run_font(run, size=9, cn=CN_FONT)
    return p


def style_cell(cell, text, size=10, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_after = Pt(0); p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold)


def add_table(doc, headers, rows, col_align=None, header_size=10, body_size=10):
    n_cols = len(headers)
    table = doc.add_table(rows=1, cols=n_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    for j, h in enumerate(headers):
        style_cell(table.rows[0].cells[j], h, size=header_size, bold=True)
    for r in rows:
        cells = table.add_row().cells
        for j, val in enumerate(r):
            al = col_align[j] if col_align else WD_ALIGN_PARAGRAPH.CENTER
            style_cell(cells[j], val, size=body_size, align=al)
    return table


def add_page_number_footer(doc):
    section = doc.sections[-1]
    p = section.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    f1 = OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = "PAGE"
    f2 = OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"), "end")
    run._element.append(f1); run._element.append(it); run._element.append(f2)
    set_run_font(run, size=10.5)


def main():
    doc = Document()
    sec = doc.sections[0]
    sec.page_height = Cm(29.7); sec.page_width = Cm(21.0)
    sec.top_margin = Cm(2.54); sec.bottom_margin = Cm(2.54)
    sec.left_margin = Cm(3.0); sec.right_margin = Cm(3.0)

    style = doc.styles["Normal"]
    style.font.name = EN_FONT; style.font.size = Pt(12)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), CN_FONT)

    # ===================== 封面 =====================
    for _ in range(2):
        add_centered(doc, "")
    add_centered(doc, "对外经济贸易大学", size=22, bold=True, cn=CN_HEI, space_after=6)
    add_centered(doc, "《世界贸易组织概论》期末论文", size=18, bold=True, cn=CN_HEI, space_after=30)
    for _ in range(2):
        add_centered(doc, "")
    add_centered(doc, "入世是\u201c污染天堂\u201d还是\u201c技术外溢\u201d？", size=16, bold=True, cn=CN_HEI, space_after=4)
    add_centered(doc, "——基于跨国面板数据对 WTO 成员身份与出口 CO\u2082 强度关系的实证检验",
                 size=13, bold=True, cn=CN_HEI, space_after=40)
    for _ in range(3):
        add_centered(doc, "")
    info = [
        ("教学单位", "经贸学院"), ("课程名称", "世界贸易组织概论"),
        ("课程代码", "ITR305"), ("课序号", "2"),
        ("任课教师", "吕越、宁静馨"), ("考试方式", "论文"),
        ("学生姓名", "＿＿＿＿＿＿＿＿＿＿"), ("学    号", "＿＿＿＿＿＿＿＿＿＿"),
        ("学    期", "2026—2027 学年度第一学期"),
    ]
    for k, v in info:
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = 1.6; p.paragraph_format.space_after = Pt(2)
        r1 = p.add_run(f"{k}："); set_run_font(r1, size=14, bold=True, cn=CN_HEI)
        r2 = p.add_run(v); set_run_font(r2, size=14)
    doc.add_page_break()

    # ===================== 摘要 =====================
    add_centered(doc, "入世是\u201c污染天堂\u201d还是\u201c技术外溢\u201d？", size=15, bold=True, cn=CN_HEI, space_after=4)
    add_centered(doc, "——基于跨国面板数据对 WTO 成员身份与出口 CO\u2082 强度关系的实证检验",
                 size=11, bold=True, cn=CN_HEI, space_after=12)
    add_h2(doc, "摘  要")
    abstract = (
        "加入世界贸易组织（WTO）通过削减关税、约束壁垒与稳定贸易预期，深刻重塑了成员的出口结构，"
        "由此引发长期争论：贸易自由化究竟会使发展中国家沦为高碳产业的\u201c污染天堂\u201d，"
        "还是通过竞争与技术外溢促使出口部门\u201c变绿\u201d？本文以 2000—2020 年覆盖约 160 个经济体、"
        "共 3480 个国家—年观测的跨国面板数据为样本，以单位出口额所含 CO\u2082 排放（出口碳强度）为被解释变量，"
        "构建双向固定效应模型并辅以事件研究、异质性分析与安慰剂检验。研究发现：第一，混合 OLS 得到的显著负向关系"
        "在引入国家与年份固定效应后消失，WTO 成员身份系数为 0.111（p=0.283），统计上不显著，安慰剂检验表明其与随机赋值难以区分；"
        "第二，出口碳强度在样本期内普遍大幅下降，但主要由收入增长（环境库兹涅茨曲线）、能源强度改善与产业结构调整驱动；"
        "第三，效应存在异质性：高能源强度国家与发展中国家呈弱\u201c污染天堂\u201d特征，低能源强度国家呈弱\u201c技术外溢\u201d特征，"
        "但均不具统计稳健性。本文据此提出以\u201c贸易—环境—技术\u201d政策协同推动出口低碳化的建议。"
    )
    add_body(doc, abstract, size=11)
    p = doc.add_paragraph(); p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.first_line_indent = Pt(22)
    r1 = p.add_run("关键词："); set_run_font(r1, size=11, bold=True, cn=CN_HEI)
    r2 = p.add_run("世界贸易组织；出口碳强度；污染天堂假说；技术外溢；双向固定效应")
    set_run_font(r2, size=11)
    doc.add_page_break()

    # ===================== 引言 =====================
    add_h1(doc, "一、引言")
    add_body(doc,
        "自 1995 年世界贸易组织（WTO）成立以来，以最惠国待遇、国民待遇和关税约束为核心的多边贸易体制"
        "极大降低了跨国贸易成本，推动全球贸易额与生产网络迅速扩张。中国于 2001 年正式加入 WTO，"
        "此后出口规模跃居世界首位，发展中国家在全球出口中的份额亦显著上升。然而，贸易扩张在带来增长红利的同时，"
        "也使\u201c贸易与环境\u201d的关系成为各方关注的焦点。")
    add_body(doc,
        "围绕这一关系存在两种针锋相对的判断。一种是\u201c污染天堂假说\u201d：在环境规制存在国别差异时，"
        "贸易自由化会促使高碳生产环节向规制宽松的发展中国家转移，使其出口部门\u201c变脏\u201d、单位出口碳排放上升。"
        "另一种是\u201c技术外溢\u201d观点：融入多边体制会带来更激烈的竞争、更严格的进口国标准以及更便利的技术与资本流入，"
        "从而倒逼并帮助出口企业提升能效、降低碳强度。两种机制方向相反，孰强孰弱本质上是一个有待数据检验的实证问题。")
    add_body(doc,
        "本文聚焦一个可度量、政策含义清晰的指标——出口碳强度（单位出口额所含 CO\u2082 排放），"
        "检验 WTO 成员身份对该指标的因果影响。边际贡献在于：其一，直接以\u201c成员身份\u201d这一制度变量为处理变量，"
        "借助双向固定效应、事件研究与安慰剂检验缓解内生性；其二，区分要素禀赋与发展阶段，刻画效应的异质性，"
        "从而回应\u201c污染天堂\u201d与\u201c技术外溢\u201d之争。")

    # ===================== 文献综述 =====================
    add_h1(doc, "二、文献综述")
    add_body(doc,
        "贸易与环境关系的研究可追溯至 Grossman 和 Krueger（1991）提出的规模、结构与技术三种效应分解框架："
        "贸易扩张通过扩大规模增加排放、通过改变产业构成影响排放、通过引入清洁技术降低单位排放，净方向取决于三者合力。"
        "环境库兹涅茨曲线（EKC）假说进一步认为，污染随人均收入呈先升后降的倒 U 形，为理解碳强度的长期演化提供了参照。")
    add_body(doc,
        "\u201c污染天堂\u201d文献方面，Copeland 和 Taylor（1994、2004）从理论上论证了环境规制差异如何塑造污染密集型产业的国际分工，"
        "并指出其经验证据常因要素禀赋效应的对冲而较弱；Levinson、Antweiler 等的实证结论分歧明显。与之相对，"
        "Frankel 和 Rose（2005）则提供了贸易可改善环境质量的证据。就 WTO/GATT 的作用而言，Rose（2004）曾质疑其贸易促进效应，"
        "引发关于成员身份效应的长期讨论。总体看，既有研究多以贸易开放度为自变量，较少直接评估\u201c成员身份\u201d这一离散制度冲击"
        "对出口碳强度的影响，且对异质性的刻画不足，本文力图在这两方面有所补充。")

    # ===================== 理论与假设 =====================
    add_h1(doc, "三、理论分析与研究假设")
    add_body(doc,
        "加入 WTO 对出口碳强度的影响可经三条路径传导：结构路径上，关税减让使一国按比较优势重配资源，"
        "若其优势位于碳密集产业，出口结构\u201c变脏\u201d、碳强度上升，对应\u201c污染天堂\u201d机制；"
        "技术路径上，进口中间品与外资的技术外溢、叠加出口目的国的环境标准，推动企业采用更清洁工艺、碳强度下降，"
        "对应\u201c技术外溢\u201d机制；规模路径则放大前两者的作用。由于结构与技术效应方向相反，二者强弱取决于成员国的禀赋与发展阶段。"
        "据此提出竞争性假设：")
    add_body(doc, "H1a（污染天堂）：加入 WTO 后，成员出口碳强度上升；", first_indent=True)
    add_body(doc, "H1b（技术外溢）：加入 WTO 后，成员出口碳强度下降；", first_indent=True)
    add_body(doc,
        "H2（异质性）：能源强度高、处于工业化中期的发展中国家更可能呈\u201c污染天堂\u201d特征；"
        "能源强度低、技术水平高的经济体更可能呈\u201c技术外溢\u201d特征。", first_indent=True)

    # ===================== 数据与方法 =====================
    add_h1(doc, "四、数据来源与研究设计")
    add_h2(doc, "（一）数据与变量")
    add_body(doc,
        "本文构建 2000—2020 年跨国面板，覆盖约 160 个经济体，经缺失值处理后共 3480 个国家—年观测。"
        "核心被解释变量为出口碳强度，定义为 CO\u2082 排放量与货物和服务出口额之比（吨/百万美元），并取自然对数 ln(CO\u2082 强度)；"
        "核心解释变量为 WTO 成员身份虚拟变量。控制变量包括人均 GDP 及其平方项（检验 EKC）、能源强度、工业增加值占比。"
        "出口额来自世界银行 WDI（指标 NE.EXP.GNFS.CD），人均 GDP、能源强度、工业占比同样取自世界银行数据库，"
        "CO\u2082 排放数据来自公开排放数据库。表 1 报告主要变量的描述性统计。")

    add_caption(doc, "表 1  主要变量描述性统计")
    add_table(doc,
        ["变量", "观测值", "均值", "标准差", "最小值", "中位数", "最大值"],
        [
            ["CO\u2082强度（吨/百万美元）", "3480", "2246.74", "3123.27", "1.14", "1349.75", "45695.13"],
            ["ln(CO\u2082强度)", "3480", "7.137", "1.182", "0.127", "7.208", "10.730"],
            ["WTO成员（0/1）", "3480", "0.792", "0.406", "0", "1", "1"],
            ["CO\u2082排放量（百万吨）", "3480", "188.36", "833.32", "0.001", "16.03", "11964.13"],
            ["人均GDP（2015美元）", "3452", "14395.0", "19659.5", "287.4", "4993.6", "118382.9"],
            ["能源强度", "3418", "5.231", "3.213", "0.110", "4.330", "30.440"],
            ["工业占GDP比重（%）", "3392", "27.30", "12.35", "3.72", "25.03", "86.67"],
        ],
        col_align=[WD_ALIGN_PARAGRAPH.LEFT] + [WD_ALIGN_PARAGRAPH.CENTER]*6,
        body_size=9.5, header_size=9.5)
    add_note(doc, "注：CO\u2082 强度右偏明显，故取对数；WTO 成员观测占比 79.2%。")

    add_body(doc,
        "图 1 直观呈现了样本特征：图 1a 比较处理组（2000 年后入世）与控制组（1995 年已为成员）的 ln(CO\u2082 强度) 趋势，"
        "两组均呈下降态势；图 1b 显示 ln(CO\u2082 强度) 大致对称、均值约 7.14。图 2 进一步对比主要出口国 2000 年与 2020 年的碳强度，"
        "中国、印度、俄罗斯、南非等国普遍下降六成以上，呈现明显的全局性\u201c变绿\u201d趋势。")
    add_image(doc, "fig1.png")
    add_caption(doc, "图 1  描述性统计：处理组/控制组趋势与 ln(CO\u2082 强度) 分布")
    add_image(doc, "fig2.png")
    add_caption(doc, "图 2  主要国家出口 CO\u2082 强度：2000 年与 2020 年对比")

    add_h2(doc, "（二）计量模型")
    add_body(doc,
        "为识别成员身份的净效应，设定双向固定效应模型：")
    add_body(doc,
        "ln(CO\u2082 强度)_{it} = β\u2081·WTO_{it} + β\u2082·ln(pgdp)_{it} + β\u2083·[ln(pgdp)]\u00b2_{it} "
        "+ β\u2084·能源强度_{it} + β\u2085·工业占比_{it} + μ_i + λ_t + ε_{it}",
        first_indent=False, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_body(doc,
        "其中 μ_i 为国家固定效应（吸收不随时间变化的禀赋与制度差异），λ_t 为年份固定效应（吸收全球技术进步、油价等共同冲击），"
        "β\u2081 即所关注的成员身份效应；标准误在国家层面聚类。在此基础上，采用事件研究检验平行趋势与动态效应，"
        "通过分组回归刻画异质性，并以安慰剂检验等方式检验稳健性。")

    # ===================== 实证结果 =====================
    add_h1(doc, "五、实证结果与讨论")
    add_h2(doc, "（一）基准回归")
    add_body(doc,
        "表 2 报告基准回归。混合 OLS 中 WTO 系数为 -0.318（p<0.05），似乎支持\u201c技术外溢\u201d；"
        "但该设定未控制国家异质性，存在严重选择性偏误——本身碳强度较低的发达经济体往往更早成为成员，"
        "从而高估负向关系。加入国家固定效应后系数降至 -0.025 且不显著；进一步加入年份固定效应（双向 FE）后，"
        "系数变为 0.111、p=0.283，不显著。这表明剥离国家禀赋与全球共同趋势后，成员身份对出口碳强度并无稳健因果影响。")
    add_caption(doc, "表 2  基准回归结果（被解释变量：ln(CO\u2082 强度)）")
    add_table(doc,
        ["变量", "(1) 混合OLS", "(2) 国家FE", "(3) 双向FE"],
        [
            ["WTO成员（β\u2081）", "-0.3181**\n(0.1388)", "-0.0251\n(0.1333)", "0.1108\n(0.1031)"],
            ["ln(人均GDP)", "2.9522***\n(0.5117)", "2.1157**\n(0.8535)", "2.9030***\n(0.7438)"],
            ["[ln(人均GDP)]\u00b2", "-0.1861***\n(0.0289)", "-0.2085***\n(0.0501)", "-0.2072***\n(0.0427)"],
            ["能源强度", "0.0968***\n(0.0231)", "0.0470***\n(0.0135)", "0.0433***\n(0.0148)"],
            ["工业增加值占比", "0.0018\n(0.0049)", "-0.0111***\n(0.0041)", "-0.0173***\n(0.0030)"],
            ["国家固定效应", "否", "是", "是"],
            ["年份固定效应", "否", "否", "是"],
        ],
        col_align=[WD_ALIGN_PARAGRAPH.LEFT] + [WD_ALIGN_PARAGRAPH.CENTER]*3,
        body_size=9.5, header_size=10)
    add_note(doc, "注：括号内为国家层面聚类稳健标准误；*、**、*** 分别表示在 10%、5%、1% 水平显著。")
    add_body(doc,
        "控制变量与理论高度吻合：ln(人均 GDP) 显著为正、其平方项显著为负，清晰刻画出碳强度随收入先升后降的倒 U 形 EKC；"
        "能源强度系数显著为正，说明能源效率是碳强度的关键决定因素；工业占比系数显著为负。"
        "可见样本期内出口碳强度的普遍下降，主要由收入增长、能效提升与结构调整解释，而非入世这一制度事件本身。")

    add_h2(doc, "（二）事件研究与动态效应")
    add_body(doc,
        "图 3 的事件研究显示，入世前各期系数均接近零且不显著，说明处理组与对照组不存在系统性差异趋势，"
        "平行趋势假定大体成立；入世后各期系数总体落在零附近并略偏负，至 t+6 期约为 -0.27，呈现微弱下降迹象，"
        "但 95% 置信区间始终包含零。即便存在某种延迟的\u201c变绿\u201d效应，其统计证据也较脆弱。")
    add_image(doc, "fig3.png")
    add_caption(doc, "图 3  事件研究：WTO 入世对出口 CO\u2082 强度的动态影响")

    add_h2(doc, "（三）异质性分析")
    add_body(doc,
        "图 4 揭示了被平均效应掩盖的结构性差异。以发达国家为基准（效应约 0），高能源强度国家效应约 +15.7%、"
        "发展中国家约 +8.3%，方向上符合\u201c污染天堂\u201d假说（H1a、H2）；低能源强度国家约 -4.7%，"
        "方向上符合\u201c技术外溢\u201d假说（H1b）。这与理论预期一致：禀赋偏向碳密集产业的经济体更易承接高碳环节，"
        "而本已清洁的经济体更多受益于竞争与技术外溢。但各组置信区间普遍较宽且跨越零点，故只能作为方向性、提示性证据。")
    add_image(doc, "fig4.png")
    add_caption(doc, "图 4  WTO 效应的异质性分析")

    add_h2(doc, "（四）稳健性与安慰剂检验")
    add_body(doc,
        "表 3 显示，剔除最大出口国与排放国中国后系数为 0.113（p=0.279），对被解释变量 1% 缩尾后为 0.122（p=0.216），"
        "均与基准一致，表明结论不依赖个别样本或极端观测。图 5 的安慰剂检验将\u201c入世年份\u201d随机赋值并重复估计 500 次，"
        "所得系数分布大致以零为中心，而真实估计值 0.111 恰落入该分布主体区间，并非尾部小概率事件。"
        "这有力说明，基准正系数与随机噪声难以区分，WTO 成员身份对出口碳强度不存在稳健、显著的因果效应。")
    add_caption(doc, "表 3  稳健性检验")
    add_table(doc,
        ["检验方案", "β 系数", "效应 (%)", "p 值"],
        [
            ["基准模型（双向FE）", "0.1108", "11.72", "0.283"],
            ["排除中国", "0.1134", "12.01", "0.279"],
            ["去除极端值（1% 缩尾）", "0.1220", "12.98", "0.216"],
        ],
        col_align=[WD_ALIGN_PARAGRAPH.LEFT] + [WD_ALIGN_PARAGRAPH.CENTER]*3,
        body_size=10, header_size=10)
    add_image(doc, "fig5.png")
    add_caption(doc, "图 5  安慰剂检验：随机入世日期的系数分布")

    # ===================== 结论 =====================
    add_h1(doc, "六、结论与政策启示")
    add_body(doc,
        "本文利用 2000—2020 年跨国面板数据，系统检验了 WTO 成员身份对出口碳强度的影响，得到三点结论。"
        "第一，从全样本平均看，在控制国家与年份固定效应后，成员身份对出口碳强度并无统计显著影响，"
        "\u201c污染天堂\u201d与\u201c技术外溢\u201d两种机制大致相互抵消；混合 OLS 的显著负向关系实为选择性偏误所致。"
        "第二，样本期内出口碳强度的普遍下降，主要由收入增长（EKC）、能源强度改善与产业结构调整驱动，而非入世本身。"
        "第三，效应存在方向性异质性：高能源强度国家与发展中国家偏向\u201c污染天堂\u201d，低能源强度国家偏向\u201c技术外溢\u201d，但均不稳健。")
    add_body(doc,
        "上述发现具有清晰的政策含义。其一，不应简单地将 WTO 成员身份或贸易自由化视为碳强度上升的根源；"
        "贸易的环境效应是\u201c中性\u201d的，其方向取决于成员国内部的禀赋、发展阶段与配套政策。"
        "其二，能源密集型发展中国家应警惕高碳产业转入风险，通过提高能效标准、完善环境规制与碳定价，避免比较优势固化于高碳环节。"
        "其三，应发挥多边体制在技术扩散方面的正向作用，推动绿色技术与环境产品自由化，并审慎设计碳边境调节机制，"
        "使其既能防止碳泄漏，又不至于演变为变相的绿色保护主义。实现\u201c贸易—环境—技术\u201d政策协同，是推动出口部门持续低碳化的关键。")
    add_body(doc,
        "本文亦存局限：以国别加总数据衡量出口碳强度，难以剥离行业构成与全球价值链分工的影响；成员身份的二值处理也未刻画关税约束深度等强度差异。"
        "未来可结合行业与企业层面微观数据，并利用投入产出表测度出口\u201c隐含碳\u201d，进一步识别贸易自由化影响碳强度的具体渠道。")

    # ===================== 参考文献 =====================
    add_h1(doc, "参考文献")
    refs = [
        "[1] Grossman, G. M., & Krueger, A. B. Environmental Impacts of a North American Free Trade Agreement[R]. NBER Working Paper No. 3914, 1991.",
        "[2] Copeland, B. R., & Taylor, M. S. North-South Trade and the Environment[J]. The Quarterly Journal of Economics, 1994, 109(3): 755-787.",
        "[3] Copeland, B. R., & Taylor, M. S. Trade, Growth, and the Environment[J]. Journal of Economic Literature, 2004, 42(1): 7-71.",
        "[4] Antweiler, W., Copeland, B. R., & Taylor, M. S. Is Free Trade Good for the Environment?[J]. American Economic Review, 2001, 91(4): 877-908.",
        "[5] Frankel, J. A., & Rose, A. K. Is Trade Good or Bad for the Environment? Sorting Out the Causality[J]. The Review of Economics and Statistics, 2005, 87(1): 85-91.",
        "[6] Rose, A. K. Do We Really Know That the WTO Increases Trade?[J]. American Economic Review, 2004, 94(1): 98-114.",
        "[7] Levinson, A., & Taylor, M. S. Unmasking the Pollution Haven Effect[J]. International Economic Review, 2008, 49(1): 223-254.",
        "[8] World Bank. World Development Indicators[DB/OL]. Washington, D.C.: The World Bank, 2024.",
        "[9] 李小平, 卢现祥. 国际贸易、污染产业转移和中国工业 CO\u2082 排放[J]. 经济研究, 2010(1): 15-26.",
        "[10] 余淼杰, 张睿. 贸易自由化、技术进步与企业碳排放[J]. 经济学(季刊), 2017.",
    ]
    for r in refs:
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.line_spacing = 1.5; pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        pf.space_after = Pt(2); pf.left_indent = Pt(24); pf.first_line_indent = Pt(-24)
        run = p.add_run(r); set_run_font(run, size=10.5)

    add_page_number_footer(doc)
    out = "WTO期末论文_入世与出口碳强度.docx"
    doc.save(out)
    print("已生成：", out)


if __name__ == "__main__":
    missing = [f for f in ["fig1.png","fig2.png","fig3.png","fig4.png","fig5.png"] if not os.path.exists(f)]
    if missing:
        raise SystemExit("缺少图片，请先运行 generate_figures.py：" + ", ".join(missing))
    main()
