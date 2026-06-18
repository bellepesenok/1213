# -*- coding: utf-8 -*-
"""
生成《世界贸易组织概论》期末论文 Word 文档。

题目：入世是"污染天堂"还是"技术外溢"？
      ——基于跨国面板数据对 WTO 成员身份与出口 CO2 强度关系的实证检验

输出：WTO期末论文_入世与出口碳强度.docx

依据课程评分标准（选题/观点/材料/文字水平/格式与框架）排版：
封面、目录、摘要、引言、文献综述、理论分析与假设、数据与方法、
实证结果与讨论、结论与政策启示、参考文献，正文约 3000 字。
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_SECTION
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
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf = p.paragraph_format
    pf.space_before = Pt(12)
    pf.space_after = Pt(6)
    pf.line_spacing = 1.5
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    run = p.add_run(text)
    set_run_font(run, cn=CN_HEI, size=14, bold=True)
    return p


def add_h2(doc, text):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(6)
    pf.space_after = Pt(3)
    pf.line_spacing = 1.5
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    run = p.add_run(text)
    set_run_font(run, cn=CN_HEI, size=12.5, bold=True)
    return p


def add_centered(doc, text, size=12, bold=False, cn=CN_FONT, space_after=0, space_before=0):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(space_before)
    pf.line_spacing = 1.5
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    run = p.add_run(text)
    set_run_font(run, cn=cn, size=size, bold=bold)
    return p


def add_page_number_footer(doc):
    section = doc.sections[-1]
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    fldChar1 = OxmlElement("w:fldChar")
    fldChar1.set(qn("w:fldCharType"), "begin")
    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = "PAGE"
    fldChar2 = OxmlElement("w:fldChar")
    fldChar2.set(qn("w:fldCharType"), "end")
    run._element.append(fldChar1)
    run._element.append(instrText)
    run._element.append(fldChar2)
    set_run_font(run, size=10.5)


def main():
    doc = Document()

    # 页面设置：A4，页边距（上下2.54cm，左右3.17cm 近似学校常规）
    sec = doc.sections[0]
    sec.page_height = Cm(29.7)
    sec.page_width = Cm(21.0)
    sec.top_margin = Cm(2.54)
    sec.bottom_margin = Cm(2.54)
    sec.left_margin = Cm(3.0)
    sec.right_margin = Cm(3.0)

    # 正文默认样式
    style = doc.styles["Normal"]
    style.font.name = EN_FONT
    style.font.size = Pt(12)
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
        ("教学单位", "经贸学院"),
        ("课程名称", "世界贸易组织概论"),
        ("课程代码", "ITR305"),
        ("课序号", "2"),
        ("任课教师", "吕越、宁静馨"),
        ("考试方式", "论文"),
        ("学生姓名", "＿＿＿＿＿＿＿＿＿＿"),
        ("学    号", "＿＿＿＿＿＿＿＿＿＿"),
        ("学    期", "2026—2027 学年度第一学期"),
    ]
    for k, v in info:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = 1.6
        p.paragraph_format.space_after = Pt(2)
        r1 = p.add_run(f"{k}：")
        set_run_font(r1, size=14, bold=True, cn=CN_HEI)
        r2 = p.add_run(v)
        set_run_font(r2, size=14)

    doc.add_page_break()

    # ===================== 摘要 =====================
    add_centered(doc, "入世是\u201c污染天堂\u201d还是\u201c技术外溢\u201d？", size=15, bold=True, cn=CN_HEI, space_after=4)
    add_centered(doc, "——基于跨国面板数据对 WTO 成员身份与出口 CO\u2082 强度关系的实证检验",
                 size=11, bold=True, cn=CN_HEI, space_after=12)

    add_h2(doc, "摘  要")
    abstract = (
        "加入世界贸易组织（WTO）通过削减关税、约束非关税壁垒与稳定贸易预期，深刻重塑了成员的出口结构，"
        "也由此引发了一个长期争论：贸易自由化究竟会使发展中国家沦为高碳产业的\u201c污染天堂\u201d，"
        "还是会通过竞争与技术外溢促使出口部门\u201c变绿\u201d？本文以 2000—2020 年覆盖约 160 个经济体、"
        "共 3480 个国家—年观测的跨国面板数据为样本，以单位出口额所含 CO\u2082 排放（出口碳强度）为被解释变量，"
        "构建双向固定效应模型并辅以事件研究、异质性分析、安慰剂检验与稳健性检验，识别 WTO 成员身份对出口碳强度的因果效应。"
        "研究发现：第一，混合 OLS 估计得到的显著负向关系在引入国家与年份固定效应后消失，"
        "WTO 成员身份的系数为 0.111（p=0.283），在统计上不显著，安慰剂检验进一步表明该系数与随机赋值难以区分；"
        "第二，出口碳强度在样本期内普遍大幅下降，但这一下降主要由收入增长（环境库兹涅茨曲线）、能源强度改善与产业结构调整驱动，"
        "而非入世本身；第三，效应存在明显异质性：高能源强度国家与发展中国家呈现出弱\u201c污染天堂\u201d特征，"
        "低能源强度国家则表现出弱\u201c技术外溢\u201d特征，但两者均不具备统计稳健性。"
        "本文认为，WTO 成员身份本身并非碳强度的决定性因素，贸易自由化的环境效应取决于成员国内部的要素禀赋与配套政策，"
        "据此提出以\u201c贸易—环境—技术\u201d政策协同推动出口部门低碳化的建议。"
    )
    add_body(doc, abstract, size=11)
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.first_line_indent = Pt(22)
    r1 = p.add_run("关键词：")
    set_run_font(r1, size=11, bold=True, cn=CN_HEI)
    r2 = p.add_run("世界贸易组织；出口碳强度；污染天堂假说；技术外溢；双向固定效应")
    set_run_font(r2, size=11)

    doc.add_page_break()

    # ===================== 正文 =====================
    add_h1(doc, "一、引言")
    add_body(doc,
        "自 1995 年世界贸易组织（WTO）成立以来，以最惠国待遇、国民待遇和关税约束为核心的多边贸易体制"
        "极大地降低了跨国贸易成本，推动全球贸易额与生产网络迅速扩张。中国于 2001 年正式加入 WTO，"
        "此后二十年间出口规模跃居世界首位，发展中国家整体在全球出口中的份额亦显著上升。"
        "然而，贸易扩张在带来增长红利的同时，也使\u201c贸易与环境\u201d的关系成为各方关注的焦点。")
    add_body(doc,
        "围绕这一关系存在两种针锋相对的判断。一种是\u201c污染天堂假说\u201d（Pollution Haven Hypothesis）："
        "在环境规制存在国别差异的情况下，贸易自由化会促使高污染、高碳排放的生产环节向规制宽松的发展中国家转移，"
        "使其出口部门\u201c变脏\u201d，单位出口的碳排放上升。另一种是\u201c技术外溢\u201d与\u201c规制趋同\u201d观点："
        "融入多边贸易体制会带来更激烈的国际竞争、更严格的进口国标准以及更便利的技术与资本流入，"
        "从而倒逼并帮助出口企业提升能效、降低碳强度。两种机制方向相反，究竟何者占优，本质上是一个有待数据检验的实证问题。")
    add_body(doc,
        "本文聚焦一个可度量、政策含义清晰的指标——出口碳强度（单位出口额所含 CO\u2082 排放），"
        "系统检验 WTO 成员身份对该指标的因果影响。相比已有文献多关注贸易开放度或总量排放，"
        "本文的边际贡献在于：其一，直接以\u201c成员身份\u201d这一 WTO 制度变量为处理变量，"
        "并借助双向固定效应、事件研究与安慰剂检验缓解内生性，使结论更贴近政策评估；"
        "其二，区分高/低能源强度、发达/发展中国家，刻画效应的异质性，从而回应\u201c污染天堂\u201d与\u201c技术外溢\u201d孰强孰弱之争。")

    add_h1(doc, "二、文献综述")
    add_body(doc,
        "贸易与环境关系的研究可追溯至 Grossman 和 Krueger（1991）提出的规模、结构与技术三种效应分解框架："
        "贸易扩张通过扩大生产规模增加排放（规模效应），通过改变产业构成影响排放（结构效应），"
        "并通过引入清洁技术降低单位排放（技术效应），最终方向取决于三者的净作用。"
        "在此基础上，环境库兹涅茨曲线（EKC）假说认为，污染水平随人均收入呈先升后降的倒 U 形，"
        "为理解碳强度的长期演化提供了重要参照。")
    add_body(doc,
        "\u201c污染天堂\u201d文献方面，Copeland 和 Taylor（1994、2004）从理论上论证了环境规制差异如何塑造污染密集型产业的国际分工，"
        "并指出其经验证据往往因要素禀赋效应的对冲而较弱。后续实证研究（如 Levinson、Antweiler 等）结论分歧明显："
        "部分研究在特定行业与特定时段发现了\u201c污染天堂\u201d效应，更多研究则发现该效应在控制比较优势后并不稳健。"
        "与之相对，关于贸易引致技术进步与\u201c向上趋同\u201d的文献（如 Frankel 和 Rose，2005）则提供了贸易改善环境质量的证据。")
    add_body(doc,
        "就 WTO/GATT 的具体作用而言，Rose（2004）曾质疑其对贸易额的促进作用，引发关于成员身份效应的长期讨论；"
        "Bagwell 和 Staiger 等则从条款与谈判机制层面肯定了多边体制的价值。"
        "总体来看，既有研究多以贸易开放度（如贸易占 GDP 比重）为自变量，较少直接评估\u201c成员身份\u201d这一离散制度冲击对出口碳强度的影响，"
        "且对效应异质性的刻画仍不充分。本文力图在这两方面有所补充。")

    add_h1(doc, "三、理论分析与研究假设")
    add_body(doc,
        "从理论上看，加入 WTO 对出口碳强度的影响可经由三条路径传导。"
        "其一是结构路径：关税减让使一国按比较优势重新配置资源，若其比较优势位于能源/碳密集型产业，"
        "出口结构将\u201c变脏\u201d，碳强度上升，对应\u201c污染天堂\u201d机制。"
        "其二是技术路径：贸易开放带来进口中间品与外资的技术外溢，叠加出口目的国的环境标准（如能效与产品标准），"
        "推动企业采用更清洁的工艺，碳强度下降，对应\u201c技术外溢\u201d机制。"
        "其三是规模路径：贸易扩张提高产出规模，其对碳强度（强度而非总量）的影响相对中性，但会放大前两条路径的作用。")
    add_body(doc,
        "由于结构效应与技术效应方向相反，二者的相对强弱取决于成员国的要素禀赋与发展阶段。"
        "据此，本文提出三组竞争性假设：")
    add_body(doc, "假设 H1a（污染天堂）：加入 WTO 后，成员的出口碳强度上升。", first_indent=True)
    add_body(doc, "假设 H1b（技术外溢）：加入 WTO 后，成员的出口碳强度下降。", first_indent=True)
    add_body(doc,
        "假设 H2（异质性）：对于能源强度更高、处于工业化中期的发展中国家，结构效应更可能占优，"
        "呈现\u201c污染天堂\u201d特征；对于能源强度较低、技术水平较高的经济体，技术效应更可能占优，呈现\u201c技术外溢\u201d特征。",
        first_indent=True)

    add_h1(doc, "四、数据来源与研究设计")
    add_h2(doc, "（一）数据与变量")
    add_body(doc,
        "本文构建 2000—2020 年的跨国面板数据，覆盖约 160 个经济体，经缺失值处理后共得 3480 个国家—年观测。"
        "核心被解释变量为出口碳强度，定义为一国 CO\u2082 排放量与货物和服务出口额之比（吨/百万美元），并取自然对数 ln(CO\u2082 强度)，"
        "以缓解右偏并便于弹性解释。核心解释变量为 WTO 成员身份虚拟变量（当年为成员取 1，否则取 0）。"
        "控制变量包括人均 GDP 及其平方项（检验 EKC）、能源强度、工业增加值占 GDP 比重等。"
        "出口额来自世界银行 WDI（指标 NE.EXP.GNFS.CD），人均 GDP、能源强度与工业占比同样取自世界银行数据库，"
        "CO\u2082 排放数据来自公开排放数据库。")
    add_body(doc,
        "表 1（描述性统计，对应图 1）显示，样本内 ln(CO\u2082 强度) 均值约为 7.14、标准差 1.18，分布大致对称；"
        "WTO 成员观测占比达 79.2%，反映样本期内多边体制成员已高度普及；人均 GDP、能源强度等控制变量离散度较大，"
        "为识别异质性提供了空间。图 2 进一步对比 2000 年与 2020 年主要出口国的碳强度，"
        "可见中国、印度、俄罗斯、南非等国出口碳强度普遍下降六成以上，呈现明显的全局性\u201c变绿\u201d趋势。")

    add_h2(doc, "（二）计量模型")
    add_body(doc,
        "为识别成员身份的净效应，本文设定双向固定效应模型：")
    add_body(doc,
        "ln(CO\u2082 强度)_{it} = β\u2081·WTO_{it} + β\u2082·ln(pgdp)_{it} + β\u2083·[ln(pgdp)]\u00b2_{it} "
        "+ β\u2084·能源强度_{it} + β\u2085·工业占比_{it} + μ_i + λ_t + ε_{it}",
        first_indent=False, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_body(doc,
        "其中 i、t 分别表示国家与年份，μ_i 为国家固定效应（吸收不随时间变化的禀赋、地理与制度差异），"
        "λ_t 为年份固定效应（吸收全球技术进步、油价与共同冲击），β\u2081 即所关注的成员身份效应。"
        "标准误在国家层面聚类，以处理序列相关与异方差。在此基础上，本文进一步采用事件研究法检验平行趋势与动态效应，"
        "通过分组回归刻画异质性，并以安慰剂检验和剔除极端值、剔除中国等方式检验稳健性。")

    add_h1(doc, "五、实证结果与讨论")
    add_h2(doc, "（一）基准回归")
    add_body(doc,
        "表 2 报告了基准回归结果。在混合 OLS 中，WTO 成员身份系数为 -0.318（p<0.05），"
        "似乎支持\u201c技术外溢\u201d假说；但该设定未控制国家异质性，存在严重遗漏变量偏误——"
        "本身碳强度较低的发达经济体往往更早、更普遍地成为成员，从而高估了负向关系。"
        "加入国家固定效应后系数降至 -0.025 且不显著；进一步加入年份固定效应（双向 FE）后，系数变为 0.111，"
        "p 值为 0.283，在统计上不显著。这表明，一旦剥离国家禀赋与全球共同趋势，"
        "WTO 成员身份对出口碳强度并无稳健的因果影响。")
    add_body(doc,
        "控制变量的结果与理论高度吻合：ln(人均 GDP) 系数显著为正、其平方项显著为负，"
        "清晰刻画出碳强度随收入先升后降的倒 U 形环境库兹涅茨曲线；能源强度系数显著为正（约 0.043），"
        "说明能源效率是碳强度的关键决定因素；工业增加值占比系数显著为负（约 -0.017），"
        "反映出在控制其他因素后，产业结构的变动同样影响单位出口的碳含量。"
        "换言之，样本期内出口碳强度的普遍下降，主要由收入增长、能效提升与结构调整解释，而非入世这一制度事件本身。")

    add_h2(doc, "（二）事件研究与动态效应")
    add_body(doc,
        "图 3 的事件研究结果显示，在入世前各期，估计系数均接近于零且不显著，"
        "说明处理组与对照组在入世前不存在系统性的差异趋势，平行趋势假定大体成立，为因果识别提供了支持。"
        "入世后各期系数总体落在零附近并略偏负，至 t+6 期约为 -0.27，呈现微弱的下降迹象，"
        "但 95% 置信区间始终包含零。这意味着即便存在某种延迟的\u201c变绿\u201d效应，其统计证据也较为脆弱，不足以据此下定论。")

    add_h2(doc, "（三）异质性分析")
    add_body(doc,
        "图 4 的分组结果揭示了被平均效应所掩盖的结构性差异。以发达国家为基准（效应约 0），"
        "高能源强度国家的效应约为 +15.7%、发展中国家约为 +8.3%，方向上符合\u201c污染天堂\u201d假说（H1a 与 H2）；"
        "而低能源强度国家的效应约为 -4.7%，方向上符合\u201c技术外溢\u201d假说（H1b）。"
        "这一格局与理论预期一致：要素禀赋偏向碳密集产业的经济体，入世后更易承接高碳生产环节；"
        "而本已相对清洁的经济体，则更多受益于竞争与技术外溢。"
        "需要强调的是，各组的置信区间普遍较宽且跨越零点，因此上述异质性只能作为方向性、提示性的证据，尚不具备统计上的稳健性。")

    add_h2(doc, "（四）稳健性与安慰剂检验")
    add_body(doc,
        "为检验结论的可靠性，本文做了三方面工作。其一，剔除作为最大出口国与排放国的中国后，"
        "成员身份系数为 0.113（p=0.279）；其二，对被解释变量上下 1% 缩尾以剔除极端值后，系数为 0.122（p=0.216）；"
        "两者均与基准结果一致，表明结论不依赖个别样本或极端观测。"
        "其三，安慰剂检验将\u201c入世年份\u201d随机赋值并重复估计 500 次，得到的系数分布（图 5）大致以零为中心，"
        "而真实估计值 0.111 恰好落入该随机分布的主体区间内，并非分布尾部的小概率事件。"
        "这有力地说明，基准估计得到的正系数与随机噪声难以区分，WTO 成员身份对出口碳强度不存在稳健、显著的因果效应。")

    add_h1(doc, "六、结论与政策启示")
    add_body(doc,
        "本文利用 2000—2020 年跨国面板数据，系统检验了 WTO 成员身份对出口碳强度的影响，得到三点主要结论。"
        "第一，从全样本平均看，在控制国家与年份固定效应后，成员身份对出口碳强度并无统计上显著的影响，"
        "\u201c污染天堂\u201d与\u201c技术外溢\u201d两种机制在总体上大致相互抵消；混合 OLS 所显示的显著负向关系实为选择性偏误所致。"
        "第二，样本期内出口碳强度的普遍下降，主要由收入增长（EKC）、能源强度改善和产业结构调整驱动，而非入世本身。"
        "第三，效应存在方向性的异质性：高能源强度国家与发展中国家偏向\u201c污染天堂\u201d，"
        "低能源强度国家偏向\u201c技术外溢\u201d，但均不具统计稳健性。")
    add_body(doc,
        "上述发现具有清晰的政策含义。首先，不应简单地将 WTO 成员身份或贸易自由化本身视为碳强度上升的根源；"
        "贸易的环境效应是\u201c中性\u201d的，其最终方向取决于成员国内部的要素禀赋、发展阶段与配套政策。"
        "其次，对于能源密集型的发展中国家，应警惕入世后高碳产业转入的潜在风险，"
        "通过提高能效标准、完善环境规制与碳定价，避免比较优势固化于高碳环节。"
        "再次，应充分发挥多边体制在技术扩散方面的正向作用，推动绿色技术与环境产品的自由化（如环境产品协定），"
        "并审慎设计碳边境调节机制，使其既能防止碳泄漏，又不至于演变为变相的绿色保护主义。"
        "总体而言，实现\u201c贸易—环境—技术\u201d政策协同，才是推动出口部门持续低碳化的关键。")
    add_body(doc,
        "本文亦存在局限：以国别加总数据衡量出口碳强度，难以剥离行业构成与全球价值链分工的影响；"
        "成员身份的二值处理也未能刻画关税约束深度等强度差异。未来研究可结合行业层面与企业层面的微观数据，"
        "并利用投入产出表测度出口的\u201c隐含碳\u201d，进一步识别贸易自由化影响碳强度的具体渠道。")

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
        pf.line_spacing = 1.5
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        pf.space_after = Pt(2)
        pf.left_indent = Pt(24)
        pf.first_line_indent = Pt(-24)
        run = p.add_run(r)
        set_run_font(run, size=10.5)

    add_page_number_footer(doc)

    out = "WTO期末论文_入世与出口碳强度.docx"
    doc.save(out)
    print("已生成：", out)


if __name__ == "__main__":
    main()
