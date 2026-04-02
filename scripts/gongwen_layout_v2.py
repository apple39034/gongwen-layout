"""
公文格式排版脚本 v2 — 修正元素顺序，直接操作 ZIP
"""
import sys, os, re, unicodedata, shutil, zipfile
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
def qn(tag): return f"{{{W}}}{tag}"

# ─── 公文格式常量 ───────────────────────────────────────────────────────────
PAGE_W      = 11906
PAGE_H      = 16838
# 边距: 36mm 上下, 27mm 左右  (pgMar 单位 = twips = 1/1440 inch; 1mm = 1440/25.4 twips)
MARGIN_VERT = round(36 * 1440 / 25.4)   # 2041
MARGIN_SIDE = round(27 * 1440 / 25.4)   # 1531
HEADER_H    = round(15 * 1440 / 25.4)   # 850
FOOTER_H    = round(27 * 1440 / 25.4)   # 1531

BODY_LINE   = round(28.95 * 20)        # 579 twips
FIRST_IND   = round(2 * 16 * 20)       # 首行缩进2字符 = 640 twips
LEFT_IND    = round(2 * 16 * 20)       # 左缩进2字符

FONT_XIAO  = "方正小标宋简体"
FONT_FANG  = "仿宋"
FONT_HEI   = "黑体"
FONT_KAI   = "楷体"
FONT_SONG  = "宋体"

# ─── OOXML schema: <w:pPr> 子元素顺序 ───────────────────────────────────────
# 0: pStyle, 1: keepNext, 1: keepLines, 1: pageBreakBefore, 2: framePr,
# 2: widowControl, 2: numPr, 3: suppressLineNumbers, 3: pBdr,
# 4: shd, 4: tabs, 4: suppressAutoHyphens, 4: kinsoku, 4: wordWrap,
# 4: overflowPunct, 4: topLinePunct, 4: autoSpaceDE, 4: autoSpaceDN,
# 4: bidi, 4: adjustRightInd, 4: snapToGrid,
# 5: spacing, 6: ind, 7: contextualSpacing, 7: mirrorIndents, 7: suppressOverlap,
# 7: jc, 8: textDirection, 8: textAlignment, 8: textboxTightWrap,
# 8: outlineLvl, 9: divId, 9: cnfStyle, 10: rPr, 11: sectPr, 11: pPrChange
# 顺序: pStyle→numPr→spacing→ind→jc→outlineLvl→sectPr→...

# ─── XML 构建辅助 ───────────────────────────────────────────────────────────

def e(tag, **attrs):
    el = etree.Element(qn(tag))
    for k, v in attrs.items():
        el.set(qn(k), str(v))
    return el

def txt_run(text, font, pt, bold=False, italic=False):
    r = etree.Element(qn("r"))
    rPr = etree.SubElement(r, qn("rPr"))
    rf = etree.SubElement(rPr, qn("rFonts"))
    for k in ("ascii","hAnsi","eastAsia","cs"):
        rf.set(qn(k), font)
    sz  = etree.SubElement(rPr, qn("sz"));   sz.set(qn("val"),  str(int(pt)*2))
    szc = etree.SubElement(rPr, qn("szCs")); szc.set(qn("val"), str(int(pt)*2))
    if bold:   etree.SubElement(rPr, qn("b"));    etree.SubElement(rPr, qn("bCs"))
    if italic: etree.SubElement(rPr, qn("i"));    etree.SubElement(rPr, qn("iCs"))
    t = etree.SubElement(r, qn("t"))
    t.set("{http://www.w3.org/XML/1998/namespace}space","preserve")
    t.text = text
    return r


def p_with_runs(children, align=None, spacing_line=None, line_rule=None,
                first_indent=None, left_indent=None, right_indent=None,
                before=None, after=None, outline_lvl=None, pStyle=None):
    """按 schema 顺序构建 <w:p>"""
    p = etree.Element(qn("p"))
    pPr = etree.SubElement(p, qn("pPr"))
    # 顺序: pStyle → spacing → ind → jc → outlineLvl
    if pStyle is not None:
        pPr.append(e("pStyle", val=pStyle))
    # spacing
    sp_kwargs = {}
    if spacing_line is not None: sp_kwargs["line"] = str(spacing_line)
    if line_rule:                sp_kwargs["lineRule"] = line_rule
    if before is not None:       sp_kwargs["before"] = str(before)
    if after is not None:        sp_kwargs["after"]  = str(after)
    if sp_kwargs:
        pPr.append(e("spacing", **sp_kwargs))
    # ind
    ind_kwargs = {}
    if first_indent is not None: ind_kwargs["firstLine"] = str(first_indent)
    if left_indent  is not None: ind_kwargs["left"]     = str(left_indent)
    if right_indent is not None: ind_kwargs["right"]    = str(right_indent)
    if ind_kwargs:
        pPr.append(e("ind", **ind_kwargs))
    # jc
    if align:
        pPr.append(e("jc", val=align))
    # outlineLvl
    if outline_lvl is not None:
        pPr.append(e("outlineLvl", val=str(outline_lvl)))
    # children
    for ch in (children or []):
        p.append(ch)
    return p


def spacer():
    """空行段落（28.95pt间距）"""
    return p_with_runs([], spacing_line=BODY_LINE, line_rule="exact",
                       before=BODY_LINE, after=0)


def title_para(text):
    """标题: 方正小标宋，二号，居中，粗体"""
    return p_with_runs(
        [txt_run(text, FONT_XIAO, 22, bold=True)],
        align="center",
        spacing_line=BODY_LINE, line_rule="exact",
        before=round(22*2*20), after=round(22*1*20),
    )


def body_para(text):
    """正文: 仿宋三号，首行缩进2字符，两端对齐"""
    return p_with_runs(
        [txt_run(text, FONT_FANG, 16)],
        align="both",
        spacing_line=BODY_LINE, line_rule="exact",
        first_indent=FIRST_IND,
    )


def h1_para(text):
    """一级标题: 黑体三号，左缩进2字符，粗体，段前段后0磅"""
    return p_with_runs(
        [txt_run(text, FONT_HEI, 16, bold=True)],
        spacing_line=BODY_LINE, line_rule="exact",
        before=0, after=0,
        left_indent=LEFT_IND,
        outline_lvl=0,
    )


def h2_para(text):
    """二级标题: 楷体三号，左缩进2字符，粗体"""
    return p_with_runs(
        [txt_run(text, FONT_KAI, 16, bold=True)],
        spacing_line=BODY_LINE, line_rule="exact",
        before=round(16*1*20), after=round(16*0.5*20),
        left_indent=LEFT_IND,
        outline_lvl=1,
    )


def ref_para(text):
    """参考文献条目: 仿宋三号，首行缩进2字符，两端对齐"""
    return p_with_runs(
        [txt_run(text, FONT_FANG, 16)],
        align="both",
        spacing_line=BODY_LINE, line_rule="exact",
        first_indent=FIRST_IND,
    )


def right_para(text, right_chars=4):
    """落款: 右对齐，右缩进 N 个字符"""
    right_dxa = round(right_chars * 16 * 20)
    return p_with_runs(
        [txt_run(text, FONT_FANG, 16)],
        align="right",
        spacing_line=BODY_LINE, line_rule="exact",
        right_indent=right_dxa,
    )


def sect_pr_xml():
    """返回 sectPr 元素（手工写XML字符串，避免属性前缀问题）"""
    xml_str = (
        f'<w:sectPr xmlns:w="{W}">'
        f'<w:pgSz w:w="{PAGE_W}" w:h="{PAGE_H}"/>'
        f'<w:pgMar w:top="{MARGIN_VERT}" w:right="{MARGIN_SIDE}"'
        f' w:bottom="{MARGIN_VERT}" w:left="{MARGIN_SIDE}"'
        f' w:header="{HEADER_H}" w:footer="{FOOTER_H}" w:gutter="0"/>'
        f'<w:cols w:space="708"/>'
        f'<w:docGrid w:linePitch="360"/>'
        f'</w:sectPr>'
    )
    return etree.fromstring(xml_str)


# ─── 主处理 ──────────────────────────────────────────────────────────────────

def process(src_docx, out_docx):
    tmp_dir = "/tmp/_gongwen_work"
    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir)

    # 1. 解压原始 docx
    with zipfile.ZipFile(src_docx) as zf:
        zf.extractall(tmp_dir)

    # 2. 读取原始 document.xml 了解 namespace
    doc_path = os.path.join(tmp_dir, "word", "document.xml")
    tree = etree.parse(doc_path)
    root = tree.getroot()
    body = root.find(qn("body"))

    # 3. 清空 body 内容（保留 sectPr 后面追加）
    for child in list(body):
        body.remove(child)

    # 4. 构建公文内容
    # 标题
    body.append(title_para("社区精神分裂症患者长效针剂与口服药对家庭负担的影响"))
    body.append(spacer())

    # 执行摘要
    body.append(h1_para("执行摘要"))
    body.append(body_para(
        "最新研究表明，长效针剂（LAI）抗精神病药在多个方面优于口服药，"
        "可显著减轻家庭负担。主要优势包括：提高依从性、降低复发率、减少再住院，"
        "改善患者结局，从而间接减轻照料者负担。"
    ))

    # 一、依从性
    body.append(h1_para("一、依从性"))
    body.append(body_para("LAI 可显著改善治疗依从性，减少因漏服导致的复发。"))
    body.append(body_para("依从性差是精神分裂症复发和再入院的主要原因。"))

    # 二、复发与再住院
    body.append(h1_para("二、复发与再住院"))
    body.append(body_para("多项研究证实 LAI 在预防复发和再住院方面优于口服药。"))
    body.append(body_para("镜像研究显示：LAI 治疗可显著降低 3 年随访期间的复发率和住院率。"))

    # 三、照料者负担
    body.append(h1_para("三、照料者负担"))
    body.append(body_para("关键发现：从口服药转换为 LAI 后，照料者负担显著改善。"))
    body.append(body_para("患者病情稳定，家庭成员花在监护上的时间减少。"))
    body.append(body_para("经济负担减轻：减少急诊和住院费用。"))

    # 四、经济负担
    body.append(h1_para("四、经济负担"))
    body.append(body_para("直接医疗成本：减少（住院减少，成本降低）。"))
    body.append(body_para("间接成本：减少（误工、护理时间减少）。"))
    body.append(body_para("总体负担：LAI 可减轻精神分裂症的经济负担。"))

    # 五、最新进展
    body.append(h1_para("五、2024—2025 年最新进展"))
    body.append(body_para("阿立哌唑微球（长效针剂）在中国获批上市，被业内专家称为重生之路。"))
    body.append(body_para("更多研究关注 LAI 在首发和早期精神分裂症患者中的应用。"))
    body.append(body_para("指南建议：多发作、依从性差的患者优先考虑 LAI。"))

    # 六、关键结论
    body.append(h1_para("六、关键结论"))
    for i, c in enumerate([
        "LAI 在依从性、复发预防、再住院方面优于口服药。",
        "LAI 可间接减轻家庭照料负担和经济负担。",
        "2024—2025 年 LAI 在国内的可及性进一步提高。",
        "需要根据患者个体情况选择合适方案。",
    ], 1):
        body.append(ref_para(f"{i}. {c}"))

    # 七、参考文献
    body.append(h1_para("七、参考文献"))
    for r in [
        "JMCP 2025 – Budget impact of AOM.",
        "Sage 2025 – Use of LAI in acute inpatient.",
        "Wiley 2024 – Assessing impact of LAI.",
        "PMC 2026 – 3-year follow-up mirror-image study.",
        "Nature 2017 – Caregiver burden improved with LAI.",
        "PubMed 2023 – Relationship of LAI with Caregiver Burden.",
        "JMCP 2024 – Adherence, costs comparison.",
        "ScienceNet 2025 – 长效针剂有望开启重生之路。",
    ]:
        body.append(ref_para(r))

    # 空三行 + 落款
    for _ in range(3):
        body.append(spacer())
    body.append(right_para("社区精神卫生服务中心", right_chars=4))
    body.append(right_para("2026 年 4 月",        right_chars=4))

    # sectPr
    body.append(sect_pr_xml())

    # 5. 写回 document.xml
    tree.write(doc_path, xml_declaration=True, encoding="UTF-8", standalone=True)
    print(f"document.xml 已写入: {doc_path}")

    # 6. 清理无引用文件及其关系条目
    for fname in ["comments.xml", "footnotes.xml"]:
        fpath = os.path.join(tmp_dir, "word", fname)
        if os.path.exists(fpath):
            os.remove(fpath)
            print(f"已删除: {fname}")
        # 从 rels 移除引用
        rels_path = os.path.join(tmp_dir, "word", "_rels", "document.xml.rels")
        if os.path.exists(rels_path):
            with open(rels_path) as f:
                content = f.read()
            # 移除对该文件的 Relationship 元素
            content = re.sub(
                rf'<Relationship[^>]+Target=".{re.escape(fname)}."[^>]*/>', '', content)
            with open(rels_path, "w") as f:
                f.write(content)
        # 从 [Content_Types].xml 移除 Override
        ct_path = os.path.join(tmp_dir, "[Content_Types].xml")
        if os.path.exists(ct_path):
            with open(ct_path) as f:
                content = f.read()
            content = re.sub(
                rf'<Override[^>]+PartName="/word/{re.escape(fname)}"[^>]*/>', '', content)
            with open(ct_path, "w") as f:
                f.write(content)

    # 7. 打包成 docx
    with zipfile.ZipFile(out_docx, "w", zipfile.ZIP_DEFLATED) as zout:
        for dirpath, dirnames, filenames in os.walk(tmp_dir):
            for filename in filenames:
                filepath = os.path.join(dirpath, filename)
                arcname = os.path.relpath(filepath, tmp_dir)
                zout.write(filepath, arcname)
    print(f"已保存到: {out_docx}")
    shutil.rmtree(tmp_dir)


if __name__ == "__main__":
    src = sys.argv[1]
    dst = sys.argv[2] if len(sys.argv) > 2 else src
    process(src, dst)
