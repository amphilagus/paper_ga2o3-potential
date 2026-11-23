from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path("/Users/amphilagusgu/手稿/Ga2O3 potential")
OUTPUT = ROOT / "氧化镓神经演化势科研进展通稿_清洁版.docx"
FIGURE_1 = ROOT / "ga2o3-nep-training-figure-v9/figure/nep-training-principle-v9.png"
FIGURE_2 = ROOT / "figures/track_diameter_comparison.png"

TITLE = "近代物理所在氧化镓神经演化势及快重离子辐照研究中获进展"
DOI_URL = "https://doi.org/10.1063/5.0323486"

BODY_PARAGRAPHS = [
    (
        "近日，中国科学院近代物理研究所科研团队构建了适用于氧化镓多晶型与辐照过程的"
        "神经演化势（NEP），并通过百万原子级分子动力学"
        "模拟揭示快重离子辐照形成的核—壳离子径迹结构，为理解不同实验中观察到的非晶化"
        "与相变现象提供了统一图景。相关成果发表于《应用物理评论》（Applied Physics "
        "Reviews）。"
    ),
    (
        "氧化镓是新一代超宽禁带半导体材料，在高功率、高频电子器件及极端环境应用中"
        "具有重要潜力。该材料存在多种晶体结构，辐照过程中会经历局域熔化、无序化和"
        "相变等强非平衡过程。第一性原理计算难以覆盖百万原子、百皮秒尺度，传统经验势"
        "和已有机器学习势又难以同时兼顾复杂结构描述与计算效率。"
    ),
    (
        "研究团队采用面向物理过程的数据构建策略，在已有数据基础上补充γ相能量—体积"
        "构型和β相升降温构型，并引入能量依赖加权，使模型重点学习稳定结构，同时保留"
        "对高能非平衡构型的描述能力。测试中，该模型的能量平均绝对误差为每原子16.4"
        "毫电子伏，约为对比模型tabGAP的三分之一；依托图形处理器分子动力学程序GPUMD，"
        "模型可高效开展百万原子级模拟。"
    ),
    (
        "基于该模型，团队模拟了约146万个原子持续200皮秒的辐照响应。结果显示，离子径迹"
        "沿径向依次呈现非晶核心、γ相壳层、变形β相和完整β相四个区域；径迹直径在电子"
        "能损超过约17千电子伏每纳米后逐渐增大，并与已有透射电子显微镜实验结果吻合。"
        "中等电子能损下，核心区保留向γ相转变的前驱结构；能损继续提高后，非晶核心和"
        "γ相壳层趋于稳定。"
    ),
    (
        "研究表明，此前实验报道的非晶径迹与γ相等不同结构并非彼此矛盾，而可能对应不同"
        "电子能损、径迹位置及辐照后弛豫或表征条件。该工作为复杂多晶型材料机器学习势"
        "的构建提供了“过程导向采样—能量依赖加权”思路，也为理解氧化镓器件在辐照环境"
        "中的结构演化奠定了原子尺度基础。研究工作得到国家自然科学基金、中国科学院"
        "人才计划青年项目、惠州市科技人才项目和甘肃省科技创新人才计划等支持。"
    ),
]


def set_run_font(run, latin="Arial Unicode MS", east_asia="Arial Unicode MS", size=11, bold=False, color="000000"):
    run.font.name = latin
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor.from_string(color)
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), east_asia)
    run._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), latin)
    run._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), latin)


def set_paragraph_language(paragraph, language="zh-CN"):
    for run in paragraph.runs:
        r_pr = run._element.get_or_add_rPr()
        lang = r_pr.find(qn("w:lang"))
        if lang is None:
            lang = OxmlElement("w:lang")
            r_pr.append(lang)
        lang.set(qn("w:val"), language)
        lang.set(qn("w:eastAsia"), language)


def add_bottom_border(paragraph, color="C9D2DE", size=6, space=5):
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = p_pr.find(qn("w:pBdr"))
    if p_bdr is None:
        p_bdr = OxmlElement("w:pBdr")
        p_pr.append(p_bdr)
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size))
    bottom.set(qn("w:space"), str(space))
    bottom.set(qn("w:color"), color)
    p_bdr.append(bottom)


def add_hyperlink(paragraph, text, url):
    part = paragraph.part
    relationship_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), relationship_id)
    new_run = OxmlElement("w:r")
    run_properties = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "315F86")
    run_properties.append(color)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    run_properties.append(underline)
    fonts = OxmlElement("w:rFonts")
    fonts.set(qn("w:ascii"), "Arial Unicode MS")
    fonts.set(qn("w:hAnsi"), "Arial Unicode MS")
    fonts.set(qn("w:eastAsia"), "Arial Unicode MS")
    run_properties.append(fonts)
    size = OxmlElement("w:sz")
    size.set(qn("w:val"), "20")
    run_properties.append(size)
    new_run.append(run_properties)
    text_element = OxmlElement("w:t")
    text_element.text = text
    new_run.append(text_element)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    set_run_font(run, size=9, color="6B7280")
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    placeholder = OxmlElement("w:t")
    placeholder.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instruction, separate, placeholder, end])


def configure_styles(document):
    styles = document.styles
    normal = styles["Normal"]
    normal.font.name = "Arial Unicode MS"
    normal.font.size = Pt(10.5)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial Unicode MS")
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    normal.paragraph_format.line_spacing = 1.28
    normal.paragraph_format.space_after = Pt(6)

    title_style = styles["Title"]
    title_style.font.name = "Arial Unicode MS"
    title_style.font.size = Pt(17)
    title_style.font.bold = True
    title_style.font.color.rgb = RGBColor.from_string("000000")
    title_style._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial Unicode MS")
    title_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_style.paragraph_format.space_before = Pt(6)
    title_style.paragraph_format.space_after = Pt(12)
    title_p_pr = title_style._element.get_or_add_pPr()
    title_border = title_p_pr.find(qn("w:pBdr"))
    if title_border is not None:
        title_p_pr.remove(title_border)

    for name in ("Heading 1", "Heading 2"):
        style = styles[name]
        style.font.name = "Arial Unicode MS"
        style.font.bold = True
        style.font.color.rgb = RGBColor.from_string("000000")
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial Unicode MS")

    if "Figure Caption CN" not in styles:
        caption_style = styles.add_style("Figure Caption CN", WD_STYLE_TYPE.PARAGRAPH)
    else:
        caption_style = styles["Figure Caption CN"]
    caption_style.font.name = "Arial Unicode MS"
    caption_style.font.size = Pt(9.5)
    caption_style.font.color.rgb = RGBColor.from_string("000000")
    caption_style._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial Unicode MS")
    caption_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    caption_style.paragraph_format.line_spacing = 1.2
    caption_style.paragraph_format.space_before = Pt(6)
    caption_style.paragraph_format.space_after = Pt(6)


def add_header_footer(section):
    header = section.header
    header.distance = Inches(0.42)
    paragraph = header.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.space_after = Pt(0)
    run = paragraph.add_run("科研进展通稿")
    set_run_font(run, east_asia="Heiti SC", size=8.5, bold=True, color="315F86")
    run = paragraph.add_run("  ·  中国科学院近代物理研究所")
    set_run_font(run, size=8.5, color="6B7280")

    footer = section.footer
    footer.distance = Inches(0.42)
    add_page_number(footer.paragraphs[0])


def set_image_alt_text(inline_shape, title, description):
    doc_pr = inline_shape._inline.docPr
    doc_pr.set("title", title)
    doc_pr.set("descr", description)


def add_figure_page(document, image_path, caption, alt_title, alt_description, width):
    document.add_section(WD_SECTION.NEW_PAGE)
    image_paragraph = document.add_paragraph()
    image_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    image_paragraph.paragraph_format.space_before = Pt(10)
    image_paragraph.paragraph_format.space_after = Pt(0)
    inline_shape = image_paragraph.add_run().add_picture(str(image_path), width=Inches(width))
    set_image_alt_text(inline_shape, alt_title, alt_description)

    caption_paragraph = document.add_paragraph(style="Figure Caption CN")
    caption_paragraph.add_run(caption)
    set_paragraph_language(caption_paragraph)


def build_document():
    if not FIGURE_1.exists() or not FIGURE_2.exists():
        missing = [str(path) for path in (FIGURE_1, FIGURE_2) if not path.exists()]
        raise FileNotFoundError(f"Missing figure(s): {missing}")

    document = Document()
    section = document.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.gutter = Inches(0)
    configure_styles(document)

    properties = document.core_properties
    properties.title = TITLE
    properties.subject = "氧化镓神经演化势与快重离子辐照径迹结构科研进展通稿"
    properties.author = "中国科学院近代物理研究所研究团队"
    properties.keywords = "氧化镓；神经演化势；快重离子；离子径迹；分子动力学"
    properties.comments = "供科研新闻宣传与审校使用"

    title_paragraph = document.add_paragraph(style="Title")
    title_paragraph.add_run(TITLE)
    set_paragraph_language(title_paragraph)

    for index, text in enumerate(BODY_PARAGRAPHS):
        paragraph = document.add_paragraph()
        paragraph.paragraph_format.first_line_indent = Pt(22)
        paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        paragraph.paragraph_format.line_spacing = 1.28
        paragraph.paragraph_format.space_after = Pt(6)
        run = paragraph.add_run(text)
        set_run_font(run, size=10.5)
        set_paragraph_language(paragraph)

    link_paragraph = document.add_paragraph()
    link_paragraph.paragraph_format.first_line_indent = Pt(22)
    link_paragraph.paragraph_format.space_before = Pt(1)
    link_paragraph.paragraph_format.space_after = Pt(0)
    run = link_paragraph.add_run("论文链接：")
    set_run_font(run, size=10, bold=False)
    run = link_paragraph.add_run(DOI_URL)
    set_run_font(run, size=10)

    add_figure_page(
        document,
        FIGURE_1,
        "图1. 面向快重离子辐照物理过程的氧化镓神经演化势训练及应用流程示意图。",
        "氧化镓神经演化势训练及应用流程",
        "流程图展示从第一性原理数据、过程导向采样和能量依赖加权，到神经演化势训练、"
        "模型验证及快重离子辐照分子动力学模拟的完整路径。",
        width=6.45,
    )
    add_figure_page(
        document,
        FIGURE_2,
        "图2. 不同电子能损条件下模拟得到的离子径迹直径与已有透射电子显微镜实验结果的对比。",
        "不同电子能损条件下的离子径迹直径对比",
        "散点图比较不同电子能损条件下分子动力学模拟与透射电子显微镜实验获得的"
        "离子径迹直径，显示模拟结果与实验结果具有良好一致性。",
        width=6.25,
    )
    document.save(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    build_document()
