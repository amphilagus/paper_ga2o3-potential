from pathlib import Path
import shutil

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from docx.text.paragraph import Paragraph
from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
RESPONSE_DIR = ROOT / "第二轮意见回复"
STAGING = RESPONSE_DIR / "response_letter_figures_to_insert"
PROCESSED = STAGING / "processed"
SOURCE_DOCX = RESPONSE_DIR / "Response to Reviewer Comments_Round2.docx"
OUT_DOCX = RESPONSE_DIR / "Response to Reviewer Comments_Round2_with_figures.docx"


def paragraph_after(paragraph, text=None, style=None):
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    if text:
        new_para.text = text
    if style:
        new_para.style = style
    return new_para


def set_keep_with_next(paragraph):
    paragraph.paragraph_format.keep_with_next = True


def caption(paragraph, text):
    run = paragraph.add_run(text)
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    run.font.size = Pt(9)
    run.font.italic = True
    run.font.color.rgb = RGBColor(80, 80, 80)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(3)
    paragraph.paragraph_format.space_after = Pt(8)


def add_image_block(after_paragraph, image_path, caption_text, width_inches=6.2):
    cap = paragraph_after(after_paragraph)
    caption(cap, caption_text)
    set_keep_with_next(cap)

    fig_p = paragraph_after(cap)
    fig_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fig_p.paragraph_format.space_before = Pt(0)
    fig_p.paragraph_format.space_after = Pt(8)
    run = fig_p.add_run()
    run.add_picture(str(image_path), width=Inches(width_inches))
    return fig_p


def copy_and_prepare_images():
    PROCESSED.mkdir(parents=True, exist_ok=True)

    # Reviewer #3 Comment 4 asks only about RDF panels, so prepare a cropped
    # response-letter image from the full Fig. 7 preview.
    fig7 = Image.open(PROCESSED / "Figure_7_response.png").convert("RGB")
    crop = fig7.crop((0, 820, 1325, fig7.height))
    crop.save(PROCESSED / "Figure_7_c-e_response.png", optimize=True)

    # Build a compact 2x2 preview of Supplementary Figs. S8-S11. This is only
    # for the response letter; formal submission figures remain separate.
    items = [
        ("Supplementary Fig. S8", PROCESSED / "Figure_S8_response.png"),
        ("Supplementary Fig. S9", PROCESSED / "Figure_S9_response.png"),
        ("Supplementary Fig. S10", PROCESSED / "Figure_S10_response.png"),
        ("Supplementary Fig. S11", PROCESSED / "Figure_S11_response.png"),
    ]
    cell_w, cell_h = 1260, 980
    margin, gutter, title_h = 60, 40, 48
    canvas_w = margin * 2 + cell_w * 2 + gutter
    canvas_h = margin * 2 + cell_h * 2 + gutter
    canvas = Image.new("RGB", (canvas_w, canvas_h), "white")
    draw = ImageDraw.Draw(canvas)
    try:
        font = ImageFont.truetype("Arial.ttf", 32)
    except OSError:
        font = ImageFont.load_default()

    for idx, (label, path) in enumerate(items):
        row, col = divmod(idx, 2)
        x = margin + col * (cell_w + gutter)
        y = margin + row * (cell_h + gutter)
        draw.text((x, y), label, fill=(50, 50, 50), font=font)
        im = Image.open(path).convert("RGB")
        im.thumbnail((cell_w, cell_h - title_h), Image.LANCZOS)
        ix = x + (cell_w - im.width) // 2
        iy = y + title_h + (cell_h - title_h - im.height) // 2
        canvas.paste(im, (ix, iy))
        draw.rectangle((x, y + title_h, x + cell_w, y + cell_h), outline=(210, 210, 210), width=2)

    canvas.save(PROCESSED / "Figure_S8-S11_composite_response.png", optimize=True)


def last_before(doc, start_text, stop_text):
    started = False
    last = None
    for para in doc.paragraphs:
        text = para.text.strip()
        if text == start_text:
            started = True
            last = para
            continue
        if started and text == stop_text:
            return last
        if started and text:
            last = para
    raise ValueError(f"Could not locate block {start_text!r} before {stop_text!r}")


def main():
    copy_and_prepare_images()
    shutil.copy2(SOURCE_DOCX, OUT_DOCX)
    doc = Document(OUT_DOCX)

    insertions = [
        (
            "Revision actions 1",
            "Comment 2",
            [
                (
                    PROCESSED / "Figure_S5_response.png",
                    "Supplementary Fig. S5. Geometry-optimization check for the equation-of-state comparison.",
                    6.2,
                ),
                (
                    PROCESSED / "Figure_S6_response.png",
                    "Supplementary Fig. S6. Controlled gamma-phase configuration-sensitivity test.",
                    6.2,
                ),
            ],
        ),
        (
            "Revision actions 2",
            "Comment 3",
            [
                (
                    PROCESSED / "Figure_S2_response.png",
                    "Supplementary Fig. S2. Linear-scale pair-energy curves for Ga-Ga, Ga-O, and O-O dimers.",
                    6.2,
                ),
                (
                    PROCESSED / "Figure_S3_response.png",
                    "Supplementary Fig. S3. Short-range/logarithmic pair-energy comparison.",
                    6.2,
                ),
            ],
        ),
        (
            "Revision actions 3",
            "Comment 4",
            [
                (
                    PROCESSED / "Figure_S8-S11_composite_response.png",
                    "Supplementary Figs. S8-S11. Compact preview of the noncrystalline Ga2O3 validation benchmarks.",
                    6.3,
                )
            ],
        ),
        (
            "Revision actions 4",
            "Minor or Recommended Revisions",
            [
                (
                    PROCESSED / "Figure_7_c-e_response.png",
                    "Fig. 7(c-e). Same-temperature RDF results reproduced for reviewer convenience.",
                    6.0,
                )
            ],
        ),
    ]

    for start, stop, figures in reversed(insertions):
        anchor = last_before(doc, start, stop)
        current = anchor
        for image, cap, width in figures:
            current = add_image_block(current, image, cap, width)

    doc.save(OUT_DOCX)
    print(OUT_DOCX)


if __name__ == "__main__":
    main()
