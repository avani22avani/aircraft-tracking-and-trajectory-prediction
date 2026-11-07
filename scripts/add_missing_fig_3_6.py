from pathlib import Path
import sys

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.shared import Inches
from docx.text.paragraph import Paragraph
from PIL import Image, ImageDraw, ImageFont


SOURCE = Path(sys.argv[1])
OUTPUT = Path(sys.argv[2])
ASSET_PATH = OUTPUT.parent / "figures" / "fig_3_6_regression_prediction.png"
CAPTION = (
    "Fig. 3.6: Regression-based short-outage prediction. Recent latitude, "
    "longitude and altitude samples define least-squares trends; "
    "constant-velocity kinematics provide a fallback when a reliable fit cannot be produced."
)


def load_font(size, bold=False):
    candidates = [
        "arialbd.ttf" if bold else "arial.ttf",
        "calibrib.ttf" if bold else "calibri.ttf",
    ]
    for name in candidates:
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            pass
    return ImageFont.load_default()


def text_box(draw, xy, text, font, fill, outline, text_fill=(20, 30, 40), pad=12):
    x, y = xy
    bbox = draw.multiline_textbbox((0, 0), text, font=font, spacing=4)
    w = bbox[2] - bbox[0] + 2 * pad
    h = bbox[3] - bbox[1] + 2 * pad
    draw.rounded_rectangle((x, y, x + w, y + h), radius=12, fill=fill, outline=outline, width=2)
    draw.multiline_text((x + pad, y + pad), text, font=font, fill=text_fill, spacing=4)
    return (x, y, x + w, y + h)


def arrow(draw, start, end, fill, width=5):
    draw.line((start[0], start[1], end[0], end[1]), fill=fill, width=width)
    ex, ey = end
    sx, sy = start
    dx = ex - sx
    dy = ey - sy
    length = max((dx * dx + dy * dy) ** 0.5, 1)
    ux, uy = dx / length, dy / length
    px, py = -uy, ux
    size = 16
    points = [
        (ex, ey),
        (ex - ux * size + px * size * 0.55, ey - uy * size + py * size * 0.55),
        (ex - ux * size - px * size * 0.55, ey - uy * size - py * size * 0.55),
    ]
    draw.polygon(points, fill=fill)


def make_figure(path):
    path.parent.mkdir(parents=True, exist_ok=True)
    width, height = 1500, 900
    img = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(img)

    title_font = load_font(36, bold=True)
    head_font = load_font(25, bold=True)
    body_font = load_font(22)
    small_font = load_font(19)

    draw.text((50, 35), "Regression-Based Short-Outage Prediction", font=title_font, fill=(12, 45, 75))
    draw.text(
        (50, 82),
        "Recent real ADS-B samples are used first. If the fit is weak, speed and heading are used as a fallback.",
        font=body_font,
        fill=(70, 80, 90),
    )

    # Main graph area.
    gx, gy, gw, gh = 90, 180, 900, 565
    draw.rounded_rectangle((gx, gy, gx + gw, gy + gh), radius=10, fill=(248, 252, 255), outline=(120, 150, 180), width=2)
    draw.line((gx + 70, gy + gh - 70, gx + gw - 45, gy + gh - 70), fill=(45, 60, 75), width=3)
    draw.line((gx + 70, gy + gh - 70, gx + 70, gy + 45), fill=(45, 60, 75), width=3)
    draw.text((gx + gw // 2 - 70, gy + gh - 35), "Time", font=small_font, fill=(45, 60, 75))
    draw.text((gx + 12, gy + 50), "Position / altitude", font=small_font, fill=(45, 60, 75))

    # Grid.
    for i in range(1, 8):
        x = gx + 70 + i * 95
        draw.line((x, gy + 55, x, gy + gh - 70), fill=(222, 232, 242), width=1)
    for i in range(1, 5):
        y = gy + 70 + i * 90
        draw.line((gx + 70, y, gx + gw - 45, y), fill=(222, 232, 242), width=1)

    # Sample points and trend.
    samples = [(170, 590), (270, 552), (370, 512), (470, 474), (570, 435), (670, 397)]
    samples = [(gx + x, gy + y - 180) for x, y in samples]
    for point in samples:
        draw.ellipse((point[0] - 9, point[1] - 9, point[0] + 9, point[1] + 9), fill=(22, 126, 184), outline="white", width=3)
    draw.line(samples, fill=(22, 126, 184), width=4)
    draw.line((samples[0][0] - 25, samples[0][1] + 12, samples[-1][0] + 240, samples[-1][1] - 90), fill=(229, 90, 36), width=5)

    pred = (samples[-1][0] + 230, samples[-1][1] - 85)
    draw.ellipse((pred[0] - 14, pred[1] - 14, pred[0] + 14, pred[1] + 14), fill=(229, 90, 36), outline="white", width=4)
    draw.text((pred[0] - 58, pred[1] - 55), "Predicted\npoint", font=small_font, fill=(170, 55, 25), spacing=2)

    # Fallback vector.
    fallback_start = samples[-1]
    fallback_end = (fallback_start[0] + 210, fallback_start[1] + 20)
    draw.line((fallback_start[0], fallback_start[1], fallback_end[0], fallback_end[1]), fill=(70, 130, 80), width=4)
    arrow(draw, fallback_start, fallback_end, (70, 130, 80), width=4)
    draw.text((fallback_end[0] - 60, fallback_end[1] + 20), "Fallback:\nspeed + heading", font=small_font, fill=(50, 105, 60), spacing=2)

    text_box(draw, (80, 770), "Blue dots = recent real samples\nOrange line = least-squares trend\nOrange dot = predicted short-gap position", small_font, (255, 247, 222), (210, 150, 50))

    # Processing blocks.
    blocks = [
        ("1. Recent trajectory\nsamples from database", (1080, 170)),
        ("2. Fit simple trends\nlat(t), lon(t), alt(t)", (1080, 300)),
        ("3. Estimate short-gap\nfuture position", (1080, 430)),
        ("4. If fit is weak,\nuse kinematic fallback", (1080, 560)),
    ]
    centers = []
    for label, pos in blocks:
        box = text_box(draw, pos, label, head_font, (235, 245, 255), (55, 110, 160))
        centers.append(((box[0] + box[2]) // 2, box[3]))
    for a, b in zip(centers, centers[1:]):
        arrow(draw, (a[0], a[1] + 5), (b[0], b[1] - 95), (70, 90, 120), width=4)

    text_box(
        draw,
        (1050, 720),
        "Output shown on dashboard as\nPRED aircraft and dotted path.\nPredicted values are not stored\nas real training samples.",
        body_font,
        (237, 255, 238),
        (70, 140, 80),
    )

    img.save(path)


def insert_paragraph_before(paragraph):
    new_p = OxmlElement("w:p")
    paragraph._p.addprevious(new_p)
    return Paragraph(new_p, paragraph._parent)


def main():
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    make_figure(ASSET_PATH)

    doc = Document(SOURCE)
    caption_index = None
    caption_para = None
    for idx, paragraph in enumerate(doc.paragraphs):
        if CAPTION in paragraph.text:
            caption_index = idx
            caption_para = paragraph
            break
    if caption_para is None:
        raise RuntimeError("Fig. 3.6 caption not found")

    # Avoid adding the same figure twice if the script is re-run on an already
    # updated document: if the paragraph before the caption contains a drawing,
    # assume Fig. 3.6 is already present.
    already_inserted = False
    if caption_index > 0:
        already_inserted = bool(doc.paragraphs[caption_index - 1]._p.xpath(".//w:drawing"))
    if not already_inserted:
        image_para = insert_paragraph_before(caption_para)
        image_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        image_para.add_run().add_picture(str(ASSET_PATH), width=Inches(5.7))

    doc.save(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    main()
