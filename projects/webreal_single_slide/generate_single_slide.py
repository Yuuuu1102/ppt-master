#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generate a single-slide PPT matching the supplied reference image.

Usage:
    python projects/webreal_single_slide/generate_single_slide.py \
        --input /path/to/source_image.png \
        --output /path/to/webreal_single_slide.pptx

The script assumes the source image is the reference screenshot. It crops the banner
area into three equal horizontal sections, styles the slide with the requested
#00ab84 accent, and writes a PowerPoint file using python-pptx.
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import List

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

ACCENT_HEX = "#00ab84"
ACCENT_RGB = tuple(int(ACCENT_HEX.lstrip("#")[i : i + 2], 16) for i in (0, 2, 4))
FONT_CANDIDATES = ["Source Han Sans CN", "Noto Sans CJK SC", "PingFang SC", "Microsoft YaHei", "Arial"]


def pick_font() -> str:
    return FONT_CANDIDATES[0]


def crop_banner_panels(image_path: Path, out_dir: Path) -> List[Path]:
    img = Image.open(image_path).convert("RGB")
    width, height = img.size
    third = width // 3

    out_files: List[Path] = []
    for idx in range(3):
        left = idx * third
        right = width if idx == 2 else (idx + 1) * third
        crop = img.crop((left, 0, right, height))
        out_path = out_dir / f"banner_panel_{idx + 1}.png"
        crop.save(out_path)
        out_files.append(out_path)
    return out_files


def add_textbox(slide, left, top, width, height, text, font_size=18, bold=False, color=(80, 80, 80), align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0

    p = tf.paragraphs[0]
    p.text = text
    p.alignment = align
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.name = pick_font()
    p.font.color.rgb = RGBColor(*color)
    return box


def add_labelled_card(slide, left, top, width, height, title, body_lines):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
    shape.line.color.rgb = RGBColor(220, 220, 220)
    shape.line.width = Pt(1)

    tf = shape.text_frame
    tf.word_wrap = True
    tf.clear()

    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.name = pick_font()
    p.font.color.rgb = RGBColor(80, 80, 80)

    for line in body_lines:
        p2 = tf.add_paragraph()
        p2.text = line
        p2.level = 0
        p2.font.size = Pt(12)
        p2.font.name = pick_font()
        p2.font.color.rgb = RGBColor(85, 85, 85)

    return shape


def build_slide(input_image: Path, output_pptx: Path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_width = prs.slide_width
    slide_height = prs.slide_height

    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, slide_width, slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(244, 248, 250)
    bg.line.fill.background()

    left_margin = Inches(0.45)
    right_margin = Inches(0.45)

    header = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left_margin, Inches(0.18), slide_width - left_margin - right_margin, Inches(0.72))
    header.fill.solid()
    header.fill.fore_color.rgb = RGBColor(255, 255, 255)
    header.line.color.rgb = RGBColor(*ACCENT_RGB)
    header.line.width = Pt(1)

    add_textbox(slide, left_margin + Inches(0.22), Inches(0.23), Inches(4.7), Inches(0.45), "WebReal  唯哲科技", 20, True, ACCENT_RGB)
    add_textbox(slide, slide_width - right_margin - Inches(3.2), Inches(0.29), Inches(2.8), Inches(0.35), "2026 | 工业智能 赋能未来", 12, False, ACCENT_RGB, PP_ALIGN.RIGHT)

    panels = crop_banner_panels(input_image, input_image.parent)
    banner_top = Inches(1.02)
    panel_width = (slide_width - left_margin - right_margin) / 3
    for index, panel_path in enumerate(panels):
        slide.shapes.add_picture(str(panel_path), left_margin + Inches(index * (panel_width / Inches(1))), banner_top, width=Inches(panel_width / Inches(1) - 0.02))

    title_y = Inches(2.92)
    add_textbox(slide, left_margin + Inches(0.12), title_y, Inches(8.7), Inches(0.7), "让工业智能，被看见!", 28, True, ACCENT_RGB)
    add_textbox(slide, left_margin + Inches(0.12), title_y + Inches(0.62), Inches(8.2), Inches(0.45), "做编程语言的苹果！", 18, False, (80, 80, 80))
    add_textbox(slide, left_margin + Inches(0.14), title_y + Inches(1.0), Inches(9.2), Inches(0.35), "工业知识、设备、芯片、模型  ➜  WebReal 编程  ➜  工业智能体", 12, False, (90, 90, 90))

    card_top = Inches(4.0)
    card_width = (slide_width - left_margin - right_margin - Inches(0.5)) / 3

    add_labelled_card(slide, left_margin, card_top, card_width, Inches(2.25), "国内开发者", ["超1000万", "面向中国 STEM 开发者", "释放 AI 编程创造力", "跨语种编程 一带一路 走向全球"])
    add_labelled_card(slide, left_margin + card_width + Inches(0.25), card_top, card_width, Inches(2.25), "公司愿景", ["营收超千亿", "市值过万亿", "重庆的科技名片"])
    add_labelled_card(slide, left_margin + 2 * card_width + Inches(0.5), card_top, card_width, Inches(2.25), "商业模型", ["长尾高附加值工业 APP 的编程开发平台", "毛利率 70%    净利率 25%", "客单价 1 万元 × 1000 万开发者 = 1 千亿"])

    add_textbox(slide, left_margin + Inches(0.1), slide_height - Inches(0.7), Inches(7.5), Inches(0.3), "hello · こんにちは · 안녕하세요 · ciao · hola · bonjour · olá", 9, False, (160, 160, 160))

    output_pptx.parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_pptx)


def main():
    parser = argparse.ArgumentParser(description="Generate a single-page WebReal-style PPT from a source image.")
    parser.add_argument("--input", required=True, help="Path to the reference screenshot image.")
    parser.add_argument("--output", default="exports/webreal_single_slide.pptx", help="Output PPTX path.")
    args = parser.parse_args()

    input_image = Path(args.input).expanduser().resolve()
    output_pptx = Path(args.output).expanduser().resolve()
    if not input_image.exists():
        raise FileNotFoundError(f"Input image not found: {input_image}")

    build_slide(input_image, output_pptx)
    print(f"PPT exported to: {output_pptx}")


if __name__ == "__main__":
    main()

