"""
生成单页 PPT 的脚本：从原始图片裁切为三等分横向片段，合并为横幅并生成 single-slide PPTX。
主色已设置为 #00ab84，字体默认使用 思源黑体 (Source Han Sans CN)——如果系统中无该字体，会使用 Arial 作为回退。

用法：
1) 把原始图片放到仓库路径 images/original.jpg（或修改 INPUT_IMAGE 路径）
2) 在虚拟环境中安装依赖：pip install pillow python-pptx
3) 运行：python scripts/generate_ppt_from_image.py
4) 输出文件：outputs/single_slide_webreal.pptx

说明：该脚本不依赖仓库特有的 ppt-master API，直接使用 python-pptx 生成 PPT，保证在任何环境可运行。
如果你仍希望改为调用 repo 内的 ppt-master 导出器，我可以再提交一个调用示例（前提是该导出器提供可用的 Python 接口）。

作者: AI Assistant
"""

from pathlib import Path
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# 配置
INPUT_IMAGE = Path("images/original.jpg")
OUTPUT_DIR = Path("outputs")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_PPTX = OUTPUT_DIR / "single_slide_webreal.pptx"
TEMP_BANNERS_DIR = Path(".tmp_banners")
TEMP_BANNERS_DIR.mkdir(parents=True, exist_ok=True)

# 主题颜色 (#00ab84)
ACCENT_RGB = (0x00, 0xab, 0x84)  # (0,171,132)
BG_LIGHT = (245, 250, 250)

# 字体（思源黑体）
PREFERRED_FONT = "Source Han Sans CN"  # 常见系统名，中文环境可能为 "思源黑体" 或 "Source Han Sans CN"
FALLBACK_FONT = "Arial"

def crop_three_vertical_slices(img_path, out_dir):
    """将图片按宽度三等分并保存三张图片（左，中，右）。返回保存路径列表。"""
    im = Image.open(img_path).convert("RGB")
    w, h = im.size
    slice_w = w // 3
    paths = []
    for i in range(3):
        left = i * slice_w
        # 最后一块包含余数
        right = (i + 1) * slice_w if i < 2 else w
        box = (left, 0, right, h)
        cropped = im.crop(box)
        p = out_dir / f"banner_part_{i+1}.jpg"
        cropped.save(p, quality=90)
        paths.append(p)
    return paths


def combine_banners_horizontally(paths, out_path, target_height=420):
    # 等比缩放到相同高度并横向拼接
    imgs = [Image.open(p).convert("RGB") for p in paths]
    resized = []
    for im in imgs:
        w, h = im.size
        new_h = target_height
        new_w = int(w * new_h / h)
        resized.append(im.resize((new_w, new_h), Image.LANCZOS))
    total_w = sum(im.size[0] for im in resized)
    banner = Image.new("RGB", (total_w, target_height), (255,255,255))
    x = 0
    for im in resized:
        banner.paste(im, (x, 0))
        x += im.size[0]
    banner.save(out_path, quality=90)
    return out_path


def build_ppt(banner_image_path, output_pptx_path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

    slide_width = prs.slide_width
    slide_height = prs.slide_height

    # 背景
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, slide_width, slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(*BG_LIGHT)
    bg.line.fill.background()

    left_margin = Inches(0.5)
    right_margin = Inches(0.5)

    # 页眉
    hdr_height = Inches(0.8)
    hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left_margin, Inches(0.2), slide_width - left_margin - right_margin, hdr_height)
    hdr.fill.solid()
    hdr.fill.fore_color.rgb = RGBColor(255,255,255)
    hdr.line.color.rgb = RGBColor(*ACCENT_RGB)

    # Logo 文本
    logo_box = slide.shapes.add_textbox(left_margin + Inches(0.2), Inches(0.25), Inches(4.5), Inches(0.6))
    tb = logo_box.text_frame
    p = tb.paragraphs[0]
    p.text = "WebReal  唯哲科技"
    p.font.bold = True
    p.font.size = Pt(20)
    # 尝试设置首选字体，否则使用回退
    try:
        p.font.name = PREFERRED_FONT
    except Exception:
        p.font.name = FALLBACK_FONT
    p.font.color.rgb = RGBColor(*ACCENT_RGB)

    # 右侧小字
    hdr_right = slide.shapes.add_textbox(slide_width - right_margin - Inches(3.0), Inches(0.32), Inches(2.8), Inches(0.4))
    p2 = hdr_right.text_frame.paragraphs[0]
    p2.text = "2026 | 工业智能 赋能未来"
    p2.font.size = Pt(12)
    try:
        p2.font.name = PREFERRED_FONT
    except Exception:
        p2.font.name = FALLBACK_FONT
    p2.font.color.rgb = RGBColor(*ACCENT_RGB)
    p2.alignment = PP_ALIGN.RIGHT

    # 插入横幅
    banner_left = left_margin
    banner_top = Inches(1.15)
    banner_width = slide_width - left_margin - right_margin
    pic = slide.shapes.add_picture(str(banner_image_path), banner_left, banner_top, width=banner_width)

    # 主标题
    title_top = banner_top + Inches(1.8)
    title_box = slide.shapes.add_textbox(left_margin + Inches(0.1), title_top, slide_width - left_margin - right_margin - Inches(0.2), Inches(1.2))
    tf = title_box.text_frame
    tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = "让工业智能，被看见!"
    p.font.size = Pt(44)
    p.font.bold = True
    try:
        p.font.name = PREFERRED_FONT
    except Exception:
        p.font.name = FALLBACK_FONT
    p.font.color.rgb = RGBColor(*ACCENT_RGB)

    p_sub = tf.add_paragraph()
    p_sub.text = "做编程语言的苹果！"
    p_sub.font.size = Pt(18)
    p_sub.font.bold = False
    try:
        p_sub.font.name = PREFERRED_FONT
    except Exception:
        p_sub.font.name = FALLBACK_FONT
    p_sub.font.color.rgb = RGBColor(80,80,80)

    # 第二行说明
    desc_box = slide.shapes.add_textbox(left_margin + Inches(0.1), title_top + Inches(1.05), Inches(9.0), Inches(0.4))
    d_tf = desc_box.text_frame
    d_tf.text = "工业知识、设备、芯片、模型  ➜  WebReal 编程  ➜  工业智能体"
    d_tf.paragraphs[0].font.size = Pt(12)
    d_tf.paragraphs[0].font.color.rgb = RGBColor(80,80,80)
    try:
        d_tf.paragraphs[0].font.name = PREFERRED_FONT
    except Exception:
        d_tf.paragraphs[0].font.name = FALLBACK_FONT

    # 三个信息卡片
    card_top = title_top + Inches(1.6)
    card_height = Inches(2.1)
    card_width = (slide_width - left_margin - right_margin - Inches(0.4)) / 3

    # 左卡
    c1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left_margin + Inches(0.0), card_top, card_width, card_height)
    c1.fill.solid()
    c1.fill.fore_color.rgb = RGBColor(255,255,255)
    c1.line.color.rgb = RGBColor(220,220,220)
    tx = c1.text_frame
    tx.text = "国内开发者\n\n超1000万\n\n面向中国 STEM 开发者\n释放 AI 编程创造力\n跨语种编程 一带一路 走向全球"
    tx.word_wrap = True
    for i, p in enumerate(tx.paragraphs):
        if i == 1:
            p.font.size = Pt(28)
            p.font.bold = True
            p.font.color.rgb = RGBColor(*ACCENT_RGB)
        else:
            p.font.size = Pt(10)
            p.font.color.rgb = RGBColor(80,80,80)
        try:
            p.font.name = PREFERRED_FONT
        except Exception:
            p.font.name = FALLBACK_FONT

    # 中卡
    c2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left_margin + card_width + Inches(0.2), card_top, card_width, card_height)
    c2.fill.solid()
    c2.fill.fore_color.rgb = RGBColor(255,255,255)
    c2.line.color.rgb = RGBColor(220,220,220)
    tx2 = c2.text_frame
    tx2.text = "公司愿景\n\n营收超千亿\n市值过万亿\n\n重庆的科技名片"
    for i, p in enumerate(tx2.paragraphs):
        if i == 1:
            p.font.size = Pt(20)
            p.font.bold = True
            p.font.color.rgb = RGBColor(*ACCENT_RGB)
        else:
            p.font.size = Pt(11)
            p.font.color.rgb = RGBColor(80,80,80)
        try:
            p.font.name = PREFERRED_FONT
        except Exception:
            p.font.name = FALLBACK_FONT

    # 右卡
    c3 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left_margin + card_width*2 + Inches(0.4), card_top, card_width, card_height)
    c3.fill.solid()
    c3.fill.fore_color.rgb = RGBColor(255,255,255)
    c3.line.color.rgb = RGBColor(220,220,220)
    tx3 = c3.text_frame
    tx3.text = "商业模型\n\n长尾高附加值工业APP的编程开发平台\n\n毛利率 70%    净利率 25%\n客单价1万元 x 1000万开发者 = 1千亿"
    for i, p in enumerate(tx3.paragraphs):
        if i == 1:
            p.font.size = Pt(12)
            p.font.bold = True
            p.font.color.rgb = RGBColor(*ACCENT_RGB)
        else:
            p.font.size = Pt(10)
            p.font.color.rgb = RGBColor(80,80,80)
        try:
            p.font.name = PREFERRED_FONT
        except Exception:
            p.font.name = FALLBACK_FONT

    # 底部语言行
    lang_box = slide.shapes.add_textbox(left_margin + Inches(0.1), slide_height - Inches(0.8), Inches(8.5), Inches(0.5))
    lang_tf = lang_box.text_frame
    lang_tf.text = "hello · こんにちは · 안녕하세요 · ciao · hola · bonjour · olá"
    lang_tf.paragraphs[0].font.size = Pt(9)
    lang_tf.paragraphs[0].font.color.rgb = RGBColor(160,160,160)
    try:
        lang_tf.paragraphs[0].font.name = PREFERRED_FONT
    except Exception:
        lang_tf.paragraphs[0].font.name = FALLBACK_FONT

    prs.save(str(output_pptx_path))
    print(f"已保存 PPT: {output_pptx_path}")


if __name__ == "__main__":
    if not INPUT_IMAGE.exists():
        print(f"请先把原始图片放到: {INPUT_IMAGE}，脚本退出。")
        raise SystemExit(1)

    parts = crop_three_vertical_slices(INPUT_IMAGE, TEMP_BANNERS_DIR)
    combined = TEMP_BANNERS_DIR / "banner_combined.jpg"
    combine_banners_horizontally(parts, combined, target_height=420)
    build_ppt(combined, OUTPUT_PPTX)
    print("完成。输出文件位于:", OUTPUT_PPTX)
