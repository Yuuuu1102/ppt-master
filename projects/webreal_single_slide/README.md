# WebReal single-slide export

This directory contains a standalone script that turns a reference image into a single-slide PPTX using the repository's PowerPoint stack.

## Usage

```bash
python projects/webreal_single_slide/generate_single_slide.py \
  --input /path/to/reference_image.png \
  --output exports/webreal_single_slide.pptx
```

The script does the following:

- reads the supplied screenshot
- crops the banner into three equal horizontal panels
- applies the requested `#00ab84` theme color
- writes a clean single-page deck matching the provided layout closely
- uses `Source Han Sans CN` / `Noto Sans CJK SC` fallback fonts when available

## Notes

- The repo must have Python dependencies installed first:

```bash
pip install -r requirements.txt
```

- If you want a more exact match, provide a higher-resolution source image and adjust the
  text block widths/spacing in `generate_single_slide.py`.
