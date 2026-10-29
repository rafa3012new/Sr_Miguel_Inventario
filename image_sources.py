import re
from pathlib import Path

BASE = Path(__file__).resolve().parent
IMAGES_DIR = "imagenes"


def safe_name(s):
    s = re.sub(r'[\\/*?:"<>|]', "_", str(s))
    s = re.sub(r"\s+", "_", s).strip("_")
    return (s[:50] or "item")

MANUAL_IMAGE_OVERRIDES = {
    ("V322453M", "STO"): "imagenes/V322453M.jpg",
}

STYLE_IMAGE_OVERRIDES = {
    "119004414": "imagenes/119004414.png",
    "117004093": "imagenes/117004093.jpg",
    "119000214": "imagenes/119000214.jpg",
    "119000455": "imagenes/119000455.jpg",
    "119001566": "imagenes/119001566.jpg",
    "119001328": "imagenes/119001328.jpg",
    "119005385": "imagenes/119005385.jpg",
    "121000428": "imagenes/121000428.jpg",
}


def extract_imagenes():
    out_dir = BASE / IMAGES_DIR
    mapping = {}
    if not out_dir.exists():
        return mapping
    for f in sorted(out_dir.iterdir()):
        if not f.is_file():
            continue
        if f.suffix.lower() not in (".png", ".jpg", ".jpeg", ".webp"):
            continue
        if f.stem not in mapping:
            mapping[f.stem] = f"{IMAGES_DIR}/{f.name}"
    return mapping


def build_maps():
    return {"imagenes": extract_imagenes()}


def combined_image(r, maps):
    _style = str(r.get("Style", "")).strip()
    _color = str(r.get("Color_code", "")).strip().upper()
    _brand = str(r.get("BRAND_NAME", "")).strip()
    if (_style, _color) in MANUAL_IMAGE_OVERRIDES:
        return MANUAL_IMAGE_OVERRIDES[(_style, _color)]
    if _style in STYLE_IMAGE_OVERRIDES:
        return STYLE_IMAGE_OVERRIDES[_style]
    brand_fname = maps["imagenes"].get(f"{safe_name(_brand)}_{_style}", "")
    if brand_fname:
        return brand_fname
    return maps["imagenes"].get(_style, "")
