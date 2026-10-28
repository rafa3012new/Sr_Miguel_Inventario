"""
Image Generator API - builds Inventario_con_imagenes.xlsx with a single
'Imagen' column pointing to the same images used by the HTML gallery
(imagenes_excel / imagenes_pdf / imagenes_xls), instead of downloading
images from the internet.

Usage:
    python main.py
"""

import argparse
import re
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent
EXCEL_IN = BASE_DIR / "sources" / "Inventario.xlsx"
EXCEL_OUT = BASE_DIR / "Inventario_con_imagenes.xlsx"
SHEET_NAME = "Data"

# Same sources/dirs as make_gallery.py
IMAGE_COL = "Imagen"
EXCEL_IMG_SRC = "sources/NAUTICA ATS TOM .xlsx"
EXCEL_IMG_DIR = "imagenes_excel"
PDF_IMG_SRC = "sources/MMX FALL ATS 6.23 LEVY GROUP 3.pdf"
PDF_IMG_SRC2 = "sources/LAUNDRY EARLY BUY LINESHEET 3.18.pdf"
PDF_IMG_DIR = "imagenes_pdf"
XLS_IMG_DIR = "imagenes_xls"
ATS_IMG_DIR = "imagenes_ats"

REQUIRED_COLS = ["TLGDIV_NAME", "Style", "Style_Name", "Color_code"]


def safe_name(s):
    s = re.sub(r'[\\/*?:"<>|]', "_", str(s))
    s = re.sub(r"\s+", "_", s).strip("_")
    return (s[:50] or "item")


def is_missing(v):
    if v is None:
        return True
    try:
        if pd.isna(v):
            return True
    except (TypeError, ValueError):
        pass
    s = str(v).strip()
    return s == "" or s.lower() in ("nan", "(blank)", "(en blanco)", "none", "nat")


try:
    import pymupdf
except Exception:
    pymupdf = None


def extract_excel_images():
    src = BASE_DIR / EXCEL_IMG_SRC
    out_dir = BASE_DIR / EXCEL_IMG_DIR
    out_dir.mkdir(exist_ok=True)
    mapping = {}
    if not src.exists():
        return mapping
    from openpyxl import load_workbook as _lw
    wb = _lw(src)
    ws = wb["Sheet"] if "Sheet" in wb.sheetnames else wb.worksheets[1]
    for im in ws._images:
        anchor = im.anchor
        from_cell = getattr(anchor, "_from", None)
        if from_cell is None:
            continue
        excel_row = from_cell.row + 1
        tl = ws.cell(row=excel_row, column=1).value
        style = ws.cell(row=excel_row, column=2).value
        if tl is None or style is None:
            continue
        style_s = str(style).strip()
        if "total" in style_s.lower():
            continue
        key = (str(tl).strip(), style_s)
        fname = f"excel_{safe_name(tl)}_{safe_name(style)}.png"
        dest = out_dir / fname
        if not dest.exists():
            try:
                dest.write_bytes(im._data())
            except Exception:
                continue
        mapping[key] = fname
    return mapping


def extract_pdf_images():
    src = BASE_DIR / PDF_IMG_SRC
    out_dir = BASE_DIR / PDF_IMG_DIR
    out_dir.mkdir(exist_ok=True)
    mapping = {}
    if pymupdf is None or not src.exists():
        return mapping
    doc = pymupdf.open(src)
    candidates = {}
    for pno in range(doc.page_count):
        page = doc[pno]
        style_blocks = []
        for b in page.get_text("blocks"):
            x0, y0, x1, y1, txt = b[:5]
            txt = txt.strip()
            m = re.search(r"\b(\d{7,})\b", txt)
            if m and x0 > 180:
                style_blocks.append((y0, y1, m.group(1)))
        for img in page.get_images(full=True):
            for r in page.get_image_rects(img[0]):
                styles = {sb[2] for sb in style_blocks if not (sb[1] < r.y0 or sb[0] > r.y1)}
                for st in styles:
                    candidates.setdefault(st, []).append((r.x0, r.y0, img[0]))
    for style, cands in candidates.items():
        cands_sorted = sorted(cands, key=lambda c: (c[0], c[1]))
        _, _, xref = cands_sorted[0]
        base = doc.extract_image(xref)
        ext = base.get("ext", "png")
        fname = f"pdf_{safe_name(style)}.{ext}"
        dest = out_dir / fname
        if not dest.exists():
            try:
                dest.write_bytes(base["image"])
            except Exception:
                continue
        mapping[style] = fname
    doc.close()
    return mapping


def extract_pdf_images2():
    src = BASE_DIR / PDF_IMG_SRC2
    out_dir = BASE_DIR / PDF_IMG_DIR
    out_dir.mkdir(exist_ok=True)
    mapping = {}
    if pymupdf is None or not src.exists():
        return mapping
    doc = pymupdf.open(src)
    used = set()
    for pno in range(doc.page_count):
        page = doc[pno]
        styles = []
        for b in page.get_text("blocks"):
            x0, y0, x1, y1, txt = b[:5]
            m = re.search(r"\b(\d{7,})\b", txt.strip())
            if m:
                styles.append((m.group(1), (x0 + x1) / 2, y0, y1))
        imgs = []
        for img in page.get_images(full=True):
            for r in page.get_image_rects(img[0]):
                imgs.append((img[0], (r.x0 + r.x1) / 2, r.y0, r.y1))
        for style, scx, sy0, sy1 in styles:
            cands = []
            for xref, icx, iy0, iy1 in imgs:
                if xref in used:
                    continue
                if iy0 >= sy1 - 5 and abs(icx - scx) <= 90:
                    cands.append((abs(iy0 - sy1) + abs(icx - scx) * 0.5, xref))
            if cands:
                cands.sort()
                xref = cands[0][1]
                used.add(xref)
                base = doc.extract_image(xref)
                ext = base.get("ext", "png")
                fname = f"pdf_{safe_name(style)}.{ext}"
                dest = out_dir / fname
                if not dest.exists():
                    try:
                        dest.write_bytes(base["image"])
                    except Exception:
                        continue
                mapping[style] = fname
    doc.close()
    return mapping


def extract_xls_images():
    out_dir = BASE_DIR / XLS_IMG_DIR
    mapping = {}
    if not out_dir.exists():
        return mapping
    for f in out_dir.iterdir():
        if not f.is_file():
            continue
        name = f.name
        if not name.startswith("xls_") or not name.lower().endswith((".png", ".jpg", ".jpeg")):
            continue
        key = name[4:].rsplit(".", 1)[0]
        mapping[key] = f"{XLS_IMG_DIR}/{name}"
    if "V623040" in mapping:
        mapping.setdefault("V623040A", mapping["V623040"])
        mapping.setdefault("V623040H", mapping["V623040"])
    return mapping


def extract_ats_images():
    out_dir = BASE_DIR / ATS_IMG_DIR
    mapping = {}
    if not out_dir.exists():
        return mapping
    by_style = {}
    for f in out_dir.iterdir():
        if not f.is_file():
            continue
        name = f.name
        if not name.startswith("ATS_") or not name.lower().endswith((".png", ".jpg", ".jpeg")):
            continue
        stem = name[4:].rsplit(".", 1)[0]
        parts = stem.split("_")
        if len(parts) < 2:
            continue
        style = parts[0]
        idx = 0
        if len(parts) >= 3 and parts[-1].isdigit():
            idx = int(parts[-1])
        by_style.setdefault(style, []).append((idx, name))
    for style, cands in by_style.items():
        cands.sort()
        mapping[style] = f"{ATS_IMG_DIR}/{cands[0][1]}"
    return mapping


MANUAL_IMAGE_OVERRIDES = {
    ("V322453M", "STO"): "V322453M_COLOR_STONE.png",
}


def combined_image(r, image_map, pdf_map, xls_map, ats_map=None):
    _style = str(r.get("Style", "")).strip()
    _color = str(r.get("Color_code", "")).strip().upper()
    if (_style, _color) in MANUAL_IMAGE_OVERRIDES:
        return MANUAL_IMAGE_OVERRIDES[(_style, _color)]
    # Imagen descargada de Rainforest/Amazon para este style
    if _style == "119004414":
        return "imagenes/119004414.png"
    if _style == "119000214":
        return "imagenes/119000214.jpg"
    if _style == "119000455":
        return "imagenes/119000455.jpg"
    if _style == "119001566":
        return "imagenes/119001566.jpg"
    if _style == "119001328":
        return "imagenes/119001328.jpg"
    if _style == "119005385":
        return "imagenes/119005385.jpg"
    if _style == "121000428":
        return "imagenes/121000428.jpg"
    excel_fname = image_map.get(
        (str(r.get("TLGDIV_NAME", "")).strip(), _style), ""
    )
    if excel_fname:
        return f"{EXCEL_IMG_DIR}/{excel_fname}"
    pdf_fname = pdf_map.get(_style, "")
    if pdf_fname:
        return f"{PDF_IMG_DIR}/{pdf_fname}"
    xls_fname = xls_map.get(_style, "")
    if xls_fname:
        return xls_fname
    if ats_map:
        ats_fname = ats_map.get(_style, "")
        if ats_fname:
            return ats_fname
    return ""


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true", help="Compatibilidad (ya no descarga).")
    parser.add_argument("--limit", type=int, default=None, help="Compatibilidad.")
    args = parser.parse_args()

    if not EXCEL_IN.exists():
        raise SystemExit(f"No se encontro {EXCEL_IN}")

    # Read Excel
    df = pd.read_excel(EXCEL_IN, sheet_name=SHEET_NAME, header=2, usecols="A:H")
    df.columns = ["TLGDIV_NAME", "Style", "Style_Name", "Color_code",
                  "Dm_Pk", "Lbl_code", "Size_Desc", "OTS_INV"]

    missing_mask = df[REQUIRED_COLS].apply(lambda c: c.map(is_missing)).any(axis=1)
    missing_idx = df.index[missing_mask].tolist()

    print(f"Filas totales: {len(df)}")
    print(f"Filas con datos faltantes (rojo): {len(missing_idx)}")

    # Build image maps (same as HTML gallery)
    image_map = extract_excel_images()
    pdf_map = extract_pdf_images()
    for _k, _v in extract_pdf_images2().items():
        pdf_map.setdefault(_k, _v)
    xls_map = extract_xls_images()
    ats_map = extract_ats_images()

    df[IMAGE_COL] = df.apply(lambda r: combined_image(r, image_map, pdf_map, xls_map, ats_map), axis=1)

    out_df = df.copy()

    out_df.to_excel(EXCEL_OUT, sheet_name=SHEET_NAME, index=False, engine="openpyxl")
    print(f"Excel escrito: {EXCEL_OUT}")

    # Post-process: red rows + hyperlinks for image column
    wb = load_workbook(EXCEL_OUT)
    ws = wb[SHEET_NAME]
    header = [c.value for c in ws[1]]
    col_index = {name: i + 1 for i, name in enumerate(header)}

    red_font = Font(color="FF0000", bold=True)
    red_fill = PatternFill("solid", fgColor="FFC7CE")

    for idx in missing_idx:
        row_num = idx + 2
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_num, column=c)
            cell.font = red_font
            cell.fill = red_fill

    if IMAGE_COL in col_index:
        cidx = col_index[IMAGE_COL]
        for r in range(2, ws.max_row + 1):
            cell = ws.cell(row=r, column=cidx)
            val = cell.value
            if val:
                rel = str(val)
                cell.hyperlink = rel
                cell.value = val

    for c in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(c)].width = 22

    wb.save(EXCEL_OUT)
    print("Post-procesado listo (filas rojas + hyperlinks).")


if __name__ == "__main__":
    main()