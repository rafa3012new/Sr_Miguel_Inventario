import html
import re
import pandas as pd
from pathlib import Path
from openpyxl import load_workbook

try:
    import pymupdf
except Exception:
    pymupdf = None

BASE = Path(__file__).resolve().parent
IMAGE_COL = "Imagen"
EXCEL_IMG_SRC = "NAUTICA ATS TOM .xlsx"
EXCEL_IMG_DIR = "imagenes_excel"
PDF_IMG_SRC = "MMX FALL ATS 6.23 LEVY GROUP 3.pdf"
PDF_IMG_SRC2 = "LAUNDRY EARLY BUY LINESHEET 3.18.pdf"
PDF_IMG_DIR = "imagenes_pdf"
XLS_IMG_DIR = "imagenes_xls"


def safe_name(s):
    s = re.sub(r'[\\/*?:"<>|]', "_", str(s))
    s = re.sub(r"\s+", "_", s).strip("_")
    return (s[:50] or "item")


def extract_excel_images():
    src = BASE / EXCEL_IMG_SRC
    out_dir = BASE / EXCEL_IMG_DIR
    out_dir.mkdir(exist_ok=True)
    mapping = {}
    if not src.exists():
        return mapping
    wb = load_workbook(src)
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
    src = BASE / PDF_IMG_SRC
    out_dir = BASE / PDF_IMG_DIR
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
    src = BASE / PDF_IMG_SRC2
    out_dir = BASE / PDF_IMG_DIR
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
    out_dir = BASE / XLS_IMG_DIR
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

image_map = extract_excel_images()
pdf_map = extract_pdf_images()
# Agregar imagenes del nuevo PDF sin sobrescribir las existentes
for _k, _v in extract_pdf_images2().items():
    pdf_map.setdefault(_k, _v)

xls_map = extract_xls_images()

df = pd.read_excel(BASE / "Inventario_con_imagenes.xlsx", sheet_name="Data")

# Columna 'hoja': registros originales (Inventario) = sheet2, nuevos (append de Inventario-1) = sheet1
ORIGINAL_ROWS = 4924
df["hoja"] = (["sheet2"] * min(ORIGINAL_ROWS, len(df))
              + ["sheet1"] * max(0, len(df) - ORIGINAL_ROWS))

# Columna "Color_Name" mapeada desde Inventario-1.xlsx por Style + Color_code
try:
    _inv1 = pd.read_excel(BASE / "Inventario-1.xlsx", sheet_name="Data")
    _cn = {}
    for _, _r in _inv1.iterrows():
        _k = (str(_r.get("Style", "")).strip(), str(_r.get("Color_code", "")).strip().upper())
        _v = _r.get("Color Name")
        if _k not in _cn and _v is not None and str(_v).strip() != "" and str(_v).strip().lower() != "nan":
            _cn[_k] = str(_v).strip()
    df["Color_Name"] = df.apply(
        lambda r: _cn.get((str(r.get("Style", "")).strip(),
                           str(r.get("Color_code", "")).strip().upper()), ""),
        axis=1,
    )
except Exception:
    df["Color_Name"] = ""



def combined_image(r):
    excel_fname = image_map.get(
        (str(r.get("TLGDIV_NAME", "")).strip(), str(r.get("Style", "")).strip()), ""
    )
    if excel_fname:
        return f"{EXCEL_IMG_DIR}/{excel_fname}"
    pdf_fname = pdf_map.get(str(r.get("Style", "")).strip(), "")
    if pdf_fname:
        return f"{PDF_IMG_DIR}/{pdf_fname}"
    xls_fname = xls_map.get(str(r.get("Style", "")).strip(), "")
    if xls_fname:
        return xls_fname
    return ""


df[IMAGE_COL] = df.apply(combined_image, axis=1)

# Ordenar por OTS_INV ascendente
df = df.sort_values("OTS_INV", kind="mergesort", na_position="last").reset_index(drop=True)


cols = ["#", "TLGDIV_NAME", "Style", "Style_Name", "Color_code", "Color_Name", "Size_Desc", "OTS_INV", "hoja", IMAGE_COL]


def is_red(row):
    return (str(row.get("TLGDIV_NAME", "")).strip().lower() in ("", "(blank)", "nan")
            or str(row.get("Style", "")).strip().lower() in ("", "(blank)", "nan"))


def unique_vals(col):
    vals = []
    seen = set()
    for v in df[col].tolist():
        if v is None or (isinstance(v, float) and pd.isna(v)):
            s = ""
        elif col == "OTS_INV":
            try:
                f = float(v)
                s = str(int(f)) if f.is_integer() else str(f)
            except (ValueError, TypeError):
                s = str(v).strip()
        elif col == IMAGE_COL:
            s = "con" if (v is not None and str(v).strip() != "" and not (isinstance(v, float) and pd.isna(v))) else "sin"
        else:
            s = str(v).strip()
        if s not in seen:
            seen.add(s)
            vals.append(s)
    return vals


dropdowns = {}
for c in cols:
    if c == "#":
        dropdowns[c] = ""
        continue
    if c == IMAGE_COL:
        vals = ["con", "sin"]
    else:
        vals = unique_vals(c)
    boxes = []
    boxes.append(f'<input type="text" class="filtersearch" placeholder="Buscar...">')
    boxes.append(f'<label class="fitem"><input type="checkbox" class="colchk" data-col="{c}" value="__ALL__" checked> (Todos)</label>')
    for v in vals:
        esc = html.escape(str(v), quote=True)
        boxes.append(f'<label class="fitem"><input type="checkbox" class="colchk" data-col="{c}" value="{esc}" checked> {html.escape(str(v))}</label>')
    dropdowns[c] = "\n".join(boxes)

rows_html = []
for i, (_, r) in enumerate(df.iterrows(), start=1):
    red = is_red(r)
    tr_class = ' class="red"' if red else ""
    cells = []
    data_attrs = []
    for c in cols:
        if c == "#":
            val = i
            s = str(i)
        else:
            val = r.get(c)
            if val is None or (isinstance(val, float) and pd.isna(val)):
                s = ""
            elif c == "OTS_INV":
                try:
                    f = float(val)
                    s = str(int(f)) if f.is_integer() else str(f)
                except (ValueError, TypeError):
                    s = str(val).strip()
            elif c == IMAGE_COL:
                s = "con" if (val is not None and str(val).strip() != "" and not (isinstance(val, float) and pd.isna(val))) else "sin"
            else:
                s = str(val).strip()
        data_attrs.append(f'data-{c}="{html.escape(s, quote=True)}"')
    for c in cols:
        if c == "#":
            cells.append(f"<td>{i}</td>")
            continue
        val = r.get(c)
        if c == IMAGE_COL and isinstance(val, str) and val.strip():
            _v = html.escape(str(val), quote=True)
            img = f'<img src="{_v}" loading="lazy" alt="{html.escape(str(val))}" class="thumb" data-full="{_v}" onclick="openModal(this.dataset.full)" onerror="this.style.display=\'none\'">'
            cells.append(f"<td>{img}<br><small>{html.escape(str(val))}</small></td>")
        else:
            if val is None or (isinstance(val, float) and pd.isna(val)):
                v = ""
            elif c == "OTS_INV":
                try:
                    f = float(val)
                    v = str(int(f)) if f.is_integer() else str(f)
                except (ValueError, TypeError):
                    v = str(val)
            else:
                v = str(val)
            cells.append(f"<td>{html.escape(str(v))}</td>")
    rows_html.append(f"<tr{tr_class} {' '.join(data_attrs)}>{''.join(cells)}</tr>")

header_cells = []
for c in cols:
    if c == "#":
        header_cells.append(f"<th>#</th>")
    else:
        header_cells.append(
            f"<th>{c} <span class='sortbtn' data-col='{c}' onclick='sortTable(this)' style='cursor:pointer'>↕</span><br><button class='filterbtn' data-col='{c}'>Filtrar</button><div class='filterpanel' data-col='{c}' style='display:none'>{dropdowns[c]}</div></th>"
        )

html_doc = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>Galeria de imagenes - Inventario</title>
<style>
body {{ font-family: Arial, sans-serif; margin: 20px; }}
table {{ border-collapse: collapse; width: 100%; }}
th, td {{ border: 1px solid #ccc; padding: 6px; vertical-align: top; font-size: 12px; }}
th {{ background: #333; color: #fff; position: sticky; top: 0; }}
tr.red {{ background: #ffcccc; }}
img {{ max-width: 120px; max-height: 120px; object-fit: contain; }}
.sortbtn {{ cursor: pointer; font-size: 12px; color: #fff; user-select: none; }}
th {{ position: relative; }}
.filterbtn {{ font-size: 11px; padding: 2px 6px; cursor: pointer; }}
.filterpanel {{ position: absolute; top: 100%; left: 0; z-index: 1000; background: #fff; color: #000; border: 1px solid #ccc; padding: 6px; max-height: 260px; overflow: auto; text-align: left; min-width: 160px; box-shadow: 0 2px 8px rgba(0,0,0,0.15); }}
.fitem {{ display: block; white-space: nowrap; font-size: 11px; color: #000; }}
.filtersearch {{ width: 95%; margin-bottom: 4px; font-size: 11px; }}
#search {{ margin-bottom: 10px; padding: 8px; width: 300px; }}
select.colfilter {{ width: 100%; font-size: 11px; }}
.thumb {{ cursor: zoom-in; }}
.modal {{ display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.85); z-index: 9999; align-items: center; justify-content: center; flex-direction: column; padding: 20px; }}
.modal.open {{ display: flex; }}
.modal img {{ max-width: 90vw; max-height: 75vh; transition: transform .15s ease; transform-origin: center center; }}
.modal-close {{ position: absolute; top: 12px; right: 20px; color: #fff; font-size: 32px; cursor: pointer; line-height: 1; }}
.modal-controls {{ margin-top: 12px; }}
.modal-controls button {{ margin: 0 4px; padding: 6px 14px; font-size: 13px; cursor: pointer; }}
</style>
</head>
<body>
<h2>Galeria de imagenes - Inventario ({len(df)} filas) <span id="rowCount" style="font-size:13px; font-weight:normal;">(mostrando {len(df)})</span></h2>
<input type="text" id="search" placeholder="Buscar texto en toda la fila...">
<table id="tbl">
<thead><tr>{''.join(header_cells)}</tr></thead>
<tbody>
{''.join(rows_html)}
</tbody>
</table>
<div id="modal" class="modal" onclick="if(event.target===this) closeModal()">
  <span class="modal-close" onclick="closeModal()">&times;</span>
  <img id="modalImg" src="" alt="">
  <div class="modal-controls">
    <button onclick="setZoom(1)">1x</button>
    <button onclick="setZoom(2)">2x</button>
    <button onclick="setZoom(3)">3x</button>
  </div>
</div>
<script>
const COLS = ["#", "TLGDIV_NAME", "Style", "Style_Name", "Color_code", "Color_Name", "Size_Desc", "OTS_INV", "hoja", "Imagen"];
const input=document.getElementById('search');
function applyFilters(){{
  const q=input.value.toLowerCase();
  const filters={{}};
  document.querySelectorAll('.filterpanel').forEach(panel=>{{
    const col = panel.dataset.col;
    const boxes = panel.querySelectorAll('.colchk');
    const checked = [...boxes].filter(b=>b.checked && b.value !== '__ALL__').map(b=>b.value);
    const allBox = panel.querySelector('.colchk[value="__ALL__"]');
    if (allBox && allBox.checked) {{
      // todos seleccionados -> sin filtro
    }} else if (checked.length === 0) {{
      filters[col] = new Set(); // nada seleccionado -> ocultar todo
    }} else {{
      filters[col] = new Set(checked);
    }}
  }});
  document.querySelectorAll('#tbl tbody tr').forEach(tr=>{{
    let show = tr.textContent.toLowerCase().includes(q);
    for(const col in filters){{
      const val = tr.getAttribute('data-'+col);
      if (filters[col] instanceof Set) {{
        if (filters[col].size === 0 || !filters[col].has(val)) {{ show=false; break; }}
      }} else if (filters[col] && val !== filters[col]) {{ show=false; break; }}
    }}
    tr.style.display = show ? '' : 'none';
  }});
  const visible = Array.from(document.querySelectorAll('#tbl tbody tr')).filter(tr=>tr.style.display!=='none').length;
  const rc = document.getElementById('rowCount');
  if(rc) rc.textContent = '(mostrando ' + visible + ' de {len(df)})';
}}
input.addEventListener('input', applyFilters);

let sortState = {{ col: null, asc: true }};
function sortTable(btn){{
  const col = btn.dataset.col;
  const asc = (sortState.col === col) ? !sortState.asc : true;
  sortState = {{ col: col, asc: asc }};
  const tbody = document.querySelector('#tbl tbody');
  const rows = Array.from(tbody.querySelectorAll('tr'));
  const idx = COLS.indexOf(col);
  if (idx < 0) return;
  rows.sort(function(a,b){{
    const va = (a.cells[idx] ? a.cells[idx].textContent : '').trim();
    const vb = (b.cells[idx] ? b.cells[idx].textContent : '').trim();
    const na = parseFloat(va.replace(/,/g,''));
    const nb = parseFloat(vb.replace(/,/g,''));
    if (!isNaN(na) && !isNaN(nb)) return asc ? (na - nb) : (nb - na);
    return asc ? va.localeCompare(vb) : vb.localeCompare(va);
  }});
  rows.forEach(function(r){{ tbody.appendChild(r); }});
  document.querySelectorAll('.sortbtn').forEach(function(b){{ b.textContent = '\u2195'; }});
  btn.textContent = asc ? '\u25B2' : '\u25BC';
  applyFilters();
}}


// Toggle de paneles de filtro
document.addEventListener('click', function(e){{
  if (e.target && e.target.classList && e.target.classList.contains('filterbtn')){{
    const col = e.target.dataset.col;
    const panel = document.querySelector('.filterpanel[data-col="'+col+'"]');
    if (panel) panel.style.display = panel.style.display === 'none' ? 'block' : 'none';
  }} else if (!e.target.closest('.filterpanel') && !(e.target.classList && e.target.classList.contains('filterbtn'))){{
    document.querySelectorAll('.filterpanel').forEach(function(p){{ p.style.display='none'; }});
  }}
}});
// Select-all y cambio por checkbox
document.addEventListener('change', function(e){{
  if (e.target && e.target.classList && e.target.classList.contains('colchk') && e.target.value === '__ALL__'){{
    const col = e.target.dataset.col;
    const panel = document.querySelector('.filterpanel[data-col="'+col+'"]');
    if (panel) panel.querySelectorAll('.colchk').forEach(function(b){{ b.checked = e.target.checked; }});
    applyFilters();
  }} else if (e.target && e.target.classList && e.target.classList.contains('colchk')){{
    const col = e.target.dataset.col;
    const panel = document.querySelector('.filterpanel[data-col="'+col+'"]');
    if (panel){{
      const allBox = panel.querySelector('.colchk[value="__ALL__"]');
      const boxes = [...panel.querySelectorAll('.colchk')].filter(function(b){{ return b.value !== '__ALL__'; }});
      if (allBox) allBox.checked = boxes.every(function(b){{ return b.checked; }});
    }}
    applyFilters();
  }}
}});
// Busqueda dentro del panel
document.addEventListener('input', function(e){{
  if (e.target && e.target.classList && e.target.classList.contains('filtersearch')){{
    const panel = e.target.closest('.filterpanel');
    const q = e.target.value.toLowerCase();
    panel.querySelectorAll('.fitem').forEach(function(lbl){{
      const txt = lbl.textContent.toLowerCase();
      lbl.style.display = txt.includes(q) ? '' : 'none';
    }});
  }}
}});

</script>
</body>
</html>
"""

out = BASE / "galeria.html"
out.write_text(html_doc, encoding="utf-8")
(BASE / "index.html").write_text(
    '<!DOCTYPE html><html lang="es"><head><meta charset="utf-8">'
    '<meta http-equiv="refresh" content="0; url=galeria.html">'
    '<title>Galeria de imagenes - Inventario</title></head>'
    '<body><p>Redirigiendo a la <a href="galeria.html">galeria</a>...</p></body></html>',
    encoding="utf-8",
)
print("Generado:", out)