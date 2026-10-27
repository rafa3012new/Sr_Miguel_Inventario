import html
import pandas as pd
from pathlib import Path
from openpyxl import load_workbook

BASE = Path(__file__).resolve().parent
IMAGE_COL = "Imagen"
EXCEL_IMG_SRC = "NAUTICA ATS TOM .xlsx"
EXCEL_IMG_DIR = "imagenes_excel"


def safe_name(s):
    import re
    s = re.sub(r'[\\\\/*?:"<>|]', "_", str(s))
    s = re.sub(r"\s+", "_", s).strip("_")
    return (s[:50] or "item")


def extract_excel_images():
    """Extract embedded images from the Excel and map (TLGDIV, Style) -> filename."""
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


image_map = extract_excel_images()

df = pd.read_excel(BASE / "Inventario_con_imagenes.xlsx", sheet_name="Data")
df[IMAGE_COL] = df.apply(
    lambda r: image_map.get(
        (str(r.get("TLGDIV_NAME", "")).strip(), str(r.get("Style", "")).strip()), ""
    ),
    axis=1,
)
cols = ["#", "TLGDIV_NAME", "Style", "Style_Name", "Color_code", "Size_Desc", "OTS_INV", IMAGE_COL]


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


filter_panels = {}
for c in cols:
    if c == "#":
        filter_panels[c] = ""
        continue
    vals = unique_vals(c)
    cbs = []
    for v in vals:
        ev = html.escape(v, quote=True)
        cbs.append(f'<label class="fp-item"><input type="checkbox" class="chk-val" data-col="{c}" value="{ev}" checked> {html.escape(v)}</label>')
    panel = (
        f'<div class="filter-wrap">'
        f'<button class="filter-btn" data-col="{c}" onclick="toggleFilter(\'{c}\')">&#9660;</button>'
        f'<div class="filter-panel" id="fp-{c}" data-col="{c}" style="display:none;">'
        f'<label class="fp-item"><input type="checkbox" class="chk-all" data-col="{c}" checked> (Todos)</label>'
        f'<div class="fp-list">{"".join(cbs)}</div>'
        f'<div class="fp-actions"><button onclick="applyFilters()">Aplicar</button>'
        f'<button onclick="clearFilter(\'{c}\')">Limpiar</button></div>'
        f'</div></div>'
    )
    filter_panels[c] = panel

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
            else:
                s = str(val).strip()
            if c == IMAGE_COL:
                s = "con" if s else "sin"
        data_attrs.append(f'data-{c}="{html.escape(s, quote=True)}"')
    for c in cols:
        if c == "#":
            cells.append(f"<td>{i}</td>")
            continue
        val = r.get(c)
        if c == IMAGE_COL and isinstance(val, str) and val.strip():
            _v = html.escape(str(val), quote=True)
            img = f'<img src="{EXCEL_IMG_DIR}/{_v}" loading="lazy" alt="{html.escape(str(val))}" class="thumb" data-full="{EXCEL_IMG_DIR}/{_v}" onclick="openModal(this.dataset.full)" onerror="this.style.display=\'none\'">'
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
        header_cells.append(f"<th>{c}<br>{filter_panels[c]}</th>")

html = f"""<!DOCTYPE html>
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
#search {{ margin-bottom: 10px; padding: 8px; width: 300px; }}
.thumb {{ cursor: zoom-in; }}
.filter-wrap {{ position: relative; display: inline-block; }}
.filter-btn {{ cursor: pointer; border: 1px solid #888; background: #fff; font-size: 11px; padding: 2px 6px; border-radius: 3px; }}
.filter-panel {{ position: absolute; top: 100%; left: 0; background: #fff; color: #000; border: 1px solid #ccc; padding: 8px; max-height: 280px; overflow: auto; z-index: 100; min-width: 200px; box-shadow: 0 2px 8px rgba(0,0,0,.25); }}
.fp-item {{ display: block; white-space: nowrap; font-size: 12px; margin: 2px 0; }}
.fp-list {{ max-height: 190px; overflow: auto; margin: 6px 0; border-top: 1px solid #eee; border-bottom: 1px solid #eee; padding: 4px 0; }}
.fp-actions {{ display: flex; gap: 6px; margin-top: 6px; }}
.fp-actions button {{ font-size: 11px; cursor: pointer; }}
.modal {{ display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.85); z-index: 9999; align-items: center; justify-content: center; flex-direction: column; padding: 20px; }}
.modal.open {{ display: flex; }}
.modal img {{ max-width: 90vw; max-height: 75vh; transition: transform .15s ease; transform-origin: center center; }}
.modal-close {{ position: absolute; top: 12px; right: 20px; color: #fff; font-size: 32px; cursor: pointer; line-height: 1; }}
.modal-controls {{ margin-top: 12px; }}
.modal-controls button {{ margin: 0 4px; padding: 6px 14px; font-size: 13px; cursor: pointer; }}
</style>
</head>
<body>
<h2>Galeria de imagenes - Inventario ({len(df)} filas)</h2>
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
const input=document.getElementById('search');
function toggleFilter(col){{
  document.querySelectorAll('.filter-panel').forEach(p=>{{ if(p.id!=='fp-'+col) p.style.display='none'; }});
  const p=document.getElementById('fp-'+col);
  if(p) p.style.display = p.style.display==='none'?'block':'none';
}}
function clearFilter(col){{
  document.querySelectorAll('.chk-val[data-col="'+col+'"]').forEach(v=>v.checked=true);
  const all=document.querySelector('.chk-all[data-col="'+col+'"]');
  if(all) all.checked=true;
  applyFilters();
}}
document.querySelectorAll('.chk-val').forEach(cb=>{{
  cb.addEventListener('change', function(){{
    const col=this.dataset.col;
    const all=document.querySelector('.chk-all[data-col="'+col+'"]');
    const vals=document.querySelectorAll('.chk-val[data-col="'+col+'"]');
    if(all) all.checked = Array.from(vals).every(v=>v.checked);
  }});
}});
document.querySelectorAll('.chk-all').forEach(cb=>{{
  cb.addEventListener('change', function(){{
    const col=this.dataset.col;
    const checked=this.checked;
    document.querySelectorAll('.chk-val[data-col="'+col+'"]').forEach(v=>v.checked=checked);
  }});
}});
document.addEventListener('click', function(e){{
  if(!e.target.closest('.filter-wrap')){{
    document.querySelectorAll('.filter-panel').forEach(p=>p.style.display='none');
  }}
}});
function applyFilters(){{
  const q=input.value.toLowerCase();
  const colSelected={{}};
  document.querySelectorAll('.filter-panel').forEach(panel=>{{
    const col=panel.dataset.col;
    const vals=panel.querySelectorAll('.chk-val');
    const checked=Array.from(vals).filter(v=>v.checked).map(v=>v.value);
    if(checked.length>0 && checked.length < vals.length){{
      colSelected[col]=new Set(checked);
    }}
  }});
  document.querySelectorAll('#tbl tbody tr').forEach(tr=>{{
    let show = tr.textContent.toLowerCase().includes(q);
    for(const col in colSelected){{
      const val = tr.getAttribute('data-'+col);
      if(!colSelected[col].has(val)){{ show=false; break; }}
    }}
    tr.style.display = show ? '' : 'none';
  }});
}}
input.addEventListener('input', applyFilters);
function openModal(src){{
  const m=document.getElementById('modal');
  const img=document.getElementById('modalImg');
  img.src=src;
  img.style.transform='scale(1)';
  m.classList.add('open');
}}
function closeModal(){{
  const m=document.getElementById('modal');
  const img=document.getElementById('modalImg');
  m.classList.remove('open');
  img.src='';
  img.style.transform='scale(1)';
}}
function setZoom(z){{
  document.getElementById('modalImg').style.transform='scale('+z+')';
}}
document.addEventListener('keydown',function(e){{ if(e.key==='Escape') closeModal(); }});
</script>
</body>
</html>
"""
out = BASE / "galeria.html"
out.write_text(html, encoding="utf-8")
(BASE / "index.html").write_text(
    '<!DOCTYPE html><html lang="es"><head><meta charset="utf-8">'
    '<meta http-equiv="refresh" content="0; url=galeria.html">'
    '<title>Galeria de imagenes - Inventario</title></head>'
    '<body><p>Redirigiendo a la <a href="galeria.html">galeria</a>...</p></body></html>',
    encoding="utf-8",
)
print("Generado:", out)