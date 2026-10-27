import html
import pandas as pd
from pathlib import Path

BASE = Path(__file__).resolve().parent
df = pd.read_excel(BASE / "Inventario_con_imagenes.xlsx", sheet_name="Data")
cols = ["#", "TLGDIV_NAME", "Style", "Style_Name", "Color_code", "Size_Desc", "OTS_INV", "imagen_1", "imagen_2", "imagen_3"]


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
        else:
            s = str(v).strip()
        if s not in seen:
            seen.add(s)
            vals.append(s)
    return vals


dropdowns = {}
for c in cols:
    if c == "#":
        dropdowns[c] = '<option value="">Todos</option>'
        continue
    opts = ['<option value="">Todos</option>']
    if c.startswith("imagen_"):
        opts += ['<option value="con">Con imagen</option>', '<option value="sin">Sin imagen</option>']
    else:
        for v in unique_vals(c):
            opts.append(f'<option value="{html.escape(v, quote=True)}">{html.escape(v)}</option>')
    dropdowns[c] = "\n".join(opts)

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
            if c.startswith("imagen_"):
                s = "con" if s else "sin"
        data_attrs.append(f'data-{c}="{html.escape(s, quote=True)}"')
    for c in cols:
        if c == "#":
            cells.append(f"<td>{i}</td>")
            continue
        val = r.get(c)
        if c.startswith("imagen_") and isinstance(val, str) and val.strip():
            _v = html.escape(str(val), quote=True)
            img = f'<img src="imagenes/{_v}" loading="lazy" alt="{html.escape(str(val))}" class="thumb" data-full="imagenes/{_v}" onclick="openModal(this.dataset.full)" onerror="this.style.display=\'none\'">'
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
            f"<th>{c}<br><select class='colfilter' data-col='{c}'>{dropdowns[c]}</select></th>"
        )

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
function applyFilters(){{
  const q=input.value.toLowerCase();
  const filters={{}};
  document.querySelectorAll('select.colfilter').forEach(s=>{{
    if(s.value) filters[s.dataset.col]=s.value;
  }});
  document.querySelectorAll('#tbl tbody tr').forEach(tr=>{{
    let show = tr.textContent.toLowerCase().includes(q);
    for(const col in filters){{
      if(tr.getAttribute('data-'+col) !== filters[col]){{ show=false; break; }}
    }}
    tr.style.display = show ? '' : 'none';
  }});
}}
input.addEventListener('input', applyFilters);
document.querySelectorAll('select.colfilter').forEach(s=>s.addEventListener('change', applyFilters));
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
print("Generado:", out)
