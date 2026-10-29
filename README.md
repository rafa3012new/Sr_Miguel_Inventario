# Image Generator API - Image Gallery

This project generates a static HTML image gallery from inventory Excel files. It maps product styles to images and provides a filterable, sortable table view.

## Features

- Static gallery (`index.html`) with a responsive table.
- Per-column checkbox filters with cross-filtering (faceted search).
- Sortable columns.
- Image mapping by style number from the `imagenes/` folder (brand + style first, then style).
- Column highlighting for actively filtered columns.
- Reset filters button.

## Project Structure

- `make_gallery.py` - Generates `index.html` and `Inventario_con_imagenes.xlsx` (single entry point).
- `inventory.py` - Reads the sources, maps images from `imagenes/`, and builds the inventory table and the Excel output.
- `index.html` - Generated gallery (single entry point).
- `sources/` - Input Excel/PDF files and additional images.
- `imagenes/` - Único carpeta de imagenes (todas las fuentes de imagen se consolidan aquí).
- `requirements.txt` - Python dependencies.

## Requirements

- Python 3.12 (recommended) with `pandas`, `openpyxl`, etc.
- Input files in `sources/` (e.g. `Inventario_con_imagenes.xlsx`, `Inventario-1.xlsx`).

## Installation (Python)

```bash
python -m pip install -r requirements.txt
```

## Generate the Gallery (Python)

```bash
python make_gallery.py
```

This regenerates `index.html` and `Inventario_con_imagenes.xlsx` from the files in `sources/`. Both are outputs; neither is needed as input.

## View the Gallery

Open `index.html` in your browser.

## Publish on GitHub Pages

1. Generate `index.html`.
2. Push the repository to GitHub.
3. Enable GitHub Pages (branch: `main`, folder: `/root`).

The site will be served from `index.html` automatically.

## Notes

- Generated caches are ignored by Git (see `.gitignore`).
- If you add new images to `imagenes/`, regenerate the gallery to include them.
- The Python generator writes directly to `index.html` (no separate `galeria.html`).

---

See also: [README.es.md](README.es.md) (Spanish).