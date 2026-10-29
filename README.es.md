# Image Generator API - Galeria de imagenes

Este proyecto genera una galeria HTML estatica a partir de archivos Excel de inventario. Mapea estilos de productos a imagenes y ofrece una tabla filtrable y ordenable.

## Caracteristicas

- Galeria estatica (`index.html`) con tabla responsive.
- Filtros por columna con checkboxes y filtro cruzado (faceted search).
- Columnas ordenables.
- Mapeo de imagenes por numero de estilo desde varias fuentes (Excel, PDF, XLS, ATS y `imagenes/`).
- Resaltado de columnas filtradas activamente.
- Boton para resetear filtros.

## Estructura del proyecto

- `make_gallery.py` - Genera `index.html` e `Inventario_con_imagenes.xlsx` (punto de entrada unico).
- `inventory.py` - Lee las fuentes y arma la tabla de inventario y el Excel de salida.
- `image_sources.py` - Reglas de mapeo de imagenes (Excel, PDF, XLS, ATS y `imagenes/`).
- `index.html` - Galeria generada (punto de entrada unico).
- `sources/` - Archivos Excel/PDF de entrada e imagenes adicionales.
- `imagenes/` - Unica carpeta de imagenes (todas las fuentes de imagen se consolidan aqui).
- `requirements.txt` - Dependencias de Python.

## Requisitos

- Python 3.12 (recomendado) con `pandas`, `openpyxl`, etc.
- Archivos de entrada en `sources/` (por ejemplo `Inventario_con_imagenes.xlsx`, `Inventario-1.xlsx`).

## Instalacion (Python)

```bash
python -m pip install -r requirements.txt
```

## Generar la galeria (Python)

```bash
python make_gallery.py
```

Esto regenera `index.html` e `Inventario_con_imagenes.xlsx` desde los archivos de `sources/`. Ambos son salidas; ninguno se necesita como entrada.

## Ver la galeria

Abre `index.html` en tu navegador.

## Publicar en GitHub Pages

1. Genera `index.html`.
2. Sube el repositorio a GitHub.
3. Activa GitHub Pages (rama: `main`, carpeta: `/root`).

El sitio se servira desde `index.html` automaticamente.

## Notas

- Caches generados estan ignorados por Git (ver `.gitignore`).
- Si agregas nuevas imagenes a `imagenes/`, regenera la galeria para incluirlas.
- El generador Python escribe directamente en `index.html` (sin `galeria.html` separado).

---

Ver tambien: [README.md](README.md) (Ingles).