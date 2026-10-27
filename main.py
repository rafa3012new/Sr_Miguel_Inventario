"""
Image Generator API - batch image fetcher for Inventario.xlsx

Usage:
    python main.py                # full run (needs API key in .env)
    python main.py --dry-run      # validate data + write Excel with red rows, no API
    python main.py --limit 20     # process only first 20 unique items (testing)
    python main.py --provider serpapi

Providers (set IMAGE_PROVIDER in .env or use --provider):
    bing      -> needs BING_API_KEY
    google    -> needs GOOGLE_API_KEY + GOOGLE_CSE_ID
    serpapi   -> needs SERPAPI_API_KEY
"""
import argparse
import asyncio
import json
import os
import re
from pathlib import Path

import httpx
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from tenacity import retry, stop_after_attempt, wait_exponential

BASE_DIR = Path(__file__).resolve().parent
EXCEL_IN = BASE_DIR / "Inventario.xlsx"
EXCEL_OUT = BASE_DIR / "Inventario_con_imagenes.xlsx"
IMG_DIR = BASE_DIR / "imagenes"
PROGRESS_FILE = BASE_DIR / "progress.json"
SHEET_NAME = "Data"
MAX_IMAGES = 3
CONCURRENCY = 5

REQUIRED_COLS = ["TLGDIV_NAME", "Style", "Style_Name", "Color_code"]
IMAGE_COLS = ["imagen_1", "imagen_2", "imagen_3"]


def load_env():
    env_path = BASE_DIR / ".env"
    if env_path.exists():
        for line in env_path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip())


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


def safe_name(s):
    s = re.sub(r'[\\/*?:"<>|]', "_", str(s))
    s = re.sub(r"\s+", "_", s).strip("_")
    return (s[:50] or "item")


def load_progress():
    if PROGRESS_FILE.exists():
        try:
            return json.loads(PROGRESS_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_progress(progress):
    PROGRESS_FILE.write_text(json.dumps(progress, ensure_ascii=False, indent=2), encoding="utf-8")


# ---------- Search providers ----------
async def search_duckduckgo(query, n):
    import asyncio as _a
    from ddgs import DDGS
    from ddgs.exceptions import RatelimitException

    last_err = None
    for attempt in range(4):
        if attempt > 0:
            await _a.sleep(5 * attempt + 5)
        try:
            def _run():
                out = []
                with DDGS() as ddgs:
                    for r in ddgs.images(query, max_results=n):
                        u = r.get("image")
                        if u:
                            out.append(u)
                            if len(out) >= n:
                                break
                return out[:n]
            return await _a.to_thread(_run)
        except RatelimitException as e:
            last_err = e
            continue
        except Exception as e:
            last_err = e
            # retry once on transient
            if attempt < 3:
                continue
            raise
    raise last_err if last_err else RuntimeError("duckduckgo failed")


async def search_bing(client, query, n, key):
    url = "https://api.bing.microsoft.com/v7.0/images/search"
    params = {"q": query, "count": n, "safeSearch": "Moderate"}
    headers = {"Ocp-Apim-Subscription-Key": key}
    r = await client.get(url, params=params, headers=headers, timeout=30)
    r.raise_for_status()
    data = r.json()
    return [it["contentUrl"] for it in data.get("value", [])][:n]


async def search_google(client, query, n, key, cx):
    url = "https://www.googleapis.com/customsearch/v1"
    params = {"q": query, "searchType": "image", "num": n, "key": key, "cx": cx}
    r = await client.get(url, params=params, timeout=30)
    r.raise_for_status()
    data = r.json()
    return [it["link"] for it in data.get("items", [])][:n]


async def search_serpapi(client, query, n, key):
    url = "https://serpapi.com/search.json"
    params = {"engine": "google_images", "q": query, "api_key": key}
    r = await client.get(url, params=params, timeout=30)
    r.raise_for_status()
    data = r.json()
    out = []
    for img in data.get("images_results", [])[:n]:
        u = img.get("original") or img.get("thumbnail")
        if u:
            out.append(u)
    return out[:n]


@retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=2, max=10))
async def do_search(client, provider, query, n):
    if provider == "duckduckgo" or provider == "ddg":
        return await search_duckduckgo(query, n)
    if provider == "bing":
        key = os.environ.get("BING_API_KEY")
        if not key:
            raise RuntimeError("Falta BING_API_KEY en .env")
        return await search_bing(client, query, n, key)
    if provider == "google":
        key = os.environ.get("GOOGLE_API_KEY")
        cx = os.environ.get("GOOGLE_CSE_ID")
        if not key or not cx:
            raise RuntimeError("Faltan GOOGLE_API_KEY / GOOGLE_CSE_ID en .env")
        return await search_google(client, query, n, key, cx)
    if provider == "serpapi":
        key = os.environ.get("SERPAPI_API_KEY")
        if not key:
            raise RuntimeError("Falta SERPAPI_API_KEY en .env")
        return await search_serpapi(client, query, n, key)
    raise RuntimeError(f"Proveedor desconocido: {provider}")


async def download_image(client, url, dest):
    r = await client.get(url, timeout=40, follow_redirects=True)
    r.raise_for_status()
    dest.write_bytes(r.content)
    # validate
    from PIL import Image
    with Image.open(dest) as im:
        im.verify()


async def handle_key(client, sem, provider, key, progress):
    key_str = "|".join(key)
    async with sem:
        if key_str in progress:
            return key, progress[key_str]

        base_query = f"{key[2]} {key[0]} {key[3]}"
        # Prioridad: Amazon (Nautica) -> sitio de Nautica -> busqueda general
        queries = [f"{base_query} site:amazon.com"]
        if "NAUTICA" in str(key[0]).upper():
            queries.append(f"{base_query} site:nautica.com")
            queries.append(f"{base_query} site:nautica.eu")
        queries.append(f"{base_query} amazon")
        queries.append(base_query)
        print(f"[SEARCH] {base_query}")
        urls = []
        for q in queries:
            try:
                urls = await do_search(client, provider, q, MAX_IMAGES)
            except Exception as e:
                print(f"  ! search error ({q}): {e}")
                urls = []
            if urls:
                break

        filenames = []
        for i, url in enumerate(urls):
            fname = f"{safe_name(key[0])}_{safe_name(key[1])}_{safe_name(key[3])}_{i+1}.jpg"
            dest = IMG_DIR / fname
            if dest.exists():
                filenames.append(fname)
                continue
            try:
                await download_image(client, url, dest)
                filenames.append(fname)
                print(f"  + downloaded {fname}")
            except Exception as e:
                print(f"  ! download error ({url}): {e}")

        progress[key_str] = filenames
        save_progress(progress)
        return key, filenames


async def main_async(args):
    load_env()
    provider = (args.provider or os.environ.get("IMAGE_PROVIDER", "duckduckgo")).lower()

    if not EXCEL_IN.exists():
        raise SystemExit(f"No se encontro {EXCEL_IN}")

    IMG_DIR.mkdir(exist_ok=True)

    # Read Excel
    df = pd.read_excel(EXCEL_IN, sheet_name=SHEET_NAME, header=2, usecols="A:H")
    df.columns = ["TLGDIV_NAME", "Style", "Style_Name", "Color_code",
                  "Dm_Pk", "Lbl_code", "Size_Desc", "OTS_INV"]

    # Missing mask over the 4 required cols
    missing_mask = df[REQUIRED_COLS].apply(lambda c: c.map(is_missing)).any(axis=1)
    missing_idx = df.index[missing_mask].tolist()
    valid_df = df[~missing_mask].copy()

    print(f"Filas totales: {len(df)}")
    print(f"Filas con datos faltantes (rojo): {len(missing_idx)}")
    print(f"Filas validas: {len(valid_df)}")

    # Dedup keys
    valid_df["_key"] = valid_df.apply(
        lambda r: (str(r["TLGDIV_NAME"]).strip(), str(r["Style"]).strip(),
                   str(r["Style_Name"]).strip(), str(r["Color_code"]).strip()),
        axis=1,
    )
    unique_keys = list(dict.fromkeys(valid_df["_key"].tolist()))
    if args.limit:
        unique_keys = unique_keys[: args.limit]
    print(f"Combinaciones unicas a procesar: {len(unique_keys)}")

    # Map key -> filenames
    key_to_files = {}
    progress = load_progress()

    if not args.dry_run:
        # seed from progress
        for k in unique_keys:
            ks = "|".join(k)
            if ks in progress:
                key_to_files[k] = progress[ks]

        to_process = [k for k in unique_keys if "|".join(k) not in progress]
        print(f"Ya en cache: {len(unique_keys) - len(to_process)} | Por procesar: {len(to_process)}")

        conc = 1 if provider in ("duckduckgo", "ddg") else CONCURRENCY
        sem = asyncio.Semaphore(conc)
        async with httpx.AsyncClient() as client:
            tasks = [handle_key(client, sem, provider, k, progress) for k in to_process]
            results = await asyncio.gather(*tasks, return_exceptions=True)
            for res in results:
                if isinstance(res, Exception):
                    print(f"  ! task error: {res}")
                    continue
                k, files = res
                key_to_files[k] = files
        # ensure cached ones present
        for k in unique_keys:
            key_to_files.setdefault(k, progress.get("|".join(k), []))
    else:
        print("DRY-RUN: no se llama a la API ni se descargan imagenes.")

    # Build image columns aligned to original df
    img1 = [""] * len(df)
    img2 = [""] * len(df)
    img3 = [""] * len(df)
    for idx, row in valid_df.iterrows():
        k = row["_key"]
        if k not in unique_keys and args.limit:
            # skipped by limit
            continue
        files = key_to_files.get(k, [])
        if len(files) > 0:
            img1[idx] = files[0]
        if len(files) > 1:
            img2[idx] = files[1]
        if len(files) > 2:
            img3[idx] = files[2]

    out_df = df.copy()
    out_df["imagen_1"] = img1
    out_df["imagen_2"] = img2
    out_df["imagen_3"] = img3
    from urllib.parse import quote
    def img_url(f):
        if not f:
            return ""
        abs_path = (IMG_DIR / f).resolve().as_posix()
        return "file:///" + quote(abs_path, safe="/:")
    # Columnas con formula para mostrar la imagen en la celda (Excel 365+) y link clickeable
    out_df["mostrar_1"] = [f'=IMAGE("{img_url(f)}")' if f else "" for f in img1]
    out_df["mostrar_2"] = [f'=IMAGE("{img_url(f)}")' if f else "" for f in img2]
    out_df["mostrar_3"] = [f'=IMAGE("{img_url(f)}")' if f else "" for f in img3]
    out_df["ver_1"] = [f'=HYPERLINK("imagenes/{f}","Ver imagen")' if f else "" for f in img1]
    out_df["ver_2"] = [f'=HYPERLINK("imagenes/{f}","Ver imagen")' if f else "" for f in img2]
    out_df["ver_3"] = [f'=HYPERLINK("imagenes/{f}","Ver imagen")' if f else "" for f in img3]

    # Write Excel
    out_df.to_excel(EXCEL_OUT, sheet_name=SHEET_NAME, index=False, engine="openpyxl")
    print(f"Excel escrito: {EXCEL_OUT}")

    # Post-process: red rows + hyperlinks
    wb = load_workbook(EXCEL_OUT)
    ws = wb[SHEET_NAME]
    header = [c.value for c in ws[1]]
    col_index = {name: i + 1 for i, name in enumerate(header)}

    red_font = Font(color="FF0000", bold=True)
    red_fill = PatternFill("solid", fgColor="FFC7CE")

    # map original df index -> excel row (header at row 1 => data row = idx+2)
    for idx in missing_idx:
        row_num = idx + 2
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_num, column=c)
            cell.font = red_font
            cell.fill = red_fill

    # hyperlinks for image cols
    for colname in IMAGE_COLS:
        if colname not in col_index:
            continue
        cidx = col_index[colname]
        for r in range(2, ws.max_row + 1):
            cell = ws.cell(row=r, column=cidx)
            val = cell.value
            if val:
                rel = f"imagenes/{val}"
                cell.hyperlink = rel
                cell.value = val

    # column widths
    for c in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(c)].width = 22

    wb.save(EXCEL_OUT)
    print("Post-procesado listo (filas rojas + hyperlinks).")
    print(f"Imagenes en: {IMG_DIR}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true", help="No API calls, solo valida y escribe Excel.")
    parser.add_argument("--limit", type=int, default=None, help="Limitar numero de items unicos (pruebas).")
    parser.add_argument("--provider", type=str, default=None, help="bing|google|serpapi")
    args = parser.parse_args()

    asyncio.run(main_async(args))


if __name__ == "__main__":
    main()
