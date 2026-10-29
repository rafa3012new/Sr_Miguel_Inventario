const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

const BASE = __dirname;
const SOURCES = path.join(BASE, "sources");
const EXCEL_IN = path.join(SOURCES, "Inventario.xlsx");
const EXCEL_IN1 = path.join(SOURCES, "Inventario-1.xlsx");
const EXCEL_IMG_SRC = path.join(SOURCES, "NAUTICA ATS TOM .xlsx");

const IMAGE_COL = "Imagen";
const EXCEL_IMG_DIR = "imagenes_excel";
const PDF_IMG_DIR = "imagenes_pdf";
const XLS_IMG_DIR = "imagenes_xls";
const ATS_IMG_DIR = "imagenes_ats";
const IMAGES_DIR = "imagenes";

const ORIGINAL_ROWS = 4924;

function safeName(s) {
  s = String(s).replace(/[\\/*?:"<>|]/g, "_");
  s = s.replace(/\s+/g, "_").replace(/^_+|_+$/g, "");
  return s.slice(0, 50) || "item";
}

function isMissing(v) {
  if (v === null || v === undefined) return true;
  const s = String(v).trim();
  return s === "" || ["nan", "(blank)", "(en blanco)", "none", "nat"].includes(s.toLowerCase());
}

function htmlEscape(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function ensureDir(d) { fs.mkdirSync(d, { recursive: true }); }

function scanExcelImages() {
  const map = {};
  const dir = path.join(BASE, EXCEL_IMG_DIR);
  if (!fs.existsSync(dir)) return map;
  for (const f of fs.readdirSync(dir)) {
    const full = path.join(dir, f);
    if (!fs.statSync(full).isFile()) continue;
    if (!f.startsWith("excel_") || !/\.(png|jpg|jpeg)$/i.test(f)) continue;
    const stem = f.slice(6).replace(/\.[^.]+$/, "");
    const m = stem.match(/(\d{7,})$/);
    if (m) {
      const style = m[1];
      if (!map[style]) map[style] = f;
    }
  }
  return map;
}

function scanPdfImages() {
  const map = {};
  const dir = path.join(BASE, PDF_IMG_DIR);
  if (!fs.existsSync(dir)) return map;
  for (const f of fs.readdirSync(dir)) {
    const full = path.join(dir, f);
    if (!fs.statSync(full).isFile()) continue;
    if (!f.startsWith("pdf_") || !/\.(png|jpg|jpeg)$/i.test(f)) continue;
    const stem = f.slice(4).replace(/\.[^.]+$/, "");
    // style may contain non-digits? original uses style as key (digits). Keep full stem.
    if (!map[stem]) map[stem] = f;
  }
  return map;
}

function scanXlsImages() {
  const map = {};
  const dir = path.join(BASE, XLS_IMG_DIR);
  if (!fs.existsSync(dir)) return map;
  for (const f of fs.readdirSync(dir)) {
    const full = path.join(dir, f);
    if (!fs.statSync(full).isFile()) continue;
    if (!f.startsWith("xls_") || !/\.(png|jpg|jpeg)$/i.test(f)) continue;
    const key = f.slice(4).replace(/\.[^.]+$/, "");
    map[key] = `${XLS_IMG_DIR}/${f}`;
  }
  if (map["V623040"]) {
    map["V623040A"] = map["V623040A"] || map["V623040"];
    map["V623040H"] = map["V623040H"] || map["V623040"];
  }
  return map;
}

function scanAtsImages() {
  const map = {};
  const dir = path.join(BASE, ATS_IMG_DIR);
  if (!fs.existsSync(dir)) return map;
  const byStyle = {};
  if (!fs.existsSync(dir)) return map;
  for (const f of fs.readdirSync(dir)) {
    const full = path.join(dir, f);
    if (!fs.statSync(full).isFile()) continue;
    if (!f.startsWith("ATS_") || !/\.(png|jpg|jpeg)$/i.test(f)) continue;
    const stem = f.slice(4).replace(/\.[^.]+$/, "");
    const parts = stem.split("_");
    if (parts.length < 2) continue;
    const style = parts[0];
    let idx = 0;
    if (parts.length >= 3 && /^\d+$/.test(parts[parts.length - 1])) {
      idx = parseInt(parts[parts.length - 1], 10);
    }
    byStyle[style] = byStyle[style] || [];
    byStyle[style].push([idx, f]);
  }
  for (const style of Object.keys(byStyle)) {
    byStyle[style].sort((a, b) => a[0] - b[0]);
    map[style] = `${ATS_IMG_DIR}/${byStyle[style][0][1]}`;
  }
  return map;
}

function scanImagenes() {
  const map = {};
  const dir = path.join(BASE, "imagenes");
  if (!fs.existsSync(dir)) return map;
  for (const f of fs.readdirSync(dir)) {
    const full = path.join(dir, f);
    if (!fs.statSync(full).isFile()) continue;
    if (!/\.(png|jpg|jpeg)$/i.test(f)) continue;
    const stem = f.replace(/\.[^.]+$/, "");
    if (!map[stem]) map[stem] = `imagenes/${f}`;
  }
  return map;
}

async function extractExcelImagesEmbedded() {
  const JSZip = require("jszip");
  const { XMLParser } = require("fast-xml-parser");
  const outDir = path.join(BASE, EXCEL_IMG_DIR);
  ensureDir(outDir);
  const map = {};
  if (!fs.existsSync(EXCEL_IMG_SRC)) return map;
  const zip = await JSZip.loadAsync(fs.readFileSync(EXCEL_IMG_SRC));
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });

  const wbXml = await zip.file("xl/workbook.xml").async("string");
  const wb = parser.parse(wbXml);
  const sheetsRaw = wb.workbook.sheets.sheet;
  const sheetArr = Array.isArray(sheetsRaw) ? sheetsRaw : [sheetsRaw];

  const wbRelsXml = await zip.file("xl/_rels/workbook.xml.rels").async("string");
  const wbRels = parser.parse(wbRelsXml);
  const wrelsRaw = wbRels.Relationships.Relationship;
  const wrels = Array.isArray(wrelsRaw) ? wrelsRaw : [wrelsRaw];
  const relMap = {};
  for (const r of wrels) relMap[r["@_Id"]] = r["@_Target"];

  let chosen = sheetArr.find((s) => s["@_name"] === "Sheet");
  if (!chosen && sheetArr.length > 1) chosen = sheetArr[1];
  if (!chosen) chosen = sheetArr[0];
  const sheetTarget = relMap[chosen["@_r:id"] || chosen["@_id"]];
  if (!sheetTarget) return map;
  const sheetPath = "xl/" + sheetTarget.replace(/^\//, "");

  const sheetXml = await zip.file(sheetPath).async("string");
  const sheetRelsPath = sheetPath.replace("xl/worksheets/", "xl/worksheets/_rels/") + ".rels";
  const sheetRelsXml = await zip.file(sheetRelsPath).async("string");
  const sheetRels = parser.parse(sheetRelsXml);
  const srelsRaw = sheetRels.Relationships.Relationship;
  const srels = Array.isArray(srelsRaw) ? srelsRaw : [srelsRaw];
  const srelMap = {};
  for (const r of srels) srelMap[r["@_Id"]] = r["@_Target"];

  const sheetObj = parser.parse(sheetXml);
  const drawing = sheetObj.worksheet ? sheetObj.worksheet.drawing : null;
  const drawingRid = drawing ? (drawing["@_r:id"] || drawing["@_id"]) : null;
  if (!drawingRid) return map;
  const drawingTarget = srelMap[drawingRid];
  if (!drawingTarget) return map;
  const drawingPath = path.posix.normalize(path.posix.join(path.posix.dirname(sheetPath), drawingTarget));
  const drawingXml = await zip.file(drawingPath).async("string");
  const drawingObj = parser.parse(drawingXml);

  const drawingRelsPath = drawingPath.replace("xl/drawings/", "xl/drawings/_rels/") + ".rels";
  const drawingRelsXml = await zip.file(drawingRelsPath).async("string");
  const drawingRels = parser.parse(drawingRelsXml);
  const drelsRaw = drawingRels.Relationships.Relationship;
  const drels = Array.isArray(drelsRaw) ? drelsRaw : [drelsRaw];
  const drelMap = {};
  for (const r of drels) drelMap[r["@_Id"]] = r["@_Target"];

  // anchors
  const wsDr = drawingObj["xdr:wsDr"] || drawingObj.wsDr || drawingObj;
  const anchors = [];
  for (const key of Object.keys(wsDr)) {
    if (key.toLowerCase().includes("anchor")) {
      const val = wsDr[key];
      const arr = Array.isArray(val) ? val : [val];
      for (const a of arr) anchors.push(a);
    }
  }

  const wb2 = XLSX.readFile(EXCEL_IMG_SRC);
  const ws2 = wb2.Sheets[chosen["@_name"]];
  const allRows = XLSX.utils.sheet_to_json(ws2, { header: 1, defval: null, blankrows: true });

  for (const anchor of anchors) {
    const from = anchor["xdr:from"] || anchor.from;
    const rowNode = from ? (from["xdr:row"] || from.row) : null;
    const excelRow = (parseInt(rowNode, 10) || 0) + 1;
    const pic = anchor["xdr:pic"] || anchor.pic;
    if (!pic) continue;
    const blipFill = pic["xdr:blipFill"] || pic.blipFill || pic;
    const blip = blipFill["a:blip"] || blipFill.blip;
    const embed = blip ? (blip["@_r:embed"] || blip["@_embed"]) : null;
    if (!embed) continue;
    const mediaTarget = drelMap[embed];
    if (!mediaTarget) continue;
    const mediaPath = path.posix.normalize(path.posix.join(path.posix.dirname(drawingPath), mediaTarget));
    const mediaFile = zip.file(mediaPath);
    if (!mediaFile) continue;
    const buf = await mediaFile.async("nodebuffer");

    const row = allRows[excelRow - 1] || [];
    const tl = row[0];
    const style = row[1];
    if (tl == null || style == null) continue;
    const tlS = String(tl).trim();
    const styleS = String(style).trim();
    if (/total/i.test(styleS)) continue;

    const ext = (path.extname(mediaPath).slice(1) || "png").toLowerCase();
    const fname = `excel_${safeName(tlS)}_${safeName(styleS)}.${ext}`;
    const dest = path.join(outDir, fname);
    if (!fs.existsSync(dest)) fs.writeFileSync(dest, buf);

    map[`${tlS}||${styleS}`] = fname;
    if (!map[styleS]) map[styleS] = fname;
  }
  return map;
}

function combinedImage(r, imageMap, pdfMap, xlsMap, atsMap, imagenesMap) {
  const style = String(r.Style ?? "").trim();
  const color = String(r.Color_code ?? "").trim().toUpperCase();
  // manual overrides
  if (style === "V322453M" && color === "STO") return "V322453M_COLOR_STONE.png";
  if (style === "119004414") return "imagenes/119004414.png";
  if (style === "117004093") return "imagenes/117004093.jpg";
  if (style === "119000214") return "imagenes/119000214.jpg";
  if (style === "119000455") return "imagenes/119000455.jpg";
  if (style === "119001566") return "imagenes/119001566.jpg";
  if (style === "119001328") return "imagenes/119001328.jpg";
  if (style === "119005385") return "imagenes/119005385.jpg";
  if (style === "121000428") return "imagenes/121000428.jpg";

  const tl = String(r.TLGDIV_NAME ?? "").trim();
  let excelFname = imageMap[`${tl}||${style}`];
  if (!excelFname) excelFname = imageMap[style];
  if (excelFname) return `${EXCEL_IMG_DIR}/${excelFname}`;

  const pdfFname = pdfMap[style];
  if (pdfFname) return `${PDF_IMG_DIR}/${pdfFname}`;

  const xlsFname = xlsMap[style];
  if (xlsFname) return xlsFname;

  const atsFname = atsMap[style];
  if (atsFname) return atsFname;

  // Imagenes de la carpeta \imagenes, solo si no hay otra ya mapeada.
  const imgFname = imagenesMap ? imagenesMap[style] : null;
  if (imgFname) return imgFname;

  return "";
}

async function main() {
  const args = process.argv.slice(2);
  const extractExcel = args.includes("--extract-excel") || args.includes("--extract-excel-images");

  // Read inventory
  if (!fs.existsSync(EXCEL_IN)) {
    console.error("No se encontro", EXCEL_IN);
    process.exit(1);
  }
  const wb = XLSX.readFile(EXCEL_IN);
  const ws = wb.Sheets["Data"];
  const rawRows = XLSX.utils.sheet_to_json(ws, { header: 1, range: 2, defval: "" });
  // rawRows[0] is header
  const header = rawRows[0] || [];
  const isExcludedRow = (r) => {
  const excluded = new Set(["", "(blank)", "nan", "(en blanco)", "total general"]);
  const tl = String(r[0] ?? "").trim().toLowerCase();
  const st = String(r[1] ?? "").trim().toLowerCase();
  return excluded.has(tl) || excluded.has(st);
};
const dataRows = rawRows.slice(1).filter((r) => r.some((v) => String(v ?? "").trim() !== "") && !isExcludedRow(r));

  const df = dataRows.map((r) => ({
    TLGDIV_NAME: r[0],
    Style: r[1],
    Style_Name: r[2],
    Color_code: r[3],
    Dm_Pk: r[4],
    Lbl_code: r[5],
    Size_Desc: r[6],
    OTS_INV: r[7],
  }));

  // Color_Name from Inventario-1.xlsx
  const colorNameMap = {};
  if (fs.existsSync(EXCEL_IN1)) {
    try {
      const wb1 = XLSX.readFile(EXCEL_IN1);
      const ws1 = wb1.Sheets["Data"];
      const rows1 = XLSX.utils.sheet_to_json(ws1, { header: 1, defval: "" });
      const h1 = rows1[0] || [];
      const idxStyle = h1.indexOf("Style");
      const idxColorCode = h1.indexOf("Color_code");
      const idxColorName = h1.indexOf("Color Name");
      for (const r of rows1.slice(1)) {
        if (!r || r.length === 0) continue;
        const k = [String(r[idxStyle] ?? "").trim(), String(r[idxColorCode] ?? "").trim().toUpperCase()];
        const v = r[idxColorName];
        if (k[0] && k[1] && v !== undefined && v !== null && String(v).trim() !== "" && String(v).trim().toLowerCase() !== "nan") {
          const key = `${k[0]}||${k[1]}`;
          if (!(key in colorNameMap)) colorNameMap[key] = String(v).trim();
        }
      }
    } catch (e) {
      console.error("Error leyendo Inventario-1:", e.message);
    }
  }

  for (const r of df) {
    const key = `${String(r.Style ?? "").trim()}||${String(r.Color_code ?? "").trim().toUpperCase()}`;
    r.Color_Name = colorNameMap[key] || "";
  }

  // image maps
  const cachePath = path.join(BASE, "image_maps.json");
  let excelMap = {};
  let pdfMap = {};
  let xlsMap = {};
  let atsMap = {};

  if (fs.existsSync(cachePath)) {
    try {
      const c = JSON.parse(fs.readFileSync(cachePath, "utf8"));
      excelMap = c.excel || {};
      pdfMap = c.pdf || {};
      xlsMap = c.xls || {};
      atsMap = c.ats || {};
    } catch (e) {}
  }

  // scan folders (always, to fill missing)
  const scannedExcel = scanExcelImages();
  const scannedPdf = scanPdfImages();
  const scannedXls = scanXlsImages();
  const scannedAts = scanAtsImages();
  const scannedImagenes = scanImagenes();

  if (Object.keys(excelMap).length === 0) excelMap = scannedExcel;
  else for (const k of Object.keys(scannedExcel)) if (!(k in excelMap)) excelMap[k] = scannedExcel[k];

  if (Object.keys(pdfMap).length === 0) pdfMap = scannedPdf;
  else for (const k of Object.keys(scannedPdf)) if (!(k in pdfMap)) pdfMap[k] = scannedPdf[k];

  if (Object.keys(xlsMap).length === 0) xlsMap = scannedXls;
  else for (const k of Object.keys(scannedXls)) if (!(k in xlsMap)) xlsMap[k] = scannedXls[k];

  if (Object.keys(atsMap).length === 0) atsMap = scannedAts;
  else for (const k of Object.keys(scannedAts)) if (!(k in atsMap)) atsMap[k] = scannedAts[k];

  if (extractExcel) {
    try {
      const extracted = await extractExcelImagesEmbedded();
      for (const k of Object.keys(extracted)) excelMap[k] = extracted[k];
      console.log("Extraidas imagenes de Excel:", Object.keys(extracted).length);
    } catch (e) {
      console.error("Error extrayendo imagenes de Excel:", e.message);
    }
  }

  // save cache
  try {
    fs.writeFileSync(cachePath, JSON.stringify({ excel: excelMap, pdf: pdfMap, xls: xlsMap, ats: atsMap }, null, 2));
  } catch (e) {}

  // combined image
  for (const r of df) {
    r[IMAGE_COL] = combinedImage(r, excelMap, pdfMap, xlsMap, atsMap, scannedImagenes);
  }

  // sort by provider (TLGDIV_NAME) asc, then by style number (Style) asc
  df.sort((a, b) => {
    const pa = (a.TLGDIV_NAME === null || a.TLGDIV_NAME === undefined) ? "" : String(a.TLGDIV_NAME).trim();
    const pb = (b.TLGDIV_NAME === null || b.TLGDIV_NAME === undefined) ? "" : String(b.TLGDIV_NAME).trim();
    const c1 = pa.localeCompare(pb);
    if (c1 !== 0) return c1;
    const sa = (a.Style === null || a.Style === undefined) ? "" : String(a.Style).trim();
    const sb = (b.Style === null || b.Style === undefined) ? "" : String(b.Style).trim();
    const na = Number(sa), nb = Number(sb);
    const aNum = !isNaN(na) && sa !== "";
    const bNum = !isNaN(nb) && sb !== "";
    if (aNum && bNum) return na - nb;
    return sa.localeCompare(sb);
  });

  const cols = ["#", "TLGDIV_NAME", "Style", "Style_Name", "Color_code", "Color_Name", "Size_Desc", "OTS_INV", IMAGE_COL];

  function isRed(row) {
    const tl = String(row.TLGDIV_NAME ?? "").trim().toLowerCase();
    const st = String(row.Style ?? "").trim().toLowerCase();
    return tl === "" || tl === "(blank)" || tl === "nan" || st === "" || st === "(blank)" || st === "nan";
  }

  function uniqueVals(col) {
    const vals = [];
    const seen = new Set();
    for (const r of df) {
      let s = "";
      if (col === "#") s = "";
      else if (col === "OTS_INV") {
        const v = r.OTS_INV;
        if (v === null || v === undefined || (typeof v === "number" && isNaN(v))) s = "";
        else {
          const f = Number(v);
          s = isNaN(f) ? String(v).trim() : (Number.isInteger(f) ? String(f) : String(f));
        }
      } else if (col === IMAGE_COL) {
        const v = r[IMAGE_COL];
        s = (v !== null && v !== undefined && String(v).trim() !== "" && !(typeof v === "number" && isNaN(v))) ? "con imagen" : "sin imagen";
      } else {
        const v = r[col];
        if (v === null || v === undefined || (typeof v === "number" && isNaN(v))) s = "";
        else s = String(v).trim();
      }
      if (!seen.has(s)) {
        seen.add(s);
        vals.push(s);
      }
    }
    return vals;
  }

  const dropdowns = {};
  for (const c of cols) {
    if (c === "#") {
      dropdowns[c] = "";
      continue;
    }
    if (c === IMAGE_COL) {
      const vals = ["con imagen", "sin imagen"];
      const boxes = [];
      boxes.push('<input type="text" class="filtersearch" placeholder="Buscar...">');
      boxes.push(`<label class="fitem"><input type="checkbox" class="colchk" data-col="${c}" value="__ALL__" checked> (Todos)</label>`);
      for (const v of vals) {
        const esc = htmlEscape(v);
        boxes.push(`<label class="fitem"><input type="checkbox" class="colchk" data-col="${c}" value="${esc}" checked> ${htmlEscape(v)}</label>`);
      }
      dropdowns[c] = boxes.join("\n");
      continue;
    }
    const vals = uniqueVals(c);
    const boxes = [];
    boxes.push('<input type="text" class="filtersearch" placeholder="Buscar...">');
    boxes.push(`<label class="fitem"><input type="checkbox" class="colchk" data-col="${c}" value="__ALL__" checked> (Todos)</label>`);
    for (const v of vals) {
      const esc = htmlEscape(v);
      boxes.push(`<label class="fitem"><input type="checkbox" class="colchk" data-col="${c}" value="${esc}" checked> ${htmlEscape(v)}</label>`);
    }
    dropdowns[c] = boxes.join("\n");
  }

  const rowsHtml = [];
  for (let i = 0; i < df.length; i++) {
    const r = df[i];
    const red = isRed(r);
    const trClass = red ? ' class="red"' : "";
    const dataAttrs = [];
    for (const c of cols) {
      if (c === "#") {
        dataAttrs.push(`data-#="${i + 1}"`);
        continue;
      }
      let val = r[c];
      let s = "";
      if (val === null || val === undefined || (typeof val === "number" && isNaN(val))) s = "";
      else if (c === "OTS_INV") {
        const f = Number(val);
        s = isNaN(f) ? String(val).trim() : (Number.isInteger(f) ? String(f) : String(f));
      } else if (c === IMAGE_COL) {
        s = (val !== null && val !== undefined && String(val).trim() !== "" && !(typeof val === "number" && isNaN(val))) ? "con imagen" : "sin imagen";
      } else {
        s = String(val).trim();
      }
      dataAttrs.push(`data-${c}="${htmlEscape(s)}"`);
    }
    const cells = [];
    for (const c of cols) {
      if (c === "#") {
        cells.push(`<td>${i + 1}</td>`);
        continue;
      }
      const val = r[c];
      if (c === IMAGE_COL && typeof val === "string" && val.trim()) {
        const v = htmlEscape(val);
        cells.push(`<td data-col='${c}'><img src="${v}" loading="lazy" alt="${htmlEscape(val)}" class="thumb" data-full="${v}" onclick="openModal(this.dataset.full)" onerror="this.style.display='none'"></td>`);
      } else {
        let v = "";
        if (val === null || val === undefined || (typeof val === "number" && isNaN(val))) v = "";
        else if (c === "OTS_INV") {
          const f = Number(val);
          v = isNaN(f) ? String(val) : (Number.isInteger(f) ? String(f) : String(f));
        } else {
          v = String(val);
        }
        cells.push(`<td data-col='${c}'>${htmlEscape(v)}</td>`);
      }
    }
    rowsHtml.push(`<tr${trClass} ${dataAttrs.join(" ")}>${cells.join("")}</tr>`);
  }

  const headerCells = [];
  for (const c of cols) {
    if (c === "#") {
      headerCells.push("<th data-col='#'>#<div class='col-resizer' data-col='#'></div></th>");
    } else {
      headerCells.push(
        `<th data-col='${c}'>${c} <span class='sortbtn' data-col='${c}' onclick='sortTable(this)' style='cursor:pointer'>&#8597;</span><br><div style='height:10px; line-height:10px;'>&nbsp;</div><button class='filterbtn' data-col='${c}'>Filtrar</button><div class='filterpanel' data-col='${c}' style='display:none'>${dropdowns[c]}</div><div class='col-resizer' data-col='${c}'></div></th>`
      );
    }
  }

  const htmlDoc = `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>Galeria de imagenes - Inventario</title>
<style>
body { font-family: Arial, sans-serif; margin: 20px; }
table { border-collapse: collapse; width: 100%; table-layout: fixed; }
th, td { border: 1px solid #ccc; padding: 6px; vertical-align: top; font-size: 12px; }
th { background: #333; color: #fff; position: sticky; top: 0; }
tr.red { background: #ffcccc; }
.col-resizer { position: absolute; right: 0; top: 0; bottom: 0; width: 6px; cursor: col-resize; z-index: 6; background: transparent; }
.col-resizer:hover { background: rgba(255,255,255,0.35); }
th[data-col="#"], td[data-col="#"] { width: 30px; min-width: 30px; }
th[data-col="Color_code"], td[data-col="Color_code"] { width: 80px; min-width: 80px; }
th[data-col="Size_Desc"], td[data-col="Size_Desc"] { width: 70px; min-width: 70px; }
th[data-col="OTS_INV"], td[data-col="OTS_INV"] { width: 55px; min-width: 55px; }
th[data-col="TLGDIV_NAME"], td[data-col="TLGDIV_NAME"] { width: 90px; min-width: 90px; }
th[data-col="Style"], td[data-col="Style"] { width: 70px; min-width: 70px; }
th[data-col="Style_Name"], td[data-col="Style_Name"] { width: 176px; min-width: 176px; }
th[data-col="Color_Name"], td[data-col="Color_Name"] { width: 80px; min-width: 80px; }
th[data-col="Imagen"], td[data-col="Imagen"] { width: 130px; max-width: 130px; }
th.filtered { background: #00bcd4 !important; color: #000 !important; }
img { max-width: 120px; max-height: 120px; object-fit: contain; }
.sortbtn { cursor: pointer; font-size: 12px; color: #fff; user-select: none; }
th { position: relative; }
.filterbtn { font-size: 11px; padding: 2px 6px; cursor: pointer; }
.filterpanel { position: absolute; top: 100%; left: 0; z-index: 1000; background: #fff; color: #000; border: 1px solid #ccc; padding: 6px; max-height: 260px; overflow: auto; text-align: left; min-width: 160px; box-shadow: 2px 8px rgba(0,0,0,0.15); }
.fitem { display: block; white-space: nowrap; font-size: 11px; color: #000; }
.filtersearch { width: 95%; margin-bottom: 4px; font-size: 11px; }
#search { margin-bottom: 10px; padding: 8px; width: 300px; }
select.colfilter { width: 100%; font-size: 11px; }
.thumb { cursor: zoom-in; }
.modal { display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.85); z-index: 9999; align-items: center; justify-content: center; flex-direction: column; padding: 20px; }
.modal.open { display: flex; }
.modal img { max-width: 90vw; max-height: 75vh; transition: transform .15s ease; transform-origin: center center; }
.modal-close { position: absolute; top: 12px; right: 20px; color: #fff; font-size: 32px; cursor: pointer; line-height: 1; }
.modal-controls { margin-top: 12px; }
.modal-controls button { margin: 0 4px; padding: 6px 14px; font-size: 13px; cursor: pointer; }
</style>
</head>
<body>
<h2>Galeria de imagenes - Inventario (${df.length} filas) <span id="rowCount" style="font-size:13px; font-weight:normal;">(mostrando ${df.length})</span></h2>
<input type="text" id="search" placeholder="Buscar texto en toda la fila...">
<button id="resetFilters" style="margin-left:10px; padding:8px 12px; cursor:pointer;">Reset filters</button>
<table id="tbl">
<thead><tr>${headerCells.join("")}</tr></thead>
<tbody>
${rowsHtml.join("\n")}
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
const COLS = ${JSON.stringify(cols)};
const input=document.getElementById('search');
function rowPassesFilters(tr, filtersToUse, q){
  if(!tr.textContent.toLowerCase().includes(q)) return false;
  for(const col in filtersToUse){
    const val = tr.getAttribute('data-'+col);
    const f = filtersToUse[col];
    if (f instanceof Set) {
      if (f.size === 0 || !f.has(val)) return false;
    } else if (f && val !== f) return false;
  }
  return true;
}

function applyFilters(){
  const q=input.value.toLowerCase();
  const filters={};
  const filteredCols = new Set();
  document.querySelectorAll('.filterpanel').forEach(panel=>{
    const col = panel.dataset.col;
    const boxes = panel.querySelectorAll('.colchk');
    const checked = [...boxes].filter(b=>b.checked && b.value !== '__ALL__').map(b=>b.value);
    const allBox = panel.querySelector('.colchk[value="__ALL__"]');
    if (allBox && allBox.checked) {
      // todos seleccionados -> sin filtro
    } else {
      filteredCols.add(col);
      if (checked.length === 0) {
        filters[col] = new Set();
      } else {
        filters[col] = new Set(checked);
      }
    }
  });

  // Filtro cruzado (faceted): disponibilidad de checkboxes segun los demas filtros.
  // Solo afecta habilitacion/visual, no marca como filtrada la columna.
  const allRows = Array.from(document.querySelectorAll('#tbl tbody tr'));
  document.querySelectorAll('.filterpanel').forEach(panel=>{
    const targetCol = panel.dataset.col;
    const otherFilters = {};
    for(const c in filters){
      if (c !== targetCol) otherFilters[c] = filters[c];
    }
    const available = new Set();
    allRows.forEach(tr=>{
      if (rowPassesFilters(tr, otherFilters, q)) {
        available.add(tr.getAttribute('data-'+targetCol));
      }
    });
    panel.querySelectorAll('.colchk').forEach(cb=>{
      if (cb.value === '__ALL__') return;
      const isAvailable = available.has(cb.value);
      cb.disabled = !isAvailable;
      const lbl = cb.parentElement;
      if (lbl) {
        lbl.style.opacity = isAvailable ? '1' : '0.45';
        lbl.style.pointerEvents = isAvailable ? '' : 'none';
      }
    });
  });

  // Resaltar solo las columnas donde el usuario deselecciona checks (no las afectadas).
  document.querySelectorAll('th').forEach(th=>{
    const btn = th.querySelector('.filterbtn');
    const col = btn ? btn.dataset.col : null;
    if (col && filteredCols.has(col)) th.classList.add('filtered');
    else th.classList.remove('filtered');
  });
  document.querySelectorAll('#tbl tbody tr').forEach(tr=>{
    let show = tr.textContent.toLowerCase().includes(q);
    for(const col in filters){
      const val = tr.getAttribute('data-'+col);
      if (filters[col] instanceof Set) {
        if (filters[col].size === 0 || !filters[col].has(val)) { show=false; break; }
      } else if (filters[col] && val !== filters[col]) { show=false; break; }
    }
    tr.style.display = show ? '' : 'none';
  });
  const visible = Array.from(document.querySelectorAll('#tbl tbody tr')).filter(tr=>tr.style.display!=='none').length;
  const rc = document.getElementById('rowCount');
  if(rc) rc.textContent = '(mostrando ' + visible + ' de ${df.length})';
}

// === Resize de columnas con el mouse ===
let _resizing = null;
document.addEventListener('mousedown', function(e){
  const handle = (e.target && e.target.classList && e.target.classList.contains('col-resizer')) ? e.target : null;
  if (!handle) return;
  const th = handle.parentElement;
  _resizing = { th: th, startX: e.clientX, startWidth: th.getBoundingClientRect().width };
  document.body.style.cursor = 'col-resize';
  document.body.style.userSelect = 'none';
  e.preventDefault();
  e.stopPropagation();
});
document.addEventListener('mousemove', function(e){
  if (!_resizing) return;
  const delta = e.clientX - _resizing.startX;
  const newWidth = Math.max(60, _resizing.startWidth + delta);
  _resizing.th.style.width = newWidth + 'px';
  _resizing.th.style.minWidth = newWidth + 'px';
});
document.addEventListener('mouseup', function(){
  if (!_resizing) return;
  // guardar anchos
  const widths = {};
  document.querySelectorAll('th[data-col]').forEach(function(th){ widths[th.dataset.col] = th.getBoundingClientRect().width; });
  try { localStorage.setItem('colWidths', JSON.stringify(widths)); } catch(err) {}
  _resizing = null;
  document.body.style.cursor = '';
  document.body.style.userSelect = '';
});
// restaurar anchos guardados
try {
  const saved = JSON.parse(localStorage.getItem('colWidths') || '{}');
  Object.keys(saved).forEach(function(col){
    const th = document.querySelector('th[data-col="'+col+'"]');
    if (th) { th.style.width = saved[col] + 'px'; th.style.minWidth = saved[col] + 'px'; }
  });
} catch(err) {}

input.addEventListener('input', applyFilters);
// inicializar disponibilidad cruzada al cargar
applyFilters();
const resetBtn = document.getElementById('resetFilters');
if (resetBtn) {
  resetBtn.addEventListener('click', function(){
    input.value = '';
    document.querySelectorAll('.filterpanel .colchk').forEach(function(b){ b.checked = true; });
    applyFilters();
  });
}

let sortState = { col: null, asc: true };
function sortTable(btn){
  const col = btn.dataset.col;
  const asc = (sortState.col === col) ? !sortState.asc : true;
  sortState = { col: col, asc: asc };
  const tbody = document.querySelector('#tbl tbody');
  const rows = Array.from(tbody.querySelectorAll('tr'));
  const idx = COLS.indexOf(col);
  if (idx < 0) return;
  rows.sort(function(a,b){
    const va = (a.cells[idx] ? a.cells[idx].textContent : '').trim();
    const vb = (b.cells[idx] ? b.cells[idx].textContent : '').trim();
    const na = parseFloat(va.replace(/,/g,''));
    const nb = parseFloat(vb.replace(/,/g,''));
    if (!isNaN(na) && !isNaN(nb)) return asc ? (na - nb) : (nb - na);
    return asc ? va.localeCompare(vb) : vb.localeCompare(va);
  });
  rows.forEach(function(r){ tbody.appendChild(r); });
  document.querySelectorAll('.sortbtn').forEach(function(b){ b.textContent = '\\u2195'; });
  btn.textContent = asc ? '\\u25B2' : '\\u25BC';
  applyFilters();
}


document.addEventListener('click', function(e){
  if (e.target && e.target.classList && e.target.classList.contains('filterbtn')){
    const col = e.target.dataset.col;
    const panel = document.querySelector('.filterpanel[data-col="'+col+'"]');
    if (panel) panel.style.display = panel.style.display === 'none' ? 'block' : 'none';
  } else if (!e.target.closest('.filterpanel') && !(e.target.classList && e.target.classList.contains('filterbtn'))){
    document.querySelectorAll('.filterpanel').forEach(function(p){ p.style.display='none'; });
  }
});
document.addEventListener('change', function(e){
  if (e.target && e.target.classList && e.target.classList.contains('colchk') && e.target.value === '__ALL__'){
    const col = e.target.dataset.col;
    const panel = document.querySelector('.filterpanel[data-col="'+col+'"]');
    if (panel) panel.querySelectorAll('.colchk').forEach(function(b){ b.checked = e.target.checked; });
    applyFilters();
  } else if (e.target && e.target.classList && e.target.classList.contains('colchk')){
    const col = e.target.dataset.col;
    const panel = document.querySelector('.filterpanel[data-col="'+col+'"]');
    if (panel){
      const allBox = panel.querySelector('.colchk[value="__ALL__"]');
      const boxes = [...panel.querySelectorAll('.colchk')].filter(function(b){ return b.value !== '__ALL__'; });
      if (allBox) allBox.checked = boxes.every(function(b){ return b.checked; });
    }
    applyFilters();
  }
});
document.addEventListener('input', function(e){
  if (e.target && e.target.classList && e.target.classList.contains('filtersearch')){
    const panel = e.target.closest('.filterpanel');
    const q = e.target.value.toLowerCase();
    panel.querySelectorAll('.fitem').forEach(function(lbl){
      const txt = lbl.textContent.toLowerCase();
      lbl.style.display = txt.includes(q) ? '' : 'none';
    });
  }
});

let zoomLevel = 1;
function openModal(src){
  const modal = document.getElementById('modal');
  const img = document.getElementById('modalImg');
  if(!modal || !img) return;
  img.src = src;
  zoomLevel = 1;
  img.style.transform = 'scale(1)';
  modal.classList.add('open');
}
function closeModal(){
  const modal = document.getElementById('modal');
  const img = document.getElementById('modalImg');
  if(modal) modal.classList.remove('open');
  if(img){ img.src = ''; img.style.transform = 'scale(1)'; }
  zoomLevel = 1;
}
function setZoom(z){
  zoomLevel = z;
  const img = document.getElementById('modalImg');
  if(img) img.style.transform = 'scale(' + z + ')';
}
document.addEventListener('keydown', function(e){ if(e.key === 'Escape') closeModal(); });
const _modalEl = document.getElementById('modal');
if(_modalEl){
  _modalEl.addEventListener('wheel', function(e){
    if(!_modalEl.classList.contains('open')) return;
    e.preventDefault();
    zoomLevel = Math.min(5, Math.max(1, zoomLevel + (e.deltaY < 0 ? 0.5 : -0.5)));
    setZoom(zoomLevel);
  }, { passive: false });
}

</script>
</body>
</html>`;

  const out = path.join(BASE, "index.html");
  fs.writeFileSync(out, htmlDoc, "utf8");
  console.log("Generado:", out);
}

main().catch((e) => { console.error(e); process.exit(1); });
