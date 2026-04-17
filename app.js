/* Photo Sheet Generator — app.js
 * Pure browser script; depends on three UMD globals loaded before this file:
 *   XLSX   (SheetJS)
 *   exifr  (exifr)
 *   JSZip  (JSZip)
 */
'use strict';

// ─── State ────────────────────────────────────────────────────────────────────

let logRows   = null;       // [{filename, caption}]
let photoMap  = new Map();  // lowercase filename → File
let thumbUrls = new Map();  // lowercase filename → objectURL (preview only)

// ─── DOM refs (assigned in init) ─────────────────────────────────────────────

let logInput, logZone, photosZone, generateBtn, statusEl, logInfo, photosInfo;

// ─── UI helpers ───────────────────────────────────────────────────────────────

function setStatus(msg, type) {
  statusEl.textContent = msg;
  statusEl.className = 'status visible' + (type ? ' ' + type : '');
}
function clearStatus() { statusEl.className = 'status'; }
function updateBtn()   { generateBtn.disabled = !(logRows && logRows.length && photoMap.size); }

function escHtml(s) {
  const d = document.createElement('div');
  d.textContent = s;
  return d.innerHTML;
}

// ─── Log parsing ──────────────────────────────────────────────────────────────

async function handleLogFile(file) {
  try {
    setStatus('Reading photo log…');
    const buf  = await file.arrayBuffer();
    const wb   = XLSX.read(buf, { type: 'array' });
    const ws   = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    logRows = parseLog(data);
    logInfo.textContent = `${file.name} — ${logRows.length} row(s)`;
    logZone.classList.add('has-file');
    updateBtn();
    renderLogPreview();
    if (photoMap.size) clearStatus(); else setStatus('Now add your photos.');
  } catch (err) {
    setStatus('Error reading log: ' + err.message, 'error');
    logRows = null;
    updateBtn();
  }
}

function detectColumns(headers, dataRows) {
  const lc = headers.map(h => String(h).toLowerCase().trim());

  const find = variants => {
    for (const v of variants) { const i = lc.indexOf(v);                    if (i >= 0) return i; }
    for (const v of variants) { const i = lc.findIndex(h => h.includes(v)); if (i >= 0) return i; }
    return -1;
  };

  // ── Filename column: scan cell VALUES first for image-file extensions ──────
  // This handles cases where the header says "Digital Photo No." or any other
  // non-obvious name but the cells contain "IMG_9781.JPG" etc.
  const imgExtRe = /\.(jpe?g|png|tiff?|gif|bmp|webp|heic|dng|cr2|nef|arw)$/i;
  let filenameCol = -1;
  for (let col = 0; col < headers.length && filenameCol < 0; col++) {
    if (dataRows.some(row => imgExtRe.test(String(row[col] ?? '')))) {
      filenameCol = col;
    }
  }
  // Fall back to header-name matching if no image-extension values were found
  if (filenameCol < 0) {
    filenameCol = find([
      'filename', 'file name', 'file_name', 'file',
      'digital photo no', 'photo no', 'photo number', 'photo num', 'photo#',
      'photo name', 'photo_name', 'photo',
      'image name', 'image_name', 'image', 'img',
    ]);
  }

  return {
    filenameCol,
    captionCol: find(['caption','description','desc','text','note','notes','label','comment','comments']),
    includeCol: find(['include','use','show','selected','select','flag','process','include?','use?','active']),
  };
}

function parseLog(data) {
  if (!data.length || !data[0].length) throw new Error('Log file appears empty.');
  const headers  = data[0].map(h => String(h ?? ''));
  const dataRows = data.slice(1);
  const { filenameCol, captionCol, includeCol } = detectColumns(headers, dataRows);
  if (filenameCol < 0) throw new Error('Cannot find a filename column. Expected a column whose cells contain image filenames (e.g. IMG_001.JPG), or a header like "Filename", "File", or "Photo".');
  if (captionCol  < 0) throw new Error('Cannot find a caption column. Expected a header like "Caption", "Description", or "Notes".');
  const rows = [];
  for (const row of dataRows) {
    const filename = String(row[filenameCol] ?? '').trim();
    if (!filename) continue;
    if (includeCol >= 0) {
      const val = String(row[includeCol] ?? '').trim().toLowerCase();
      if (!['yes','true','1','y','x'].includes(val)) continue;
    }
    rows.push({ filename, caption: String(row[captionCol] ?? '').trim() });
  }
  if (!rows.length) throw new Error('No rows to process. If an include column exists, check that values are "Yes" or "TRUE".');
  return rows;
}

// ─── Photo file handling ──────────────────────────────────────────────────────

function handlePhotoFiles(files) {
  for (const f of files) {
    if (!f.type.startsWith('image/')) continue;
    const key = f.name.toLowerCase();
    if (thumbUrls.has(key)) { URL.revokeObjectURL(thumbUrls.get(key)); thumbUrls.delete(key); }
    photoMap.set(key, f);
  }
  if (photoMap.size) {
    photosInfo.textContent = `${photoMap.size} image(s) loaded`;
    photosZone.classList.add('has-file');
  } else {
    photosInfo.textContent = 'No image files found in that selection.';
    photosZone.classList.remove('has-file');
  }
  updateBtn();
  renderPhotosPreview();
  renderLogPreview();
  if (logRows && photoMap.size) clearStatus();
}

function removePhoto(name) {
  if (thumbUrls.has(name)) { URL.revokeObjectURL(thumbUrls.get(name)); thumbUrls.delete(name); }
  photoMap.delete(name);
  if (photoMap.size) {
    photosInfo.textContent = `${photoMap.size} image(s) loaded`;
    renderPhotosPreview();
  } else {
    photosInfo.textContent = '';
    photosZone.classList.remove('has-file');
    document.getElementById('photos-preview').style.display = 'none';
    document.getElementById('photos-grid').innerHTML = '';
  }
  updateBtn();
  renderLogPreview();
}

function clearPhotos() {
  for (const url of thumbUrls.values()) URL.revokeObjectURL(url);
  thumbUrls.clear();
  photoMap.clear();
  photosZone.classList.remove('has-file');
  photosInfo.textContent = '';
  document.getElementById('photos-preview').style.display = 'none';
  document.getElementById('photos-grid').innerHTML = '';
  updateBtn();
  renderLogPreview();
}

function clearLog() {
  logRows = null;
  logZone.classList.remove('has-file');
  logInfo.textContent = '';
  document.getElementById('log-preview').style.display = 'none';
  document.getElementById('log-tbody').innerHTML = '';
  updateBtn();
  clearStatus();
}

// ─── Preview: log table ───────────────────────────────────────────────────────

function renderLogPreview() {
  if (!logRows) return;
  const panel     = document.getElementById('log-preview');
  const tbody     = document.getElementById('log-tbody');
  const meta      = document.getElementById('log-preview-meta');
  const hasPhotos = photoMap.size > 0;

  meta.textContent = `${logRows.length} row(s)`;
  tbody.innerHTML  = '';

  logRows.forEach((row, i) => {
    const found = photoMap.has(row.filename.toLowerCase());
    const badge = !hasPhotos
      ? '<span class="badge badge-pending">—</span>'
      : found
        ? '<span class="badge badge-found">Found</span>'
        : '<span class="badge badge-missing">Missing</span>';

    const tr = document.createElement('tr');
    tr.innerHTML =
      `<td class="col-num">${i + 1}</td>` +
      `<td class="col-file">${escHtml(row.filename)}</td>` +
      `<td class="col-caption">${escHtml(row.caption)}</td>` +
      `<td class="col-status">${badge}</td>`;
    tbody.appendChild(tr);
  });

  panel.style.display = '';
}

// ─── Preview: photos grid ─────────────────────────────────────────────────────

function renderPhotosPreview() {
  if (!photoMap.size) return;
  const panel = document.getElementById('photos-preview');
  const grid  = document.getElementById('photos-grid');
  const meta  = document.getElementById('photos-preview-meta');

  meta.textContent = `${photoMap.size} file(s)`;
  grid.innerHTML   = '';

  for (const [name, file] of photoMap) {
    if (!thumbUrls.has(name)) thumbUrls.set(name, URL.createObjectURL(file));

    const item = document.createElement('div');
    item.className = 'photo-item';

    const wrap = document.createElement('div');
    wrap.className = 'thumb-wrap';

    const img = document.createElement('img');
    img.src     = thumbUrls.get(name);
    img.alt     = name;
    img.loading = 'lazy';

    const removeBtn = document.createElement('button');
    removeBtn.className   = 'photo-remove-btn';
    removeBtn.title       = 'Remove';
    removeBtn.textContent = '×';
    removeBtn.addEventListener('click', () => removePhoto(name));

    wrap.append(img, removeBtn);

    const lbl = document.createElement('div');
    lbl.className   = 'thumb-name';
    lbl.textContent = name;
    lbl.title       = name;

    item.append(wrap, lbl);
    grid.appendChild(item);
  }

  panel.style.display = '';
}

// ─── EXIF orientation correction via canvas ───────────────────────────────────

async function processImageFile(file) {
  let orientation = 1;
  try {
    const tags = await exifr.parse(file, ['Orientation']);
    if (tags && tags.Orientation) orientation = tags.Orientation;
  } catch (_) { /* no EXIF — orientation 1 */ }

  const url = URL.createObjectURL(file);
  let img;
  try   { img = await loadImage(url); }
  finally { URL.revokeObjectURL(url); }

  const canvas = applyOrientation(img, orientation);
  const buffer = await new Promise((resolve, reject) => {
    canvas.toBlob(blob => {
      if (!blob) { reject(new Error('Canvas.toBlob returned null')); return; }
      blob.arrayBuffer().then(resolve, reject);
    }, 'image/jpeg', 0.92);
  });
  return { data: buffer, width: canvas.width, height: canvas.height };
}

function loadImage(url) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload  = () => resolve(img);
    img.onerror = () => reject(new Error('Failed to decode image'));
    img.src = url;
  });
}

function applyOrientation(img, orientation) {
  const w = img.naturalWidth, h = img.naturalHeight;
  const canvas = document.createElement('canvas');
  const ctx    = canvas.getContext('2d');
  const swap   = orientation >= 5;
  canvas.width  = swap ? h : w;
  canvas.height = swap ? w : h;
  switch (orientation) {
    case 2: ctx.transform(-1, 0,  0,  1,  w, 0); break;
    case 3: ctx.transform(-1, 0,  0, -1,  w, h); break;
    case 4: ctx.transform( 1, 0,  0, -1,  0, h); break;
    case 5: ctx.transform( 0, 1,  1,  0,  0, 0); break;
    case 6: ctx.transform( 0, 1, -1,  0,  h, 0); break;
    case 7: ctx.transform( 0,-1, -1,  0,  h, w); break;
    case 8: ctx.transform( 0,-1,  1,  0,  0, w); break;
  }
  ctx.drawImage(img, 0, 0);
  return canvas;
}

// ─── DOCX layout constants ────────────────────────────────────────────────────

const EMU_PER_INCH = 914400;
const EMU_PER_PX   = EMU_PER_INCH / 96;
const MAX_W_EMU    = Math.round(6.5 * EMU_PER_INCH);
const MAX_H_EMU    = Math.round(4.0 * EMU_PER_INCH);
const SP_IMG_CAP   = 40;
const SP_CAP_NEXT  = 240;

function calcEmu(pixelW, pixelH) {
  let w = Math.round(pixelW * EMU_PER_PX);
  let h = Math.round(pixelH * EMU_PER_PX);
  if (w > MAX_W_EMU || h > MAX_H_EMU) {
    const scale = Math.min(MAX_W_EMU / w, MAX_H_EMU / h);
    w = Math.round(w * scale);
    h = Math.round(h * scale);
  }
  return { w, h };
}

// ─── OOXML helpers ────────────────────────────────────────────────────────────

function xe(s) {
  return String(s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function imgPara(relId, docPrId, wEmu, hEmu, pageBreak) {
  const pb = pageBreak ? '<w:pageBreakBefore/>' : '';
  return `<w:p>` +
    `<w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="${SP_IMG_CAP}"/>${pb}</w:pPr>` +
    `<w:r><w:drawing>` +
    `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
      `<wp:extent cx="${wEmu}" cy="${hEmu}"/>` +
      `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
      `<wp:docPr id="${docPrId}" name="Image ${docPrId}"/>` +
      `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
        `<pic:pic>` +
          `<pic:nvPicPr>` +
            `<pic:cNvPr id="${docPrId}" name="Image ${docPrId}"/>` +
            `<pic:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1"/></pic:cNvPicPr>` +
          `</pic:nvPicPr>` +
          `<pic:blipFill>` +
            `<a:blip r:embed="${relId}"/>` +
            `<a:stretch><a:fillRect/></a:stretch>` +
          `</pic:blipFill>` +
          `<pic:spPr bwMode="auto">` +
            `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${wEmu}" cy="${hEmu}"/></a:xfrm>` +
            `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
            `<a:noFill/>` +
          `</pic:spPr>` +
        `</pic:pic>` +
      `</a:graphicData></a:graphic>` +
    `</wp:inline>` +
    `</w:drawing></w:r></w:p>`;
}

function missingPara(filename, pageBreak) {
  const pb = pageBreak ? '<w:pageBreakBefore/>' : '';
  return `<w:p>` +
    `<w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="${SP_IMG_CAP}"/>${pb}</w:pPr>` +
    `<w:r><w:rPr><w:b/></w:rPr><w:t>MISSING PHOTO: ${xe(filename)}</w:t></w:r>` +
    `</w:p>`;
}

function captionPara(caption, afterTwips) {
  return `<w:p>` +
    `<w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="${afterTwips}"/></w:pPr>` +
    `<w:r><w:t xml:space="preserve">${xe(caption)}</w:t></w:r>` +
    `</w:p>`;
}

// ─── DOCX assembly (JSZip + raw OOXML) ───────────────────────────────────────

async function buildDocx(rows, photos, onProgress) {
  const zip       = new JSZip();
  const bodyParts = [];
  const imgRels   = [];
  let   imgCount  = 0;

  for (let i = 0; i < rows.length; i++) {
    onProgress(`Processing photo ${i + 1} of ${rows.length}…`);
    const { filename, caption } = rows[i];
    const newPage = i > 0 && i % 2 === 0;
    const isFirst = i % 2 === 0;
    const file    = photos.get(filename.toLowerCase());

    if (file) {
      let proc = null;
      try { proc = await processImageFile(file); } catch (_) {}
      if (proc) {
        imgCount++;
        const relId = `rId${imgCount + 1}`;
        const path  = `media/image${imgCount}.jpg`;
        zip.file(`word/${path}`, proc.data);
        imgRels.push(
          `<Relationship Id="${relId}"` +
          ` Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"` +
          ` Target="${path}"/>`
        );
        const { w, h } = calcEmu(proc.width, proc.height);
        bodyParts.push(imgPara(relId, imgCount, w, h, newPage));
      } else {
        bodyParts.push(missingPara(filename, newPage));
      }
    } else {
      bodyParts.push(missingPara(filename, newPage));
    }

    bodyParts.push(captionPara(caption, isFirst ? SP_CAP_NEXT : 0));
  }

  const docXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document` +
    ` xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"` +
    ` xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"` +
    ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"` +
    ` xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"` +
    ` xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
    ` xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"` +
    ` xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"` +
    ` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"` +
    ` xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"` +
    ` mc:Ignorable="w14 w15">` +
    `<w:body>` + bodyParts.join('') +
    `<w:sectPr>` +
      `<w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>` +
      `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>` +
    `</w:sectPr>` +
    `</w:body></w:document>`;

  const docRels =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
    imgRels.join('') +
    `</Relationships>`;

  const stylesXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
    ` xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"` +
    ` mc:Ignorable="w14 w15"` +
    ` xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"` +
    ` xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">` +
    `<w:docDefaults>` +
      `<w:rPrDefault><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>` +
      `<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>` +
    `</w:docDefaults>` +
    `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
    `</w:styles>`;

  const contentTypes =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml"  ContentType="application/xml"/>` +
    `<Default Extension="jpg"  ContentType="image/jpeg"/>` +
    `<Default Extension="jpeg" ContentType="image/jpeg"/>` +
    `<Default Extension="png"  ContentType="image/png"/>` +
    `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
    `<Override PartName="/word/styles.xml"   ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
    `</Types>`;

  const rootRels =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
    `</Relationships>`;

  zip.file('[Content_Types].xml',           contentTypes);
  zip.file('_rels/.rels',                  rootRels);
  zip.file('word/document.xml',            docXml);
  zip.file('word/_rels/document.xml.rels', docRels);
  zip.file('word/styles.xml',              stylesXml);

  return zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
}

// ─── Generate handler ─────────────────────────────────────────────────────────

async function generate() {
  generateBtn.disabled = true;
  try {
    const blob = await buildDocx(logRows, photoMap, msg => setStatus(msg));
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href = url; a.download = 'photo-sheet.docx';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    setStatus(`Done! photo-sheet.docx downloaded (${logRows.length} photo(s)).`, 'success');
  } catch (err) {
    setStatus('Error: ' + err.message, 'error');
    console.error(err);
  } finally {
    generateBtn.disabled = false;
  }
}

// ─── Drag-and-drop ───────────────────────────────────────────────────────────

async function filesFromDrop(dt) {
  const files = [];

  // Synchronously extract ALL entries/files before the first await.
  // DataTransfer items become inaccessible once the event loop yields,
  // so calling webkitGetAsEntry() inside an async loop loses every item
  // after the first await gap — causing only one file to be captured.
  const entries = [];
  for (const item of [...(dt.items || [])]) {
    const entry = item.webkitGetAsEntry?.();
    if (entry) entries.push(entry);
    else { const f = item.getAsFile?.(); if (f) files.push(f); }
  }

  // Walk directory trees asynchronously now that all entries are captured
  async function walk(entry) {
    if (entry.isFile) {
      files.push(await new Promise((res, rej) => entry.file(res, rej)));
    } else if (entry.isDirectory) {
      const reader = entry.createReader();
      let batch;
      do {
        batch = await new Promise((res, rej) => reader.readEntries(res, rej));
        for (const e of batch) await walk(e);
      } while (batch.length);
    }
  }

  for (const entry of entries) await walk(entry);
  return files;
}

function wireZone(zone, { onFile, onFiles }) {
  zone.addEventListener('dragenter', e => { e.preventDefault(); zone.classList.add('drag-over'); });
  zone.addEventListener('dragover',  e => { e.preventDefault(); });
  zone.addEventListener('dragleave', e => { if (!zone.contains(e.relatedTarget)) zone.classList.remove('drag-over'); });
  zone.addEventListener('drop', async e => {
    e.preventDefault();
    zone.classList.remove('drag-over');
    const files = await filesFromDrop(e.dataTransfer);
    if (onFile) {
      const target = files.find(f => /\.(xlsx|xls|csv)$/i.test(f.name)) ?? files[0];
      if (target) await onFile(target);
    }
    if (onFiles) onFiles(files);
  });
}

// ─── Init ─────────────────────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', () => {
  logInput          = document.getElementById('log-input');
  const photosFilesInput  = document.getElementById('photos-files-input');
  const photosFolderInput = document.getElementById('photos-folder-input');
  logZone     = document.getElementById('log-zone');
  photosZone  = document.getElementById('photos-zone');
  generateBtn = document.getElementById('generate-btn');
  statusEl    = document.getElementById('status');
  logInfo     = document.getElementById('log-info');
  photosInfo  = document.getElementById('photos-info');

  // Reset value before each click so re-selecting the same files re-fires change
  document.getElementById('log-btn').addEventListener('click', () => {
    logInput.value = ''; logInput.click();
  });
  document.getElementById('photos-files-btn').addEventListener('click', () => {
    photosFilesInput.value = ''; photosFilesInput.click();
  });
  document.getElementById('photos-folder-btn').addEventListener('click', () => {
    photosFolderInput.value = ''; photosFolderInput.click();
  });

  logInput         .addEventListener('change', () => { if (logInput.files[0])              handleLogFile(logInput.files[0]); });
  photosFilesInput .addEventListener('change', () => { if (photosFilesInput.files.length)  handlePhotoFiles([...photosFilesInput.files]); });
  photosFolderInput.addEventListener('change', () => { if (photosFolderInput.files.length) handlePhotoFiles([...photosFolderInput.files]); });

  wireZone(logZone,    { onFile:  f  => handleLogFile(f)     });
  wireZone(photosZone, { onFiles: fs => handlePhotoFiles(fs) });

  document.getElementById('log-clear-btn')   .addEventListener('click', clearLog);
  document.getElementById('photos-clear-btn').addEventListener('click', clearPhotos);

  generateBtn.addEventListener('click', generate);
});
