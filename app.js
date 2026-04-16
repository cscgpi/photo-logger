import * as XLSX from 'https://esm.sh/xlsx@0.20.3';
import * as docx from 'https://esm.sh/docx@7.8.2';
import exifr from 'https://esm.sh/exifr@7.1.3';

// ─── App state ───────────────────────────────────────────────────────────────

let logRows = null;       // [{filename, caption}]
let photoMap = new Map(); // lowercase filename → File

// ─── DOM refs ────────────────────────────────────────────────────────────────

const logInput    = document.getElementById('log-input');
const photosInput = document.getElementById('photos-input');
const logZone     = document.getElementById('log-zone');
const photosZone  = document.getElementById('photos-zone');
const generateBtn = document.getElementById('generate-btn');
const statusEl    = document.getElementById('status');
const logInfo     = document.getElementById('log-info');
const photosInfo  = document.getElementById('photos-info');

// ─── UI helpers ──────────────────────────────────────────────────────────────

function setStatus(msg, type = '') {
  statusEl.textContent = msg;
  statusEl.className = 'status visible' + (type ? ' ' + type : '');
}

function updateBtn() {
  generateBtn.disabled = !(logRows !== null && photoMap.size > 0);
}

// ─── Photo log parsing ───────────────────────────────────────────────────────

async function handleLogFile(file) {
  try {
    setStatus('Reading photo log…');
    const buf  = await file.arrayBuffer();
    const wb   = XLSX.read(buf, { type: 'array' });
    const ws   = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    logRows = parseLog(data);
    logInfo.textContent = `${file.name} — ${logRows.length} photo(s) to process`;
    logZone.classList.add('has-file');
    updateBtn();
    if (logRows.length && photoMap.size) setStatus('');
    else if (!photoMap.size) setStatus('Now add photos.');
  } catch (err) {
    setStatus('Error reading log: ' + err.message, 'error');
    logRows = null;
    updateBtn();
  }
}

function detectColumns(headers) {
  const lc = headers.map(h => String(h).toLowerCase().trim());

  function find(variants) {
    for (const v of variants) {
      const i = lc.indexOf(v);
      if (i >= 0) return i;
    }
    for (const v of variants) {
      const i = lc.findIndex(h => h.includes(v));
      if (i >= 0) return i;
    }
    return -1;
  }

  return {
    filenameCol: find([
      'filename', 'file name', 'file_name', 'file', 'photo name',
      'photo_name', 'photo', 'image name', 'image_name', 'image', 'img',
    ]),
    captionCol: find([
      'caption', 'description', 'desc', 'text', 'note', 'notes', 'label',
      'comment', 'comments',
    ]),
    includeCol: find([
      'include', 'use', 'show', 'selected', 'select', 'flag',
      'process', 'include?', 'use?', 'active',
    ]),
  };
}

function parseLog(data) {
  if (!data.length || !data[0].length) throw new Error('Log file appears empty.');

  const headers = data[0].map(h => String(h ?? ''));
  const { filenameCol, captionCol, includeCol } = detectColumns(headers);

  if (filenameCol < 0) {
    throw new Error(
      'Could not find a filename column. ' +
      'Expected a header like "Filename", "File", "Photo", or "Image".'
    );
  }
  if (captionCol < 0) {
    throw new Error(
      'Could not find a caption column. ' +
      'Expected a header like "Caption", "Description", or "Notes".'
    );
  }

  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const filename = String(row[filenameCol] ?? '').trim();
    if (!filename) continue;

    // Apply include filter only when the column exists
    if (includeCol >= 0) {
      const val = String(row[includeCol] ?? '').trim().toLowerCase();
      if (!['yes', 'true', '1', 'y', 'x'].includes(val)) continue;
    }

    rows.push({
      filename,
      caption: String(row[captionCol] ?? '').trim(),
    });
  }

  if (!rows.length) {
    throw new Error(
      'No rows to process. ' +
      'If an include column exists, check that the values are "Yes" or "TRUE".'
    );
  }
  return rows;
}

// ─── Photo file handling ──────────────────────────────────────────────────────

function handlePhotoFiles(files) {
  photoMap.clear();
  for (const file of files) {
    if (!file.type.startsWith('image/')) continue;
    photoMap.set(file.name.toLowerCase(), file);
  }
  if (photoMap.size > 0) {
    photosInfo.textContent = `${photoMap.size} image(s) loaded`;
    photosZone.classList.add('has-file');
  } else {
    photosInfo.textContent = 'No image files found in selection.';
    photosZone.classList.remove('has-file');
  }
  updateBtn();
  if (logRows && photoMap.size) setStatus('');
}

// ─── EXIF orientation correction ─────────────────────────────────────────────

async function processImageFile(file) {
  let orientation = 1;
  try {
    const tags = await exifr.parse(file, ['Orientation']);
    if (tags?.Orientation) orientation = tags.Orientation;
  } catch (_) { /* no EXIF — treat as orientation 1 */ }

  const url = URL.createObjectURL(file);
  let img;
  try {
    img = await loadImage(url);
  } finally {
    URL.revokeObjectURL(url);
  }

  const canvas = applyOrientation(img, orientation);

  const buffer = await new Promise((resolve, reject) => {
    canvas.toBlob(async blob => {
      if (!blob) { reject(new Error('Canvas toBlob returned null')); return; }
      resolve(await blob.arrayBuffer());
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
  const w = img.naturalWidth;
  const h = img.naturalHeight;
  const canvas = document.createElement('canvas');
  const ctx = canvas.getContext('2d');

  // Orientations 5–8 involve a 90°/270° rotation that swaps width and height
  const swapped = orientation >= 5;
  canvas.width  = swapped ? h : w;
  canvas.height = swapped ? w : h;

  // Apply the affine transform that maps the stored pixels to the correct view
  switch (orientation) {
    case 2: ctx.transform(-1,  0,  0,  1,  w,  0); break;
    case 3: ctx.transform(-1,  0,  0, -1,  w,  h); break;
    case 4: ctx.transform( 1,  0,  0, -1,  0,  h); break;
    case 5: ctx.transform( 0,  1,  1,  0,  0,  0); break;
    case 6: ctx.transform( 0,  1, -1,  0,  h,  0); break;
    case 7: ctx.transform( 0, -1, -1,  0,  h,  w); break;
    case 8: ctx.transform( 0, -1,  1,  0,  0,  w); break;
    default: break; // orientation 1 — no transform needed
  }

  ctx.drawImage(img, 0, 0);
  return canvas;
}

// ─── Document layout constants ────────────────────────────────────────────────

// OOXML image dimensions are specified in EMUs (English Metric Units).
// 1 inch = 914 400 EMU.  We treat photos as 96 DPI for layout purposes.
const EMU_PER_INCH  = 914400;
const EMU_PER_PX    = EMU_PER_INCH / 96; // 9 525 EMU per pixel at 96 DPI
const MAX_W_EMU     = Math.round(6.5 * EMU_PER_INCH); // full printable width
const MAX_H_EMU     = Math.round(4.0 * EMU_PER_INCH); // half printable height

// Paragraph spacing is in twips (1 twip = 1/1440 inch)
const GAP_IMG_CAPTION = 40;   // ~2 mm between photo and its caption
const GAP_BETWEEN     = 240;  // ~4 mm between the two photo blocks on a page

function calcEmuDimensions(pixelW, pixelH) {
  let wEMU = Math.round(pixelW * EMU_PER_PX);
  let hEMU = Math.round(pixelH * EMU_PER_PX);

  // Scale down to fit the max box; never upscale
  if (wEMU > MAX_W_EMU || hEMU > MAX_H_EMU) {
    const scale = Math.min(MAX_W_EMU / wEMU, MAX_H_EMU / hEMU);
    wEMU = Math.round(wEMU * scale);
    hEMU = Math.round(hEMU * scale);
  }
  return { width: wEMU, height: hEMU };
}

// ─── Document generation ─────────────────────────────────────────────────────

async function buildDocument(rows, photos, onProgress) {
  const children = [];

  for (let i = 0; i < rows.length; i++) {
    onProgress(`Processing photo ${i + 1} of ${rows.length}…`);

    const row       = rows[i];
    const isNewPage = i > 0 && i % 2 === 0;
    const isFirst   = i % 2 === 0;   // first of the two blocks on a page
    const file      = photos.get(row.filename.toLowerCase());

    // ── Photo block ─────────────────────────────────────────────────────────
    if (file) {
      let processed;
      try {
        processed = await processImageFile(file);
      } catch (err) {
        // Fall through to missing-photo placeholder on processing failure
        processed = null;
        console.warn(`Could not process ${row.filename}: ${err.message}`);
      }

      if (processed) {
        const { data, width, height } = processed;
        const dims = calcEmuDimensions(width, height);

        children.push(new docx.Paragraph({
          alignment:      docx.AlignmentType.CENTER,
          pageBreakBefore: isNewPage,
          spacing:        { before: 0, after: GAP_IMG_CAPTION },
          children: [
            new docx.ImageRun({
              data,
              transformation: dims,
            }),
          ],
        }));
      } else {
        children.push(missingPara(row.filename, isNewPage));
      }
    } else {
      children.push(missingPara(row.filename, isNewPage));
    }

    // ── Caption ──────────────────────────────────────────────────────────────
    children.push(new docx.Paragraph({
      alignment: docx.AlignmentType.CENTER,
      spacing:   { before: 0, after: isFirst ? GAP_BETWEEN : 0 },
      children:  [new docx.TextRun({ text: row.caption })],
    }));
  }

  const doc = new docx.Document({
    sections: [{
      properties: {
        page: {
          // Letter: 8.5 × 11 inches in twips (1 twip = 1/1440 inch)
          size:   { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      children,
    }],
  });

  return docx.Packer.toBlob(doc);
}

function missingPara(filename, isNewPage) {
  return new docx.Paragraph({
    alignment:       docx.AlignmentType.CENTER,
    pageBreakBefore: isNewPage,
    spacing:         { before: 0, after: GAP_IMG_CAPTION },
    children: [
      new docx.TextRun({ text: `MISSING PHOTO: ${filename}`, bold: true }),
    ],
  });
}

// ─── Generate handler ─────────────────────────────────────────────────────────

async function generate() {
  generateBtn.disabled = true;
  setStatus('Starting…');

  try {
    const blob = await buildDocument(logRows, photoMap, msg => setStatus(msg));

    const url = URL.createObjectURL(blob);
    const a   = document.createElement('a');
    a.href     = url;
    a.download = 'photo-sheet.docx';
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);

    setStatus(
      `Done! photo-sheet.docx downloaded (${logRows.length} photo(s)).`,
      'success'
    );
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

  async function walkEntry(entry) {
    if (entry.isFile) {
      const file = await new Promise((res, rej) => entry.file(res, rej));
      files.push(file);
    } else if (entry.isDirectory) {
      const reader = entry.createReader();
      let batch;
      do {
        batch = await new Promise((res, rej) => reader.readEntries(res, rej));
        for (const e of batch) await walkEntry(e);
      } while (batch.length > 0);
    }
  }

  for (const item of [...dt.items]) {
    const entry = item.webkitGetAsEntry?.();
    if (entry) {
      await walkEntry(entry);
    } else {
      const f = item.getAsFile?.();
      if (f) files.push(f);
    }
  }
  return files;
}

function wireZone(zone, { onFile, onFiles }) {
  zone.addEventListener('dragover', e => {
    e.preventDefault();
    zone.classList.add('drag-over');
  });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
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

// ─── Event wiring ────────────────────────────────────────────────────────────

document.getElementById('log-btn').addEventListener('click', () => logInput.click());
document.getElementById('photos-btn').addEventListener('click', () => photosInput.click());

logInput.addEventListener('change', () => {
  if (logInput.files[0]) handleLogFile(logInput.files[0]);
});

photosInput.addEventListener('change', () => {
  if (photosInput.files.length) handlePhotoFiles([...photosInput.files]);
});

wireZone(logZone,    { onFile:  f  => handleLogFile(f)     });
wireZone(photosZone, { onFiles: fs => handlePhotoFiles(fs) });

generateBtn.addEventListener('click', generate);
