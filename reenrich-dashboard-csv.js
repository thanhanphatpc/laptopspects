/**
 * reenrich-dashboard-csv.js (06/08/2026)
 *
 * Doc lai TOAN BO dashboard/data.csv hien co (khong can artifact scrape moi),
 * tra cuu tab "Part #" / "Segment" BAN MOI NHAT, tinh lai cac cot enrichment
 * (Vendor, Series Group, Segment, CPU Segment, CPU, RAM, SSD, Screen, GPU,
 * V-RAM, Part #, Focus Model) cho TOAN BO dong trong file, roi ghi de lai.
 * Cac cot RAW (Date/Hour/Dealers/SKU/SRP/Promotion Price/.../Status) GIU
 * NGUYEN - chi cot phu thuoc Part# duoc tinh lai. SP co Status=EOL trong
 * Part# se bi loai khoi ket qua (giong logic goc trong enrich-and-export.js).
 *
 * Muc dich: Phuc sua Part# (Focus Model, Series Group, CPU Segment...) thi
 * dashboard cap nhat ngay (~30-60s, chi 2 Sheet API call), khong can doi lan
 * scrape sau (~20-25 phut) hay trigger scrape thu cong.
 *
 * Best-effort: loi o dau thi log ra dashboard/.reenrich-debug.log va dung,
 * KHONG dung gi toi data.csv cu.
 */
const fs = require('fs');
const path = require('path');
const os = require('os');
const { google } = require('googleapis');
// Chuan hoa spec cho dong KHONG khop tab Part # — them 02/09/2026.
const { normalizeCpu } = require('./spec_normalize.js');
const { K } = require('./spec_dictionary.js');

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const CREDS_PATH = path.join(os.tmpdir(), 'reenrich-gcreds.json');
const DATA_CSV_PATH = path.join(__dirname, 'dashboard', 'data.csv');
const DEBUG_LOG_PATH = path.join(__dirname, 'dashboard', '.reenrich-debug.log');
const REQ_TIMEOUT_MS = 25000;

function debugLog(msg) {
  const line = `[${new Date().toISOString()}] ${msg}\n`;
  try { fs.appendFileSync(DEBUG_LOG_PATH, line); } catch (_) {}
  console.log(msg);
}
function csvEscape(v) {
  const s = (v === null || v === undefined) ? '' : String(v);
  if (/[",\n\r]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}
function isFormula(cell) {
  return typeof cell === 'string' && cell.trim().startsWith('=');
}
function buildHeaderIndex(headerRow) {
  const idx = {};
  (headerRow || []).forEach((h, i) => {
    const key = String(h || '').trim();
    if (key) idx[key] = i;
  });
  return idx;
}
// CSV parser don gian (RFC4180 co ban, co quote) - dung 1:1 voi ban trong
// enrich-and-export.js vi cung doc file do chinh script kia sinh ra.
function parseCsvSimple(text) {
  const rows = [];
  let row = [], field = '', inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQuotes) {
      if (c === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else inQuotes = false; }
      else field += c;
    } else {
      if (c === '"') inQuotes = true;
      else if (c === ',') { row.push(field); field = ''; }
      else if (c === '\n') { row.push(field); rows.push(row); row = []; field = ''; }
      else if (c === '\r') { /* skip */ }
      else field += c;
    }
  }
  if (field.length > 0 || row.length > 0) { row.push(field); rows.push(row); }
  return rows.filter(r => r.length > 1 || r[0] !== '');
}
async function readTabFormula(sheets, tabName, lastCol) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${tabName}'!A1:${lastCol}`,
    valueRenderOption: 'FORMULA',
  }, { timeout: REQ_TIMEOUT_MS });
  return res.data.values || [];
}

// Vi tri cot trong dashboard/data.csv (0-indexed) - phai khop dung HEADER
// trong enrich-and-export.js.
const COL = {
  Date: 0, Hour: 1, Dealers: 2, SKU: 3, SRP: 4, PromoPrice: 5, Change: 6,
  Sold: 7, Rate: 8, Vendor: 9, SeriesGroup: 10, Segment: 11, CPUSegment: 12,
  CPU: 13, RAM: 14, SSD: 15, Screen: 16, GPU: 17, VRAM: 18, Weight: 19,
  Link: 20, PartNo: 21, FocusModel: 22, Status: 23,
};

async function main() {
  try { fs.writeFileSync(DEBUG_LOG_PATH, ''); } catch (_) {}
  debugLog('--- Bắt đầu reenrich-dashboard-csv.js ---');

  if (!fs.existsSync(DATA_CSV_PATH)) {
    debugLog('⚠ Chưa có dashboard/data.csv — không có gì để re-enrich, dừng.');
    return;
  }
  if (!SPREADSHEET_ID || !process.env.GOOGLE_CREDENTIALS) {
    debugLog('⚠ Thiếu SPREADSHEET_ID/GOOGLE_CREDENTIALS — bỏ qua.');
    return;
  }

  const raw = fs.readFileSync(DATA_CSV_PATH, 'utf8');
  const parsed = parseCsvSimple(raw);
  const header = parsed[0];
  const dataRows = parsed.slice(1);
  debugLog(`Đọc data.csv hiện có: ${dataRows.length} dòng.`);
  if (dataRows.length === 0) {
    debugLog('⚠ data.csv rỗng — không có gì để re-enrich, dừng.');
    return;
  }

  fs.writeFileSync(CREDS_PATH, process.env.GOOGLE_CREDENTIALS);
  const auth = new google.auth.GoogleAuth({
    keyFile: CREDS_PATH,
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
  });
  const sheets = google.sheets({ version: 'v4', auth });

  // FIX 19/09/2026 [fork thanhanphatpc/laptopspects]: xem chu thich cung
  // trong enrich-and-export.js — Sheet nay chua co tab "Part #"/"Segment",
  // bat loi rieng cho tung tab thay vi de nem loi lam dung ca script.
  let partRows = [];
  try {
    debugLog('Đọc Part # (FORMULA render)...');
    partRows = await readTabFormula(sheets, 'Part #', 'AE');
    debugLog(`  ${Math.max(0, partRows.length - 1)} dòng`);
  } catch (e) {
    debugLog(`⚠ Không đọc được tab "Part #" (${e.message}) — bỏ qua, giữ nguyên spec thô hiện có trong data.csv.`);
  }

  let segmentRows = [];
  try {
    debugLog('Đọc Segment (FORMULA render)...');
    segmentRows = await readTabFormula(sheets, 'Segment', 'D');
    debugLog(`  ${Math.max(0, segmentRows.length - 1)} dòng`);
  } catch (e) {
    debugLog(`⚠ Không đọc được tab "Segment" (${e.message}) — bỏ qua, Series Group/Segment sẽ để trống.`);
  }

  try { fs.unlinkSync(CREDS_PATH); } catch (_) {}

  // Segment!C (tên dòng SP) -> Series Group — 1:1 logic voi enrich-and-export.js
  const segmentTable = [];
  for (let i = 1; i < segmentRows.length; i++) {
    const r = segmentRows[i];
    const segName = r[2];
    if (segName && !isFormula(segName)) segmentTable.push({ segment: String(segName), seriesGroup: r[3] || '' });
  }
  const segmentExactMap = new Map(segmentTable.map(s => [s.segment.toLowerCase(), s]));
  function guessSegment(productName) {
    if (!productName) return null;
    const lower = String(productName).toLowerCase();
    let best = null;
    for (const s of segmentTable) {
      if (lower.includes(s.segment.toLowerCase())) {
        if (!best || s.segment.length > best.segment.length) best = s;
      }
    }
    return best;
  }

  const partHIdx = buildHeaderIndex(partRows[0]);
  const iName = partHIdx['Product Name'];
  const iVendor = partHIdx['Vendor'];
  const iPartNo = partHIdx['Part #'];
  const iCpuSeg = partHIdx['CPU Segment'];
  const iCpu = partHIdx['CPU'];
  const iRam = partHIdx['RAM'];
  const iSsd = partHIdx['SSD'];
  const iScreen = partHIdx['Screen'];
  const iGpu = partHIdx['GPU'];
  const iVram = partHIdx['V-RAM'];
  const iStatus = partHIdx['Status'];
  const iSeriesGroup = partHIdx['Series Group'];
  const iSegment = partHIdx['Segment'];
  const iFocus = partHIdx['Focus Model'];
  debugLog(`Part# header index: ${JSON.stringify(partHIdx)}`);
  if (iName === undefined || iFocus === undefined) {
    debugLog('⚠ Part# thiếu cột bắt buộc (Product Name hoặc Focus Model) — kiểm tra lại header tab.');
  }

  const partMap = new Map();
  for (let i = 1; i < partRows.length; i++) {
    const r = partRows[i];
    const name = iName !== undefined ? r[iName] : undefined;
    if (!name) continue;
    const status = (iStatus !== undefined ? r[iStatus] : '') || '';
    const lRaw = iSeriesGroup !== undefined ? r[iSeriesGroup] : undefined;
    const mRaw = iSegment !== undefined ? r[iSegment] : undefined;

    let segment;
    if (mRaw && !isFormula(mRaw)) segment = String(mRaw);
    else segment = (guessSegment(name) || {}).segment || '';

    let seriesGroup;
    if (lRaw && !isFormula(lRaw)) seriesGroup = String(lRaw);
    else {
      const exact = segmentExactMap.get(String(segment).toLowerCase());
      seriesGroup = exact ? exact.seriesGroup : '';
    }

    const focusRaw = iFocus !== undefined ? r[iFocus] : '';
    const focusModel = (focusRaw && !isFormula(focusRaw)) ? String(focusRaw) : 'No';

    partMap.set(String(name), {
      vendor: (iVendor !== undefined && r[iVendor]) || '',
      partNumber: (iPartNo !== undefined && r[iPartNo]) || '',
      cpuSegment: (iCpuSeg !== undefined && r[iCpuSeg]) || '',
      cpu: (iCpu !== undefined && r[iCpu]) || '',
      ram: (iRam !== undefined && r[iRam]) || '',
      ssd: (iSsd !== undefined && r[iSsd]) || '',
      screen: (iScreen !== undefined && r[iScreen]) || '',
      gpu: (iGpu !== undefined && r[iGpu]) || '',
      vram: (iVram !== undefined && r[iVram]) || '',
      status: String(status).trim(),
      seriesGroup, segment, focusModel,
    });
  }
  debugLog(`Part# map: ${partMap.size} SKU`);

  // Tinh lai enrichment cho TOAN BO dong hien co trong data.csv. Giu nguyen
  // cac cot RAW; chi ghi de cot enrichment KHI Part# co gia tri moi (khong
  // xoa trang gia tri cu neu Part# lan nay khong khop - an toan hon).
  let matched = 0, skippedEOL = 0, normalized = 0;
  const outRows = [];
  for (const row of dataRows) {
    const sku = row[COL.SKU];
    const info = partMap.get(sku);
    if (info && info.status === 'EOL') { skippedEOL++; continue; }
    if (info) {
      matched++;
      row[COL.Vendor] = info.vendor || row[COL.Vendor];
      row[COL.SeriesGroup] = info.seriesGroup || row[COL.SeriesGroup];
      row[COL.Segment] = info.segment || row[COL.Segment];
      row[COL.CPUSegment] = info.cpuSegment || row[COL.CPUSegment];
      row[COL.CPU] = info.cpu || row[COL.CPU];
      row[COL.RAM] = info.ram || row[COL.RAM];
      row[COL.SSD] = info.ssd || row[COL.SSD];
      row[COL.Screen] = info.screen || row[COL.Screen];
      row[COL.GPU] = info.gpu || row[COL.GPU];
      row[COL.VRAM] = info.vram || row[COL.VRAM];
      row[COL.PartNo] = info.partNumber || row[COL.PartNo];
      row[COL.FocusModel] = info.focusModel || row[COL.FocusModel];
    } else {
      // KHONG khop tab Part # -> chuoi tho cua retailer nam nguyen trong
      // data.csv, nen cung mot con chip hien ra hang chuc dang khac nhau tren
      // dashboard. Chuan hoa lai o day.
      //
      // CHI chay o nhanh nay: dong da khop Part # dung gia tri Phuc go tay,
      // khong duoc dung vao.
      //
      // Luon giu gia tri cu lam lop cuoi: luat khong doc duoc thi de nguyen,
      // KHONG de trong — mat du lieu te hon hien thi xau.
      const before = [row[COL.SeriesGroup], row[COL.Segment], row[COL.CPUSegment], row[COL.CPU], row[COL.RAM], row[COL.SSD], row[COL.GPU]].join('\u0000');

      // Products Family (cot Segment) + Series Group: guessSegment() doi chieu
      // ten SP voi danh sach dong may trong tab Segment, KHONG can khop Part #.
      // Truoc day chi dien khi khop Part #, nen 937 dong An Phat trong tron cot
      // "Products Family" tren dashboard — trong khi Google Sheet lai co, vi
      // multi_dealer_scraper co goi matchSegment. Hai noi lech nhau.
      if (!row[COL.Segment]) {
        const g = guessSegment(row[COL.SKU]);
        if (g) {
          row[COL.Segment] = g.segment;
          if (!row[COL.SeriesGroup]) row[COL.SeriesGroup] = g.seriesGroup || '';
        }
      }

      const n = normalizeCpu(row[COL.CPU] || '');
      if (n.confidence === 'full') row[COL.CPU] = n.cpu;
      if (!row[COL.CPUSegment] && n.segment) row[COL.CPUSegment] = n.segment;
      row[COL.RAM] = K.ram(row[COL.RAM] || '') || row[COL.RAM];
      row[COL.SSD] = K.ssd(row[COL.SSD] || '') || row[COL.SSD];
      row[COL.GPU] = K.gpu(row[COL.GPU] || '') || row[COL.GPU];
      const after = [row[COL.SeriesGroup], row[COL.Segment], row[COL.CPUSegment], row[COL.CPU], row[COL.RAM], row[COL.SSD], row[COL.GPU]].join('\u0000');
      if (before !== after) normalized++;
    }
    outRows.push(row);
  }
  debugLog(`Re-enrich xong: ${matched}/${dataRows.length} dòng khớp Part # (loại ${skippedEOL} SP EOL), ${normalized} dòng được chuẩn hoá spec.`);

  const csv = [header, ...outRows].map(r => r.map(csvEscape).join(',')).join('\n') + '\n';
  fs.writeFileSync(DATA_CSV_PATH, csv, 'utf8');
  debugLog(`✅ Đã ghi lại ${DATA_CSV_PATH} — ${outRows.length + 1} dòng (kể cả header), ${(csv.length / 1024).toFixed(0)} KB`);
}

main().catch(err => {
  debugLog(`💥 Lỗi: ${err && err.stack ? err.stack : String(err)}`);
  process.exitCode = 1;
});
