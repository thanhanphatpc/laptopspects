/**
 * enrich-and-export.js
 *
 * 03/08/2026: Thay thế hoàn toàn luồng "Dailly SRP Tracking" (QUERY + 5 MAP
 * formula sống, đang bị nghẽn — mọi thao tác đọc UNFORMATTED_VALUE trên
 * spreadsheet này đều timeout, kể cả RAW DATA). Script này:
 *
 *   1. Đọc data VỪA SCRAPE từ artifact JSON (không qua Sheets — giống 1:1
 *      logic runCombineMode() trong multi_dealer_scraper.js).
 *   2. Đọc Part# / Segment / Key Focus model qua Sheets API bằng
 *      valueRenderOption='FORMULA' — KHÔNG kích hoạt tính toán nên luôn
 *      nhanh, dù bên trong có formula nặng hay không (đã xác nhận qua
 *      diagnose-sheet.js chạy ổn định nhiều lần).
 *   3. Tự tính enrichment (Vendor/SeriesGroup/Segment/CPUSegment/Part#/
 *      Focus Model) bằng JS, port lại 1:1 logic từ các formula gốc:
 *        - Part#!K = 'EOL'  → loại khỏi kết quả (= WHERE V<>'EOL' cũ)
 *        - Segment (Part#!M): nếu ô đó vẫn là formula sống → tự đoán bằng
 *          cách tìm tên Segment!C dài nhất xuất hiện trong tên SP (thay
 *          cho BYROW+REGEXMATCH+SORT+LEN). Nếu Phuc đã ghi đè tay (không
 *          phải formula) → dùng nguyên giá trị đó.
 *        - Series Group (Part#!L): map Segment → Segment!D (exact match),
 *          hoặc dùng giá trị ghi đè tay nếu có.
 *        - Focus Model: map Part# → Key Focus model!D, default 'No'.
 *   4. Ghi thẳng ra dashboard/data.csv (rolling ~35 ngày), KHÔNG ghi lại
 *      vào Google Sheets — dashboard đọc file này trực tiếp.
 *
 * Best-effort: lỗi ở đâu thì log ra dashboard/.enrich-debug.log và dừng,
 * không đụng data.csv cũ, không làm fail job chính.
 */
const fs = require('fs');
const path = require('path');
const os = require('os');
const { google } = require('googleapis');
// Chuan hoa spec cho duong lui (khi SP khong khop tab Part #) — 02/09/2026.
const { normalizeCpu } = require('./spec_normalize.js');
const { K } = require('./spec_dictionary.js');

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const CREDS_PATH = path.join(os.tmpdir(), 'enrich-gcreds.json');
const ARTIFACT_DIR = process.env.COMBINE_INPUT_DIR || './scrape-artifacts';
const DATA_CSV_PATH = path.join(__dirname, 'dashboard', 'data.csv');
const DEBUG_LOG_PATH = path.join(__dirname, 'dashboard', '.enrich-debug.log');
const REQ_TIMEOUT_MS = 25000;
const MAX_HISTORY_DAYS = 35; // dashboard chỉ cần tối đa 30 ngày, dư 5 ngày

function debugLog(msg) {
  const line = `[${new Date().toISOString()}] ${msg}\n`;
  try { fs.appendFileSync(DEBUG_LOG_PATH, line); } catch (_) {}
  console.log(msg);
}

function formatDate(d) {
  return `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;
}
function formatTime(d) {
  return `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
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
function parseDMY(s) {
  const [d, m, y] = String(s).split('/').map(Number);
  return new Date(y, m - 1, d).getTime();
}

// CSV parser đơn giản (RFC4180 cơ bản, có quote) — đủ dùng cho file
// data.csv do chính mình sinh ra.
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

function listJsonFilesRecursive(dir) {
  let results = [];
  let entries;
  try { entries = fs.readdirSync(dir, { withFileTypes: true }); }
  catch (e) { debugLog(`⚠ Không đọc được ${dir}: ${e.message}`); return results; }
  for (const entry of entries) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) results = results.concat(listJsonFilesRecursive(full));
    else if (entry.name.endsWith('.json')) results.push(full);
  }
  return results;
}

async function readTabFormula(sheets, tabName, lastCol) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${tabName}'!A1:${lastCol}`,
    valueRenderOption: 'FORMULA',
  }, { timeout: REQ_TIMEOUT_MS });
  return res.data.values || [];
}

async function main() {
  try { fs.writeFileSync(DEBUG_LOG_PATH, ''); } catch (_) {}
  debugLog('--- Bắt đầu enrich-and-export.js ---');

  // 1) Đọc data vừa scrape từ artifact
  const files = listJsonFilesRecursive(ARTIFACT_DIR);
  let allProducts = [];
  for (const f of files) {
    try {
      const arr = JSON.parse(fs.readFileSync(f, 'utf8'));
      allProducts = allProducts.concat(arr);
    } catch (e) {
      debugLog(`⚠ Lỗi đọc ${f}: ${e.message}`);
    }
  }
  debugLog(`Đọc được ${allProducts.length} SP từ ${files.length} file artifact (${ARTIFACT_DIR})`);
  if (allProducts.length === 0) {
    debugLog('⚠ 0 SP — bỏ qua, giữ data.csv cũ.');
    return;
  }

  if (!SPREADSHEET_ID || !process.env.GOOGLE_CREDENTIALS) {
    debugLog('⚠ Thiếu SPREADSHEET_ID/GOOGLE_CREDENTIALS — bỏ qua.');
    return;
  }

  fs.writeFileSync(CREDS_PATH, process.env.GOOGLE_CREDENTIALS);
  const auth = new google.auth.GoogleAuth({
    keyFile: CREDS_PATH,
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
  });
  const sheets = google.sheets({ version: 'v4', auth });

  // FIX 19/09/2026 [fork thanhanphatpc/laptopspects]: Sheet moi tao KHONG co
  // 2 tab "Part #" / "Segment" (thuoc template phan loai san pham thu cong
  // cua repo goc, chua duoc thiet lap o day). Truoc day loi doc tab (VD:
  // "Unable to parse range: 'Part #'!A1:AE") se NEM LOI va lam ca ham main()
  // dung ngay lap tuc, khong ghi duoc data.csv nao ca (du step co
  // continue-on-error). Gio bat loi rieng cho tung tab: neu tab khong ton
  // tai/khong doc duoc thi coi nhu rong ([]) va CHAY TIEP — cac dong ben
  // duoi da tu dong fallback ve du lieu tho luc scrape (p.cpu/p.gpu/p.ram/...)
  // khi partMap khong co SP do, nen dashboard van co gia/CPU/RAM binh
  // thuong, chi rieng Segment/Series Group se de trong cho toi khi tao tab.
  let partRows = [];
  try {
    debugLog('Đọc Part # (FORMULA render)...');
    partRows = await readTabFormula(sheets, 'Part #', 'AE');
    debugLog(`  ${Math.max(0, partRows.length - 1)} dòng`);
  } catch (e) {
    debugLog(`⚠ Không đọc được tab "Part #" (${e.message}) — bỏ qua, dùng dữ liệu thô lúc scrape thay thế.`);
  }

  let segmentRows = [];
  try {
    debugLog('Đọc Segment (FORMULA render)...');
    segmentRows = await readTabFormula(sheets, 'Segment', 'D');
    debugLog(`  ${Math.max(0, segmentRows.length - 1)} dòng`);
  } catch (e) {
    debugLog(`⚠ Không đọc được tab "Segment" (${e.message}) — bỏ qua, Series Group/Segment sẽ để trống.`);
  }

  fs.unlinkSync(CREDS_PATH);

  // 2) Segment!C (tên dòng SP) -> Series Group
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

  // 3) Part# tab -> spec đầy đủ. FIX 06/08/2026: tra theo TÊN CỘT (giống
  // buildPartLookupMap() trong multi_dealer_scraper.js) — KHÔNG dùng vị trí
  // cố định r[1],r[2]... nữa, vì Phuc có thể chèn/sắp lại cột tuỳ ý. Đồng
  // thời KHÔNG còn đọc tab "Key Focus model" riêng (đã bỏ, không còn tồn tại
  // — Focus Model giờ là 1 CỘT ngay trong tab "Part #", đọc trực tiếp ở đây).
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
    debugLog(`⚠ Part# thiếu cột bắt buộc (Product Name hoặc Focus Model) — kiểm tra lại header tab.`);
  }

  const partMap = new Map();
  // FIX 26/08/2026 [Vendor "MacBook" vs "Macbook" tach thanh 2 nhom]:
  // Cot Vendor lay tu tab "Part #" khi khop ten SP, nhung khi SP CHUA co
  // trong tab Part# thi rot ve p.brand do scraper tu suy ra (detectBrand).
  // Hai nguon nay dung chinh ta khac nhau -> vd "MacBook Air 15 M2 2023"
  // (Part# rong) ra Vendor "MacBook" trong khi 46 SP MacBook khac ra
  // "Macbook" => dashboard hien 2 vendor rieng. Xay ra deu moi ngay 22-26/08.
  //
  // Sua tan goc thay vi dien tay Part# cho 1 SKU: gom TU DIEN Vendor chuan
  // tu chinh tab Part# (nguon su that duy nhat), roi ep p.brand ve dung
  // chinh ta do khi khop khong phan biet hoa/thuong. Tu dong dung cho moi
  // brand khac trong tuong lai, khong hardcode danh sach.
  const vendorCanon = new Map(); // vendor viet thuong -> vendor dung chinh ta trong Part#
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

    const vendorRaw = String((iVendor !== undefined && r[iVendor]) || '').trim();
    if (vendorRaw && !vendorCanon.has(vendorRaw.toLowerCase())) {
      vendorCanon.set(vendorRaw.toLowerCase(), vendorRaw);
    }

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
  debugLog(`Vendor chuan tu Part#: ${vendorCanon.size} gia tri -> ${JSON.stringify([...vendorCanon.values()])}`);

  // Ep brand do scraper suy ra ve dung chinh ta cua tab Part#. Neu khong khop
  // gia tri nao (brand hoan toan moi) thi giu nguyen — khong bia them.
  const canonVendor = (brand) => {
    const b = String(brand || '').trim();
    if (!b) return '';
    return vendorCanon.get(b.toLowerCase()) || b;
  };

  // 5) Tính enrichment, loại EOL
  const today = new Date();
  const dateStr = formatDate(today);
  const timeStr = formatTime(today);
  const HEADER = ['Date', 'Hour', 'Dealers', 'SKU', 'SRP', 'Promotion Price', 'Change', 'Sold', 'Rate', 'Vendor', 'Series Group', 'Segment', 'CPU Segment', 'CPU', 'RAM', 'SSD', 'Screen', 'GPU', 'V-RAM', 'Weight', 'Products link', 'Part #', 'Focus Model', 'Status'];

  let skippedEOL = 0;
  const newRows = [];
  for (const p of allProducts) {
    const info = partMap.get(p.name);
    if (info && info.status === 'EOL') { skippedEOL++; continue; }

    const partNumber = (info && info.partNumber) || '';

    // CHUAN HOA SPEC — them 02/09/2026.
    // Chi ap dung cho DUONG LUI (khi khong khop tab Part #). Gia tri lay tu
    // tab Part # la do Phuc go, KHONG dung vao.
    // Truoc day duong lui do thang chuoi tho cua retailer vao data.csv, nen
    // cung mot con chip hien ra hang chuc dang khac nhau tren dashboard
    // ("NVIDIA GeForce RTX 3050 6GB GDDR6, Boost Clock 1732MHz, TGP 95W"...).
    // Rieng An Phat co 937 dong deu di duong lui vi chua co trong tab Part #.
    //
    // Luon giu chuoi tho lam lop cuoi: neu luat khong doc duoc thi de nguyen,
    // KHONG de trong — mat du lieu con te hon hien thi xau.
    const nCpu = normalizeCpu(p.cpu || '');
    const cpuVal = (info && info.cpu) || (nCpu.confidence === 'full' ? nCpu.cpu : '') || p.cpu || '';
    const cpuSegVal = (info && info.cpuSegment) || nCpu.segment || '';
    const gpuVal = (info && info.gpu) || K.gpu(p.gpu || '') || p.gpu || '';
    const ramVal = (info && info.ram) || K.ram(p.ram || '') || p.ram || '';
    const ssdVal = (info && info.ssd) || K.ssd(p.storage || '') || p.storage || '';

    // Products Family (Segment) + Series Group cho dong KHONG khop tab Part #.
    // guessSegment() doi chieu ten SP voi danh sach dong may trong tab Segment,
    // khong can Part #.
    // LOI 03/09: em chi sua cho nay trong reenrich-dashboard-csv.js ma QUEN
    // file nay — ma duong scrape lai chay file nay. Ket qua: chay scrape xong
    // cot Products Family cua An Phat van trong. Hai file phai sua CUNG NHAU.
    const guessed = info ? null : guessSegment(p.name);
    const segVal = (info && info.segment) || (guessed && guessed.segment) || '';
    const sgVal  = (info && info.seriesGroup) || (guessed && guessed.seriesGroup) || '';

    // FIX 14/08/2026: SRP fallback — khi website bỏ gạch giá gốc, origPrice=0
    // nhưng salePrice vẫn có → dùng salePrice làm SRP để không để trống cột SRP.
    // origPrice=0 && salePrice>0: website bán 1 giá, không có KM hiển thị.
    const srp   = p.origPrice || p.salePrice || '';
    const promo = p.salePrice || '';
    newRows.push([
      dateStr, timeStr, p.dealer, p.name,
      srp, promo, p.discount || '',
      p.sold || '', p.rating || '',
      (info && info.vendor) || canonVendor(p.brand) || '',
      sgVal,
      segVal,
      cpuSegVal,
      cpuVal,
      ramVal,
      ssdVal,
      (info && info.screen) || p.screen || '',
      gpuVal,
      (info && info.vram) || '',
      p.weight || '',
      p.link || '',
      partNumber,
      (info && info.focusModel) || 'No',
      p.stockStatus || '',
    ]);
  }
  debugLog(`Tính xong ${newRows.length} dòng enriched (loại ${skippedEOL} SP EOL)`);

  // 6) Gộp vào data.csv hiện có, bỏ dòng cũ của HÔM NAY (tránh trùng nếu
  //    chạy lại), cắt còn tối đa MAX_HISTORY_DAYS ngày gần nhất.
  let existingRows = [];
  if (fs.existsSync(DATA_CSV_PATH)) {
    const raw = fs.readFileSync(DATA_CSV_PATH, 'utf8');
    const parsed = parseCsvSimple(raw);
    existingRows = parsed.slice(1).filter(r => r[0] !== dateStr);
    debugLog(`data.csv cũ: ${parsed.length - 1} dòng, giữ lại ${existingRows.length} dòng (loại dòng ${dateStr} nếu có, sẽ thêm bản mới)`);
  } else {
    debugLog('Chưa có data.csv — tạo mới.');
  }

  // 6b) FIX 18/08/2026 — CHẶN GIÁ BẤT THƯỜNG (lớp bảo vệ CHUNG).
  // Sự cố cụ thể (6 SKU MBW bị ghi 1.590.000đ do parse trúng giá phụ kiện
  // trong carousel "Sản phẩm liên quan") đã sửa tận gốc ở
  // multi_dealer_scraper.js. Nhưng bài học chung là: MỘT selector sai ở bất
  // kỳ dealer nào cũng có thể bơm giá rác vào dashboard mà không ai biết —
  // đây đã là lần thứ hai. Nên chặn thêm ở đây, độc lập với scraper:
  // so giá hôm nay với TRUNG VỊ của CHÍNH SKU đó trong 7 ngày gần nhất; lệch
  // quá 55% thì bỏ giá (giữ dòng, để trống giá → dashboard hiện badge "no
  // data since" thay vì vẽ mức giảm 93% sai sự thật).
  // Ngưỡng 55% chọn có chủ đích: khuyến mãi laptop thực tế hiếm khi vượt
  // 30-40%, nên 55% gần như chắc chắn là lỗi parse, không phải giá thật.
  const PRICE_DEVIATION_MAX = 0.55;
  const PRICE_HISTORY_DAYS = 7;
  const histDates = [...new Set(existingRows.map(r => r[0]))]
    .sort((a, b) => parseDMY(b) - parseDMY(a))
    .slice(0, PRICE_HISTORY_DAYS);
  const histSet = new Set(histDates);
  const priceHistory = new Map(); // "dealer|sku" -> [giá...]
  for (const r of existingRows) {
    if (!histSet.has(r[0])) continue;
    const p = parseInt(String(r[5] || r[4] || '').replace(/\D/g, '')) || 0;
    if (!p) continue;
    const k = `${r[2]}|${r[3]}`;
    if (!priceHistory.has(k)) priceHistory.set(k, []);
    priceHistory.get(k).push(p);
  }
  let outlierCount = 0;
  for (const row of newRows) {
    const today = parseInt(String(row[5] || row[4] || '').replace(/\D/g, '')) || 0;
    if (!today) continue;
    const hist = priceHistory.get(`${row[2]}|${row[3]}`);
    if (!hist || hist.length < 3) continue; // chưa đủ lịch sử để kết luận
    const sorted = [...hist].sort((a, b) => a - b);
    const mid = Math.floor(sorted.length / 2);
    const median = sorted.length % 2 ? sorted[mid] : Math.round((sorted[mid - 1] + sorted[mid]) / 2);
    if (!median) continue;
    const dev = Math.abs(today - median) / median;
    if (dev > PRICE_DEVIATION_MAX) {
      outlierCount++;
      if (outlierCount <= 15) {
        debugLog(`  ⚠ Giá bất thường: ${row[2]} | ${String(row[3]).substring(0, 55)} — ${today.toLocaleString('vi')}đ vs trung vị ${median.toLocaleString('vi')}đ (lệch ${(dev * 100).toFixed(0)}%) → bỏ giá`);
      }
      row[4] = ''; row[5] = ''; row[6] = '';
    }
  }
  if (outlierCount > 0) {
    debugLog(`⚠ Đã loại giá của ${outlierCount} dòng vì lệch > ${PRICE_DEVIATION_MAX * 100}% so với trung vị ${PRICE_HISTORY_DAYS} ngày của chính SKU đó${outlierCount > 15 ? ' (chỉ log 15 dòng đầu)' : ''}`);
  } else {
    debugLog(`✅ Không có giá bất thường (kiểm tra ${newRows.length} dòng vs trung vị ${PRICE_HISTORY_DAYS} ngày)`);
  }

  let allRows = existingRows.concat(newRows);
  const uniqueDates = [...new Set(allRows.map(r => r[0]))];
  uniqueDates.sort((a, b) => parseDMY(a) - parseDMY(b));
  if (uniqueDates.length > MAX_HISTORY_DAYS) {
    const keepDates = new Set(uniqueDates.slice(-MAX_HISTORY_DAYS));
    allRows = allRows.filter(r => keepDates.has(r[0]));
    debugLog(`Cắt còn ${MAX_HISTORY_DAYS} ngày gần nhất (${allRows.length} dòng)`);
  }

  const csv = [HEADER, ...allRows].map(row => row.map(csvEscape).join(',')).join('\n') + '\n';
  fs.mkdirSync(path.dirname(DATA_CSV_PATH), { recursive: true });
  fs.writeFileSync(DATA_CSV_PATH, csv, 'utf8');
  debugLog(`✅ Đã ghi ${DATA_CSV_PATH} — ${allRows.length + 1} dòng (kể cả header), ${(csv.length / 1024).toFixed(0)} KB`);
}

main().catch(err => {
  debugLog(`💥 Lỗi: ${err && err.stack ? err.stack : String(err)}`);
  process.exitCode = 0;
});
