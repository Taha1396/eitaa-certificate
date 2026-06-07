<!DOCTYPE html>
<html lang="fa" dir="rtl">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=yes, shrink-to-fit=no, viewport-fit=cover">
  <title>پنل صدور گواهی</title>
  
  <!-- SDK ایتا - حتماً قبل از همه -->
  <script src="https://developer.eitaa.com/eitaa-web-app.js"></script>
  
  <!-- فونت وزیرمتن -->
  <link href="https://cdn.jsdelivr.net/gh/rastikerdar/vazirmatn-font@v33.003/dist/font-face.css" rel="stylesheet">
  
  <!-- کتابخونه‌ها -->
  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.min.js"></script>
  
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; font-family: 'Vazirmatn', 'Tahoma', sans-serif; }
    body { 
      background: #f0f2f5; 
      display: flex; 
      flex-direction: column; 
      align-items: center; 
      padding: 16px; 
      padding-bottom: 40px; 
      min-height: 100vh; 
      transition: background 0.3s;
    }
    
    /* تم تاریک */
    body.dark { background: #1a1a2e; color: #e0e0e0; }
    body.dark .container { background: #16213e; }
    body.dark h2 { color: #e0e0e0; }
    body.dark .upload-area { background: #1a1a2e; border-color: #2a2a4a; }
    body.dark .control-card { background: #1a1a2e; border-color: #2a2a4a; }
    body.dark .staff-table th { background: #1a1a2e; }
    body.dark .staff-table td { border-color: #2a2a4a; }
    body.dark .input-row input { background: #1a1a2e; border-color: #2a2a4a; color: #e0e0e0; }
    body.dark .btn { background: #1a1a2e; color: #e0e0e0; border-color: #2a2a4a; }
    body.dark .stats { background: #1a1a2e; }
    body.dark .gcard { background: #1a1a2e; }
    body.dark .gcard-info { background: #0f3460; }
    
    .container { width: 100%; max-width: 550px; background: white; border-radius: 24px; box-shadow: 0 20px 40px rgba(0,0,0,0.08); padding: 20px 16px; display: flex; flex-direction: column; gap: 16px; }
    h2 { font-size: 1.5rem; font-weight: 700; color: #1e293b; text-align: center; }

    /* هدر برنامک */
    .app-header { display: flex; justify-content: space-between; align-items: center; padding: 8px 0; }
    .close-btn { 
      background: #fee2e2; color: #dc2626; border: none; border-radius: 10px; 
      padding: 8px 14px; font-weight: 700; cursor: pointer; font-size: 0.8rem;
      display: flex; align-items: center; gap: 4px;
    }
    .close-btn:active { background: #fecaca; }

    /* Upload Areas */
    .upload-area { background: #f8fafc; border: 2px dashed #cbd5e1; border-radius: 16px; padding: 14px; display: flex; flex-direction: column; align-items: center; gap: 6px; cursor: pointer; text-align: center; transition: 0.2s; font-size: 0.9rem; }
    .upload-area:active { background: #f1f5f9; border-color: #64748b; }
    .upload-area.bg { border-color: #3b82f6; background: #eff6ff; }
    .upload-area.font { border-color: #eab308; background: #fefce8; color: #854d0e; }
    .upload-area.stamp { border-color: #a855f7; background: #faf5ff; color: #6b21a8; }
    .upload-area.excel { border-color: #22c55e; background: #f0fdf4; color: #166534; }
    .file-input { display: none; }

    /* Font List */
    .font-list { display: flex; flex-wrap: wrap; gap: 6px; margin-top: 6px; }
    .font-tag { background: #fef3c7; border: 1px solid #f59e0b; border-radius: 20px; padding: 4px 10px; font-size: 0.75rem; font-weight: 600; cursor: pointer; display: flex; align-items: center; gap: 4px; }
    .font-tag.active { background: #f59e0b; color: white; }
    .font-tag .remove-font { margin-right: 4px; color: #dc2626; font-weight: bold; cursor: pointer; }

    /* Manual Input */
    .input-row { display: flex; gap: 10px; flex-wrap: wrap; }
    .input-row input { flex: 1 1 120px; padding: 12px 14px; border: 2px solid #e2e8f0; border-radius: 12px; font-size: 0.9rem; background: #f8fafc; transition: 0.2s; }
    .input-row input:focus { border-color: #2563eb; outline: none; background: white; }
    .add-btn { background: #2563eb; color: white; border: none; border-radius: 12px; padding: 12px 20px; font-weight: 700; font-size: 0.9rem; cursor: pointer; white-space: nowrap; transition: 0.2s; }
    .add-btn:active { background: #1d4ed8; transform: scale(0.96); }

    /* Preview */
    .preview-box { border-radius: 16px; overflow: hidden; border: 2px solid #e2e8f0; touch-action: manipulation; background: #f8fafc; }
    canvas { display: block; width: 100%; height: auto; background: white; }

    /* Controls */
    .controls { display: flex; flex-direction: column; gap: 12px; background: #f8fafc; padding: 16px; border-radius: 20px; }
    .control-card { background: white; border: 2px solid #e2e8f0; border-radius: 14px; padding: 14px; }
    .control-card.name-card { border-color: #3b82f6; }
    .control-card.code-card { border-color: #10b981; }
    .control-card.stamp-card { border-color: #a855f7; }
    .control-header { font-weight: 700; font-size: 0.9rem; margin-bottom: 10px; display: flex; align-items: center; gap: 8px; }
    .btn-row { display: flex; gap: 8px; flex-wrap: wrap; align-items: center; margin-bottom: 8px; }
    .btn { background: white; border: 1px solid #cbd5e1; border-radius: 10px; padding: 10px 14px; font-weight: 600; font-size: 1rem; cursor: pointer; min-width: 44px; text-align: center; user-select: none; transition: 0.15s; }
    .btn:active { background: #e2e8f0; transform: scale(0.95); }
    .btn.rotate-btn { background: #fef3c7; border-color: #f59e0b; color: #92400e; }
    .slider-row { display: flex; align-items: center; gap: 10px; }
    .slider-row label { font-size: 0.8rem; font-weight: 600; white-space: nowrap; }
    .slider-row input[type="range"] { flex: 1; accent-color: #2563eb; height: 6px; }
    .slider-row .val { font-weight: 700; font-size: 0.85rem; min-width: 32px; text-align: center; }

    /* Color */
    .color-row { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; background: white; padding: 12px; border-radius: 14px; border: 2px solid #e2e8f0; }
    .color-row label { font-weight: 600; font-size: 0.9rem; }
    .color-dots { display: flex; gap: 8px; }
    .color-dot { width: 30px; height: 30px; border-radius: 50%; border: 2px solid #e2e8f0; cursor: pointer; transition: 0.15s; }
    .color-dot.active { border: 3px solid #2563eb; transform: scale(1.15); box-shadow: 0 0 8px rgba(37,99,235,0.3); }

    /* Table */
    .staff-table { width: 100%; border-collapse: collapse; margin: 8px 0; font-size: 0.85rem; }
    .staff-table th { background: #f1f5f9; padding: 10px 8px; text-align: center; font-weight: 700; border-bottom: 2px solid #e2e8f0; }
    .staff-table td { padding: 8px; border-bottom: 1px solid #e2e8f0; text-align: center; }
    .delete-btn { background: #fee2e2; color: #dc2626; border: none; border-radius: 8px; padding: 6px 12px; font-weight: 700; cursor: pointer; font-size: 0.8rem; }

    /* Generated Cards */
    .generated-section { margin-top: 8px; }
    .generated-grid { display: flex; flex-wrap: wrap; gap: 16px; justify-content: center; margin-top: 12px; }
    .gcard { background: white; border-radius: 16px; box-shadow: 0 8px 24px rgba(0,0,0,0.08); overflow: hidden; width: 100%; max-width: 300px; }
    .gcard canvas { width: 100%; height: auto; display: block; }
    .gcard-info { padding: 10px; text-align: center; font-weight: 600; font-size: 0.8rem; background: #f8fafc; display: flex; flex-direction: column; gap: 6px; }
    .gcard-btn { background: #2563eb; color: white; border: none; border-radius: 8px; padding: 8px 12px; font-weight: 600; font-size: 0.75rem; cursor: pointer; width: 100%; }

    /* Stats & Actions */
    .stats { display: flex; justify-content: space-between; align-items: center; background: #f1f5f9; padding: 10px 14px; border-radius: 14px; font-size: 0.9rem; }
    .action-btn { background: #0f172a; color: white; border: none; border-radius: 16px; padding: 16px; font-weight: 700; font-size: 1.1rem; cursor: pointer; width: 100%; transition: 0.2s; }
    .action-btn:disabled { background: #94a3b8; cursor: not-allowed; }
    .danger-btn { background: #fee2e2; color: #dc2626; border: none; border-radius: 10px; padding: 8px 14px; font-weight: bold; cursor: pointer; font-size: 0.85rem; }
    .info-msg { padding: 12px; border-radius: 12px; font-size: 0.85rem; text-align: center; font-weight: 600; display: none; }
    .progress-container { width: 100%; background: #e2e8f0; border-radius: 10px; overflow: hidden; display: none; margin-top: 4px; }
    .progress-bar { height: 26px; background: linear-gradient(90deg, #2563eb, #7c3aed); width: 0%; transition: width 0.3s; display: flex; align-items: center; justify-content: center; color: white; font-size: 0.8rem; font-weight: bold; }

    /* Footer */
    .footer { text-align: center; padding: 16px 0 8px; color: #94a3b8; font-size: 0.8rem; font-weight: 500; }
    .footer strong { color: #64748b; }

    @media print {
      body { margin: 0; padding: 0; background: white; }
      .container, .info-msg, .progress-container, .generated-section, .footer { display: none !important; }
      .print-area { display: block !important; width: 100%; }
      .print-page { width: 210mm; height: 148mm; page-break-after: always; display: flex; align-items: center; justify-content: center; padding: 10mm; box-sizing: border-box; background: white; }
      .print-page canvas { max-width: 100%; max-height: 100%; object-fit: contain; }
      @page { size: A5 landscape; margin: 0; }
    }
  </style>
</head>
<body>
<div class="container">
  
  <!-- هدر برنامک -->
  <div class="app-header">
    <h2>🎓 پنل صدور گواهی</h2>
    <button class="close-btn" id="closeAppBtn">✕ بستن</button>
  </div>

  <div class="upload-area bg" id="bgUploadBtn"><span>📁 بارگذاری زمینه (تصویر یا PDF)</span><small>تصویر اصلی گواهی یا فایل PDF صفحه اول</small><input type="file" id="bgFile" class="file-input" accept="image/*,application/pdf"></div>
  
  <div class="upload-area font" id="fontUploadBtn">
    <span>🔤 بارگذاری فونت (چندگانه)</span>
    <small>ttf, otf, woff - در مرورگر ذخیره می‌شود</small>
    <input type="file" id="fontFile" class="file-input" accept=".ttf,.otf,.woff,.woff2" multiple>
  </div>
  <div class="font-list" id="fontList"></div>

  <div class="upload-area stamp" id="stampUploadBtn"><span>🏷️ بارگذاری برچسب / مهر</span><small>PNG دوربری شده</small><input type="file" id="stampFile" class="file-input" accept="image/png"></div>
  <div class="upload-area excel" id="excelUploadBtn"><span>📊 بارگذاری لیست اکسل</span><small>ستون ۱: نام | ستون ۲: کد پرسنلی</small><input type="file" id="excelFile" class="file-input" accept=".xlsx,.xls,.csv"></div>
  
  <div class="info-msg" id="infoMsg"></div>

  <div class="preview-box" id="previewBox"><canvas id="previewCanvas"></canvas></div>

  <div class="controls">
    <div class="color-row">
      <label>🎨 رنگ نوشتار:</label>
      <input type="color" id="colorPicker" value="#0f172a" style="width:40px;height:40px;border:none;cursor:pointer;border-radius:10px;">
      <div class="color-dots">
        <span class="color-dot active" style="background:#0f172a;" data-color="#0f172a"></span>
        <span class="color-dot" style="background:#b91c1c;" data-color="#b91c1c"></span>
        <span class="color-dot" style="background:#1e40af;" data-color="#1e40af"></span>
        <span class="color-dot" style="background:#047857;" data-color="#047857"></span>
        <span class="color-dot" style="background:#7c3aed;" data-color="#7c3aed"></span>
        <span class="color-dot" style="background:#000000;" data-color="#000000"></span>
        <span class="color-dot" style="background:#ffffff; border:2px solid #cbd5e1;" data-color="#ffffff"></span>
      </div>
    </div>

    <div class="control-card name-card">
      <div class="control-header">👤 نام و نام خانوادگی</div>
      <div class="btn-row">
        <button class="btn" id="nUp">↑</button><button class="btn" id="nDown">↓</button>
        <button class="btn" id="nLeft">←</button><button class="btn" id="nRight">→</button>
      </div>
      <div class="slider-row">
        <label>سایز:</label><input type="range" id="nSize" min="20" max="200" value="60"><span class="val" id="nSizeVal">60</span>
      </div>
    </div>

    <div class="control-card code-card">
      <div class="control-header">🆔 کد پرسنلی</div>
      <div class="btn-row">
        <button class="btn" id="cUp">↑</button><button class="btn" id="cDown">↓</button>
        <button class="btn" id="cLeft">←</button><button class="btn" id="cRight">→</button>
      </div>
      <div class="slider-row">
        <label>سایز:</label><input type="range" id="cSize" min="20" max="200" value="50"><span class="val" id="cSizeVal">50</span>
      </div>
    </div>

    <div class="control-card stamp-card">
      <div class="control-header">🏷️ برچسب / مهر / لوگو</div>
      <div class="btn-row">
        <button class="btn" id="sUp">↑</button><button class="btn" id="sDown">↓</button>
        <button class="btn" id="sLeft">←</button><button class="btn" id="sRight">→</button>
        <button class="btn rotate-btn" id="sRotL">↺</button><button class="btn rotate-btn" id="sRotR">↻</button>
      </div>
      <div class="slider-row">
        <label>سایز:</label><input type="range" id="sSize" min="30" max="400" value="120"><span class="val" id="sSizeVal">120</span>
      </div>
    </div>
  </div>

  <div style="background:white; border-radius:16px; border:2px solid #e2e8f0; padding:12px;">
    <div class="input-row" style="margin-bottom:12px;">
      <input type="text" id="nameInput" placeholder="نام و نام خانوادگی">
      <input type="text" id="codeInput" placeholder="کد پرسنلی">
      <button class="add-btn" id="addBtn">➕ افزودن</button>
    </div>
    <div class="stats" style="margin-bottom:8px;">
      <span>👥 افراد: <strong id="countDisplay">0</strong> نفر</span>
      <button class="danger-btn" id="clearBtn">🗑️ حذف همه</button>
    </div>
    <div style="max-height:200px; overflow-y:auto;">
      <table class="staff-table" id="staffTable">
        <thead><tr><th>نام و نام خانوادگی</th><th>کد پرسنلی</th><th>حذف</th></tr></thead>
        <tbody></tbody>
      </table>
    </div>
  </div>

  <button class="action-btn" id="generateBtn" style="background:#2563eb;">⚡ تولید گواهی‌ها</button>

  <div class="generated-section" id="generatedSection" style="display:none;">
    <h3 style="text-align:center; color:#1e293b; margin-bottom:8px;">📋 گواهی‌های تولید شده</h3>
    <div class="generated-grid" id="generatedGrid"></div>
    <div style="display:flex; gap:10px; margin-top:12px;">
      <button class="action-btn" id="downloadZipBtn" style="background:#7c3aed;">📦 ذخیره ZIP همه گواهی‌ها</button>
    </div>
  </div>

  <div class="progress-container" id="progressContainer"><div class="progress-bar" id="progressBar">0%</div></div>

  <button class="action-btn" id="savePdfBtn">💾 ذخیره PDF همه گواهی‌ها</button>
  <button class="action-btn" id="printAllBtn" style="background:#1e40af;">🖨️ چاپ همه گواهی‌ها</button>

  <div class="footer">
    🛠️ طراحی شده توسط <strong>عباس آمون</strong>
  </div>
</div>

<div id="printArea" style="display:none;" class="print-area"></div>

<script>
  (() => {
    // ==================== EITAA SDK ====================
    const eitaa = window.Eitaa?.WebApp;
    
    // تنظیم تم ایتا
    if (eitaa) {
      if (eitaa.colorScheme === 'dark') {
        document.body.classList.add('dark');
      }
      // اعلام آمادگی
      eitaa.ready();
      // گسترش به تمام صفحه
      eitaa.expand();
      // دکمه بستن
      document.getElementById('closeAppBtn').onclick = () => eitaa.close();
    } else {
      document.getElementById('closeAppBtn').style.display = 'none';
    }

    // تابع کمکی برای دانلود با SDK ایتا
    function downloadWithEitaa(url, filename) {
      if (eitaa && eitaa.downloadFile) {
        eitaa.downloadFile({ url, file_name: filename });
      } else {
        // fallback برای مرورگر
        const link = document.createElement('a');
        link.download = filename;
        link.href = url;
        link.click();
      }
    }

    // ==================== STATE ====================
    let bgImg = null, stampImg = null, staff = [], generatedCanvases = [];
    let uploadedFonts = [];
    let activeFontFamily = null;
    let nX = 1240, nY = 700, nSize = 60;
    let cX = 1240, cY = 920, cSize = 50;
    let sX = 2100, sY = 1300, sSize = 120, sRot = 0;
    let textColor = '#0f172a';
    const CW = 2480, CH = 1748;
    const DB_NAME = 'CertificateForge';
    const DB_VERSION = 1;
    const STORE_NAME = 'fonts';

    // PDF.js configuration
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';

    // ==================== DOM ====================
    const canvas = document.getElementById('previewCanvas'), ctx = canvas.getContext('2d');
    canvas.width = CW; canvas.height = CH;
    const bgFile = document.getElementById('bgFile'), fontFile = document.getElementById('fontFile');
    const stampFile = document.getElementById('stampFile'), excelFile = document.getElementById('excelFile');
    const infoMsg = document.getElementById('infoMsg'), countDisp = document.getElementById('countDisplay');
    const progressContainer = document.getElementById('progressContainer'), progressBar = document.getElementById('progressBar');
    const savePdfBtn = document.getElementById('savePdfBtn'), printAllBtn = document.getElementById('printAllBtn');
    const printArea = document.getElementById('printArea');
    const nameInput = document.getElementById('nameInput'), codeInput = document.getElementById('codeInput');
    const staffTableBody = document.querySelector('#staffTable tbody');
    const generatedSection = document.getElementById('generatedSection');
    const generatedGrid = document.getElementById('generatedGrid');
    const fontListDiv = document.getElementById('fontList');

    // ==================== HELPERS ====================
    function showMsg(t, ok = true) {
      infoMsg.style.display = 'block'; infoMsg.textContent = t;
      infoMsg.style.background = ok ? '#dcfce7' : '#fee2e2'; infoMsg.style.color = ok ? '#166534' : '#991b1b';
      setTimeout(() => infoMsg.style.display = 'none', 5000);
    }
    function escapeHtml(t) { const d = document.createElement('div'); d.textContent = t; return d.innerHTML; }

    // ==================== PDF to Image ====================
    async function loadPdfAsImage(file) {
      try {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        const page = await pdf.getPage(1);
        const viewport = page.getViewport({ scale: 2.0 });
        const tempCanvas = document.createElement('canvas');
        const context = tempCanvas.getContext('2d');
        tempCanvas.width = CW;
        tempCanvas.height = CH;
        const scaleX = CW / viewport.width;
        const scaleY = CH / viewport.height;
        const scale = Math.max(scaleX, scaleY);
        const scaledViewport = page.getViewport({ scale: scale });
        tempCanvas.width = scaledViewport.width;
        tempCanvas.height = scaledViewport.height;
        const renderContext = { canvasContext: context, viewport: scaledViewport };
        await page.render(renderContext).promise;
        const img = new Image();
        return new Promise((resolve) => {
          img.onload = () => resolve(img);
          img.src = tempCanvas.toDataURL('image/png');
        });
      } catch (err) {
        console.error('PDF load error:', err);
        throw new Error('خطا در خواندن فایل PDF');
      }
    }

    // ==================== BACKGROUND UPLOAD ====================
    async function handleBackgroundFile(file) {
      if (!file) return;
      const fileType = file.type;
      if (fileType.startsWith('image/')) {
        const reader = new FileReader();
        reader.onload = (ev) => {
          const img = new Image();
          img.onload = () => {
            bgImg = img;
            draw();
            showMsg('✅ تصویر زمینه بارگذاری شد');
          };
          img.src = ev.target.result;
        };
        reader.readAsDataURL(file);
      } else if (fileType === 'application/pdf') {
        showMsg('⏳ در حال پردازش فایل PDF...');
        try {
          const pdfImg = await loadPdfAsImage(file);
          bgImg = pdfImg;
          draw();
          showMsg('✅ صفحه اول PDF به عنوان زمینه بارگذاری شد');
        } catch (err) {
          showMsg('❌ خطا در بارگذاری PDF', false);
        }
      } else {
        showMsg('❌ فرمت فایل پشتیبانی نمی‌شود (فقط تصویر یا PDF)', false);
      }
    }

    document.getElementById('bgUploadBtn').onclick = () => bgFile.click();
    bgFile.onchange = e => { const f = e.target.files[0]; if (f) handleBackgroundFile(f); };

    // ==================== IndexedDB ====================
    function openDB() {
      return new Promise((resolve, reject) => {
        const request = indexedDB.open(DB_NAME, DB_VERSION);
        request.onupgradeneeded = (e) => {
          const db = e.target.result;
          if (!db.objectStoreNames.contains(STORE_NAME)) {
            db.createObjectStore(STORE_NAME, { keyPath: 'family' });
          }
        };
        request.onsuccess = (e) => resolve(e.target.result);
        request.onerror = (e) => reject(e.target.error);
      });
    }

    async function saveFontToDB(font) {
      const db = await openDB();
      return new Promise((resolve, reject) => {
        const tx = db.transaction(STORE_NAME, 'readwrite');
        const store = tx.objectStore(STORE_NAME);
        store.put({ family: font.family, name: font.name, dataUrl: font.dataUrl });
        tx.oncomplete = () => resolve();
        tx.onerror = (e) => reject(e.target.error);
      });
    }

    async function deleteFontFromDB(family) {
      const db = await openDB();
      return new Promise((resolve, reject) => {
        const tx = db.transaction(STORE_NAME, 'readwrite');
        const store = tx.objectStore(STORE_NAME);
        store.delete(family);
        tx.oncomplete = () => resolve();
        tx.onerror = (e) => reject(e.target.error);
      });
    }

    async function loadFontsFromDB() {
      const db = await openDB();
      return new Promise((resolve, reject) => {
        const tx = db.transaction(STORE_NAME, 'readonly');
        const store = tx.objectStore(STORE_NAME);
        const request = store.getAll();
        request.onsuccess = () => resolve(request.result);
        request.onerror = (e) => reject(e.target.error);
      });
    }

    // ==================== FONT MANAGEMENT ====================
    function getFont(size) {
      if (activeFontFamily) {
        return `${size}px '${activeFontFamily}', 'Vazirmatn', 'Tahoma'`;
      }
      return `${size}px 'Vazirmatn', 'Tahoma', sans-serif`;
    }

    async function loadAndRegisterFonts() {
      try {
        const fonts = await loadFontsFromDB();
        for (const f of fonts) {
          try {
            const fontFace = new FontFace(f.family, `url(${f.dataUrl})`);
            const loaded = await fontFace.load();
            document.fonts.add(loaded);
            uploadedFonts.push({ name: f.name, family: f.family, dataUrl: f.dataUrl });
          } catch (err) {
            console.error('Error loading font:', f.name, err);
          }
        }
        if (uploadedFonts.length > 0) {
          activeFontFamily = uploadedFonts[uploadedFonts.length - 1].family;
        }
        renderFontList();
        draw();
      } catch (err) {
        console.error('Error loading fonts from DB:', err);
        renderFontList();
        draw();
      }
    }

    function renderFontList() {
      fontListDiv.innerHTML = '';
      if (uploadedFonts.length === 0) {
        fontListDiv.innerHTML = '<small style="color:#64748b;">هیچ فونتی بارگذاری نشده</small>';
        return;
      }
      uploadedFonts.forEach((f, i) => {
        const tag = document.createElement('span');
        tag.className = 'font-tag';
        if (activeFontFamily === f.family) tag.classList.add('active');
        tag.innerHTML = `${f.name} <span class="remove-font" data-idx="${i}">✕</span>`;
        tag.onclick = (e) => {
          if (e.target.classList.contains('remove-font')) return;
          activeFontFamily = activeFontFamily === f.family ? null : f.family;
          renderFontList();
          draw();
        };
        fontListDiv.appendChild(tag);
      });
      document.querySelectorAll('.remove-font').forEach(btn => {
        btn.onclick = async (e) => {
          e.stopPropagation();
          const idx = parseInt(btn.dataset.idx);
          const font = uploadedFonts[idx];
          if (activeFontFamily === font.family) activeFontFamily = null;
          uploadedFonts.splice(idx, 1);
          await deleteFontFromDB(font.family);
          renderFontList();
          draw();
        };
      });
    }

    // ==================== DRAW ====================
    function draw() {
      ctx.clearRect(0, 0, CW, CH);
      if (bgImg) ctx.drawImage(bgImg, 0, 0, CW, CH);
      else { 
        ctx.fillStyle = '#f1f5f9'; ctx.fillRect(0, 0, CW, CH); 
        ctx.fillStyle = '#94a3b8'; ctx.font = 'bold 36px Vazirmatn'; ctx.textAlign = 'center'; 
        ctx.fillText('تصویر زمینه بارگذاری نشده', CW / 2, CH / 2); return; 
      }
      const dn = staff[0]?.name || 'نام و نام خانوادگی';
      const dc = staff[0]?.code || 'کد پرسنلی';
      ctx.save(); ctx.fillStyle = textColor; ctx.textAlign = 'center'; ctx.direction = 'rtl';
      ctx.font = getFont(nSize); ctx.fillText(dn, nX, nY);
      ctx.font = getFont(cSize); ctx.fillText(dc, cX, cY);
      ctx.restore();
      if (stampImg) {
        ctx.save(); ctx.translate(sX, sY); ctx.rotate((sRot * Math.PI) / 180);
        const sh = (stampImg.height / stampImg.width) * sSize;
        ctx.drawImage(stampImg, -sSize / 2, -sh / 2, sSize, sh);
        ctx.restore();
      }
      const m = 118; ctx.strokeStyle = '#2563eb'; ctx.lineWidth = 2; ctx.setLineDash([10, 10]);
      ctx.strokeRect(m, m, CW - 2 * m, CH - 2 * m); ctx.setLineDash([]);
    }

    function createCertificateCanvas(person) {
      const tc = document.createElement('canvas'); tc.width = CW; tc.height = CH; const tctx = tc.getContext('2d');
      tctx.drawImage(bgImg, 0, 0, CW, CH);
      tctx.fillStyle = textColor; tctx.textAlign = 'center'; tctx.direction = 'rtl';
      tctx.font = getFont(nSize); tctx.fillText(person.name, nX, nY);
      tctx.font = getFont(cSize); tctx.fillText(person.code, cX, cY);
      if (stampImg) {
        tctx.save(); tctx.translate(sX, sY); tctx.rotate((sRot * Math.PI) / 180);
        const sh = (stampImg.height / stampImg.width) * sSize;
        tctx.drawImage(stampImg, -sSize / 2, -sh / 2, sSize, sh);
        tctx.restore();
      }
      return tc;
    }

    // ==================== TABLE ====================
    function renderTable() {
      staffTableBody.innerHTML = '';
      staff.forEach((p, i) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td>${escapeHtml(p.name)}</td><td>${escapeHtml(p.code)}</td><td><button class="delete-btn" data-idx="${i}">🗑️</button></td>`;
        staffTableBody.appendChild(tr);
      });
      countDisp.textContent = staff.length;
      document.querySelectorAll('.delete-btn').forEach(b => b.onclick = function() {
        staff.splice(parseInt(this.dataset.idx), 1); renderTable(); draw();
      });
    }

    // ==================== FONT UPLOAD ====================
    document.getElementById('fontUploadBtn').onclick = () => fontFile.click();
    fontFile.onchange = async (e) => {
      const files = Array.from(e.target.files);
      if (!files.length) return;
      let loaded = 0;
      for (const file of files) {
        const reader = new FileReader();
        const dataUrl = await new Promise((resolve) => {
          reader.onload = ev => resolve(ev.target.result);
          reader.readAsDataURL(file);
        });
        const fontName = file.name.replace(/\.[^/.]+$/, '');
        const family = `font_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
        try {
          const fontFace = new FontFace(family, `url(${dataUrl})`);
          const loadedFont = await fontFace.load();
          document.fonts.add(loadedFont);
          const fontObj = { name: fontName, family, dataUrl };
          uploadedFonts.push(fontObj);
          await saveFontToDB(fontObj);
          loaded++;
        } catch (err) {
          console.error('Font load error:', err);
        }
      }
      if (loaded > 0) {
        renderFontList();
        showMsg(`✅ ${loaded} فونت بارگذاری و ذخیره شد`);
        if (!activeFontFamily && uploadedFonts.length > 0) {
          activeFontFamily = uploadedFonts[uploadedFonts.length - 1].family;
          renderFontList();
        }
        draw();
      } else {
        showMsg('❌ هیچ فونتی بارگذاری نشد', false);
      }
      fontFile.value = '';
    };

    // ==================== STAMP UPLOAD ====================
    document.getElementById('stampUploadBtn').onclick = () => stampFile.click();
    stampFile.onchange = e => { 
      const f = e.target.files[0]; if (!f) return; 
      const r = new FileReader(); 
      r.onload = ev => { 
        const i = new Image(); 
        i.onload = () => { stampImg = i; draw(); showMsg('✅ برچسب بارگذاری شد'); }; 
        i.src = ev.target.result; 
      }; 
      r.readAsDataURL(f); 
    };

    // ==================== EXCEL UPLOAD ====================
    document.getElementById('excelUploadBtn').onclick = () => excelFile.click();
    excelFile.onchange = e => {
      const f = e.target.files[0]; if (!f) return; const r = new FileReader();
      r.onload = ev => {
        try {
          const wb = XLSX.read(new Uint8Array(ev.target.result), { type: 'array' });
          const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
          let added = 0;
          for (let i = 1; i < rows.length; i++) { 
            const row = rows[i]; if (!row || row.length < 2) continue; 
            const name = String(row[0]).trim(), code = String(row[1]).trim(); 
            if (name && code && !staff.some(p => p.name === name && p.code === code)) { staff.push({ name, code }); added++; } 
          }
          renderTable(); showMsg(`✅ ${added} نفر اضافه شد (کل: ${staff.length})`); draw();
        } catch (err) { showMsg('❌ خطا در خواندن اکسل', false); }
      };
      r.readAsArrayBuffer(f); excelFile.value = '';
    };

    // ==================== MANUAL ADD ====================
    document.getElementById('addBtn').onclick = () => {
      const n = nameInput.value.trim(), c = codeInput.value.trim();
      if (!n || !c) return showMsg('⚠️ نام و کد را وارد کنید', false);
      if (staff.some(p => p.name === n && p.code === c)) return showMsg('⚠️ تکراری است', false);
      staff.push({ name: n, code: c }); nameInput.value = ''; codeInput.value = ''; renderTable(); draw();
    };

    // ==================== CLEAR ====================
    document.getElementById('clearBtn').onclick = () => { 
      if (!staff.length) return; 
      if (confirm(`حذف ${staff.length} نفر؟`)) { 
        staff = []; renderTable(); printArea.innerHTML = ''; generatedGrid.innerHTML = ''; 
        generatedSection.style.display = 'none'; generatedCanvases = []; draw(); 
      } 
    };

    // ==================== COLOR ====================
    document.getElementById('colorPicker').oninput = e => { textColor = e.target.value; updateDots(); draw(); };
    document.querySelectorAll('.color-dot').forEach(d => d.onclick = () => { 
      textColor = d.dataset.color; document.getElementById('colorPicker').value = textColor; updateDots(); draw(); 
    });
    function updateDots() { document.querySelectorAll('.color-dot').forEach(d => d.classList.toggle('active', d.dataset.color === textColor)); }

    // ==================== CONTROLS ====================
    document.getElementById('nSize').oninput = e => { nSize = +e.target.value; document.getElementById('nSizeVal').textContent = nSize; draw(); };
    ['nUp', 'nDown', 'nLeft', 'nRight'].forEach((id, i) => document.getElementById(id).onclick = () => { 
      const d = 25; i === 0 ? nY -= d : i === 1 ? nY += d : i === 2 ? nX -= d : nX += d; draw(); 
    });
    document.getElementById('cSize').oninput = e => { cSize = +e.target.value; document.getElementById('cSizeVal').textContent = cSize; draw(); };
    ['cUp', 'cDown', 'cLeft', 'cRight'].forEach((id, i) => document.getElementById(id).onclick = () => { 
      const d = 25; i === 0 ? cY -= d : i === 1 ? cY += d : i === 2 ? cX -= d : cX += d; draw(); 
    });
    document.getElementById('sSize').oninput = e => { sSize = +e.target.value; document.getElementById('sSizeVal').textContent = sSize; draw(); };
    ['sUp', 'sDown', 'sLeft', 'sRight'].forEach((id, i) => document.getElementById(id).onclick = () => { 
      const d = 25; i === 0 ? sY -= d : i === 1 ? sY += d : i === 2 ? sX -= d : sX += d; draw(); 
    });
    document.getElementById('sRotL').onclick = () => { sRot -= 15; draw(); };
    document.getElementById('sRotR').onclick = () => { sRot += 15; draw(); };

    // ==================== TOUCH ====================
    let active = null;
    canvas.ontouchstart = e => { 
      e.preventDefault(); const t = e.touches[0], r = canvas.getBoundingClientRect(), 
      sx = CW / r.width, sy = CH / r.height, cx = (t.clientX - r.left) * sx, cy = (t.clientY - r.top) * sy; 
      if (Math.abs(cx - nX) < 100 && Math.abs(cy - nY) < 100) active = 'n'; 
      else if (Math.abs(cx - cX) < 100 && Math.abs(cy - cY) < 100) active = 'c'; 
      else if (stampImg && Math.abs(cx - sX) < 120 && Math.abs(cy - sY) < 120) active = 's'; 
    };
    canvas.ontouchmove = e => { 
      e.preventDefault(); if (!active) return; 
      const t = e.touches[0], r = canvas.getBoundingClientRect(), 
      sx = CW / r.width, sy = CH / r.height, cx = (t.clientX - r.left) * sx, cy = (t.clientY - r.top) * sy; 
      if (active === 'n') { nX = cx; nY = cy; } 
      else if (active === 'c') { cX = cx; cY = cy; } 
      else if (active === 's') { sX = cx; sY = cy; } 
      draw(); 
    };
    canvas.ontouchend = () => active = null;

    // ==================== GENERATE ====================
    document.getElementById('generateBtn').onclick = () => {
      if (!bgImg) return showMsg('❌ تصویر زمینه الزامی است', false);
      if (!staff.length) return showMsg('❌ لیست خالی است', false);
      generatedGrid.innerHTML = '';
      generatedCanvases = [];
      staff.forEach((p, i) => {
        const tc = createCertificateCanvas(p);
        generatedCanvases.push({ canvas: tc, name: p.name, code: p.code });
        const card = document.createElement('div');
        card.className = 'gcard';
        const preview = document.createElement('canvas');
        preview.width = CW; preview.height = CH;
        preview.style.width = '100%'; preview.style.height = 'auto';
        preview.getContext('2d').drawImage(tc, 0, 0);
        card.appendChild(preview);
        const info = document.createElement('div');
        info.className = 'gcard-info';
        info.innerHTML = `<span>${i + 1}. ${escapeHtml(p.name)} - ${escapeHtml(p.code)}</span>`;
        const dlBtn = document.createElement('button');
        dlBtn.className = 'gcard-btn';
        dlBtn.textContent = '💾 ذخیره عکس این گواهی';
        dlBtn.onclick = () => {
          const dataUrl = tc.toDataURL('image/png');
          const filename = `گواهی_${p.name.replace(/\s/g, '_')}_${p.code}.png`;
          // استفاده از SDK ایتا برای دانلود
          downloadWithEitaa(dataUrl, filename);
        };
        info.appendChild(dlBtn);
        card.appendChild(info);
        generatedGrid.appendChild(card);
      });
      generatedSection.style.display = 'block';
      generatedGrid.scrollIntoView({ behavior: 'smooth' });
      showMsg(`✅ ${staff.length} گواهی تولید شد`);
    };

    // ==================== DOWNLOAD ZIP ====================
    document.getElementById('downloadZipBtn').onclick = async () => {
      if (!generatedCanvases.length) return showMsg('❌ ابتدا گواهی‌ها را تولید کنید', false);
      const zip = new JSZip();
      generatedCanvases.forEach((g, i) => {
        const data = g.canvas.toDataURL('image/png').split(',')[1];
        zip.file(`${i + 1}_${g.name.replace(/\s/g, '_')}_${g.code}.png`, data, { base64: true });
      });
      const content = await zip.generateAsync({ type: 'blob' });
      const url = URL.createObjectURL(content);
      downloadWithEitaa(url, 'گواهی_های_وبینار.zip');
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      showMsg('✅ فایل ZIP همه گواهی‌ها ذخیره شد');
    };

    // ==================== PDF ====================
    savePdfBtn.onclick = async () => {
      if (!bgImg) return showMsg('❌ تصویر زمینه الزامی است', false);
      if (!staff.length) return showMsg('❌ لیست خالی است', false);
      const total = staff.length; savePdfBtn.disabled = true; savePdfBtn.textContent = '⏳ در حال ساخت PDF...';
      progressContainer.style.display = 'block'; progressBar.style.width = '0%'; progressBar.textContent = '0%';
      try {
        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a5' });
        const pw = 210, ph = 148, margin = 10;
        for (let i = 0; i < total; i += 10) {
          const end = Math.min(i + 10, total);
          for (let j = i; j < end; j++) {
            const tc = createCertificateCanvas(staff[j]);
            if (j > 0) pdf.addPage();
            pdf.addImage(tc.toDataURL('image/jpeg', 0.95), 'JPEG', margin, margin, pw - 2 * margin, ph - 2 * margin);
            tc.remove();
          }
          progressBar.style.width = `${Math.round((end / total) * 100)}%`;
          progressBar.textContent = `${Math.round((end / total) * 100)}%`;
          await new Promise(r => setTimeout(r, 50));
        }
        const pdfBlob = pdf.output('blob');
        const url = URL.createObjectURL(pdfBlob);
        downloadWithEitaa(url, `گواهی_وبینار_${total}_نفره.pdf`);
        setTimeout(() => URL.revokeObjectURL(url), 1000);
        showMsg(`✅ PDF با ${total} گواهی ذخیره شد`);
      } catch (err) { showMsg('❌ خطا در ساخت PDF', false); } finally {
        savePdfBtn.disabled = false; savePdfBtn.textContent = '💾 ذخیره PDF همه گواهی‌ها';
        setTimeout(() => progressContainer.style.display = 'none', 2000);
      }
    };

    // ==================== PRINT ====================
    printAllBtn.onclick = () => {
      if (!bgImg) return showMsg('❌ تصویر زمینه الزامی است', false);
      if (!staff.length) return showMsg('❌ لیست خالی است', false);
      printArea.innerHTML = '';
      const total = staff.length;
      printAllBtn.disabled = true;
      printAllBtn.textContent = '⏳ آماده‌سازی چاپ...';
      progressContainer.style.display = 'block';
      progressBar.style.width = '0%';
      progressBar.textContent = '0%';
      let idx = 0;
      function process() {
        if (idx >= total) {
          progressContainer.style.display = 'none';
          printAllBtn.disabled = false;
          printAllBtn.textContent = '🖨️ چاپ همه گواهی‌ها';
          window.print();
          return;
        }
        const tc = createCertificateCanvas(staff[idx]);
        const page = document.createElement('div');
        page.className = 'print-page';
        page.appendChild(tc);
        printArea.appendChild(page);
        idx++;
        progressBar.style.width = `${Math.round((idx / total) * 100)}%`;
        progressBar.textContent = `${Math.round((idx / total) * 100)}%`;
        setTimeout(process, 0);
      }
      process();
    };

    // ==================== INIT ====================
    renderTable();
    loadAndRegisterFonts();
    
    // نمایش پیغام خوش‌آمد با SDK
    if (eitaa) {
      const user = eitaa.initDataUnsafe?.user;
      if (user) {
        console.log(`👤 کاربر: ${user.first_name} ${user.last_name || ''}`);
      }
    }
    
    console.log('🚀 برنامک پنل صدور گواهی آماده است!');
  })();
</script>
</body>
</html>
