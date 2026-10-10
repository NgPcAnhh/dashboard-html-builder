/**
 * EXCEL SMART COMPARISON & REMARKS ENGINE
 * - Focuses strictly on FORMAT & SCHEMA inspection (NO data row-by-row diff between periods)
 * - Evaluates 14 comprehensive format checks with user checklist
 * - Generates ready-to-copy concise prompt feedback for data providers
 * - Renders dual side-by-side Excel viewers for direct visual inspection
 * - Exports clean standardized Excel
 */

(function () {
  'use strict';

  // =========================================================================
  // 1. BUSINESS ALIAS & SYNONYM DICTIONARY (VIETNAMESE & BUSINESS CODES)
  // =========================================================================
  const SYNONYM_GROUPS = [
    {
      canonical: 'ma_nv',
      aliases: ['ma nv', 'manv', 'ma_nv', 'mã nv', 'mã nhân viên', 'ma nhan vien', 'ma_nhan_vien', 'emp_id', 'employee_id', 'staff_code', 'id nhân viên', 'id_nhan_vien', 'ma can bo', 'mã cán bộ']
    },
    {
      canonical: 'ho_ten',
      aliases: ['họ và tên', 'họ tên', 'ho va ten', 'ho ten', 'ho_ten', 'full_name', 'fullname', 'tên nhân viên', 'ten nhan vien', 'ten_nhan_vien', 'name', 'tên', 'ten', 'họ & tên']
    },
    {
      canonical: 'so_dien_thoai',
      aliases: ['sđt', 'sdt', 'số điện thoại', 'so dien thoai', 'so_dien_thoai', 'phone', 'telephone', 'mobile', 'điện thoại', 'dien thoai', 'cellphone', 'tel', 'sdt liên hệ']
    },
    {
      canonical: 'email',
      aliases: ['email', 'e-mail', 'thư điện tử', 'thu dien tu', 'dia chi email', 'dia_chi_email', 'mail', 'hộp thư']
    },
    {
      canonical: 'ngay_sinh',
      aliases: ['ngày sinh', 'ngay sinh', 'ngay_sinh', 'ngaysinh', 'dob', 'date of birth', 'birthdate', 'nam sinh', 'năm sinh']
    },
    {
      canonical: 'dia_chi',
      aliases: ['địa chỉ', 'dia chi', 'dia_chi', 'address', 'nơi ở', 'noi o', 'hộ khẩu', 'ho khau', 'thuong tru', 'thường trú', 'địa chỉ thường trú']
    },
    {
      canonical: 'phong_ban',
      aliases: ['phòng ban', 'phong ban', 'phong_ban', 'bộ phận', 'bo phan', 'bo_phan', 'department', 'dept', 'khối', 'khoi', 'đơn vị', 'don vi']
    },
    {
      canonical: 'chuc_vu',
      aliases: ['chức vụ', 'chuc vu', 'chuc_vu', 'vị trí', 'vi tri', 'vi_tri', 'title', 'position', 'role', 'vai trò', 'vai tro']
    },
    {
      canonical: 'doanh_so',
      aliases: ['doanh số', 'doanh so', 'doanh_so', 'doanh thu', 'doanh_thu', 'revenue', 'sales', 'tổng doanh số', 'tong doanh so', 'tong doanh thu', 'tổng doanh thu', 'tong_tien']
    },
    {
      canonical: 'luong',
      aliases: ['lương', 'luong', 'luong_cb', 'lương cơ bản', 'luong co ban', 'salary', 'thu nhập', 'thu nhap', 'income', 'wage', 'tong luong', 'tổng lương']
    },
    {
      canonical: 'ngay_vao_lam',
      aliases: ['ngày vào làm', 'ngay vao lam', 'ngay_vao', 'ngay_vao_lam', 'join date', 'hire date', 'start date', 'ngay hop dong']
    },
    {
      canonical: 'trang_thai',
      aliases: ['trạng thái', 'trang thai', 'trang_thai', 'tình trạng', 'tinh trang', 'status', 'state', 'active']
    },
    {
      canonical: 'ghi_chu',
      aliases: ['ghi chú', 'ghi chu', 'ghi_chu', 'note', 'notes', 'comment', 'diễn giải', 'dien giai', 'remarks']
    },
    {
      canonical: 'cccd',
      aliases: ['cccd', 'cmnd', 'căn cước', 'can cuoc', 'chứng minh nhân dân', 'chung minh', 'id card', 'so cccd', 'số cccd']
    },
    {
      canonical: 'gioi_tinh',
      aliases: ['giới tính', 'gioi tinh', 'gioi_tinh', 'gender', 'sex', 'nam/nữ', 'nam nu']
    }
  ];

  // Helper: Normalize string (remove Vietnamese accents, symbols, spaces)
  function normalizeText(str) {
    if (!str) return '';
    return str
      .toString()
      .trim()
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[đĐ]/g, 'd')
      .replace(/[^a-z0-9]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function levenshteinDistance(a, b) {
    const matrix = [];
    for (let i = 0; i <= b.length; i++) matrix[i] = [i];
    for (let j = 0; j <= a.length; j++) matrix[0][j] = j;

    for (let i = 1; i <= b.length; i++) {
      for (let j = 1; j <= a.length; j++) {
        if (b.charAt(i - 1) === a.charAt(j - 1)) {
          matrix[i][j] = matrix[i - 1][j - 1];
        } else {
          matrix[i][j] = Math.min(
            matrix[i - 1][j - 1] + 1,
            Math.min(matrix[i][j - 1] + 1, matrix[i - 1][j] + 1)
          );
        }
      }
    }
    return matrix[b.length][a.length];
  }

  function stringSimilarity(s1, s2) {
    const norm1 = normalizeText(s1);
    const norm2 = normalizeText(s2);
    if (norm1 === norm2) return 1.0;
    if (!norm1 || !norm2) return 0.0;

    const maxLen = Math.max(norm1.length, norm2.length);
    const dist = levenshteinDistance(norm1, norm2);
    const charSim = 1.0 - dist / maxLen;

    const tokens1 = new Set(norm1.split(' '));
    const tokens2 = new Set(norm2.split(' '));
    let intersection = 0;
    tokens1.forEach(t => { if (tokens2.has(t)) intersection++; });
    const tokenSim = (2 * intersection) / (tokens1.size + tokens2.size);

    return Math.max(charSim, tokenSim * 0.9);
  }

  function getSynonymCanonical(header) {
    const norm = normalizeText(header);
    for (const group of SYNONYM_GROUPS) {
      for (const alias of group.aliases) {
        if (norm === normalizeText(alias)) {
          return group.canonical;
        }
      }
    }
    return null;
  }

  function isSynonymMatch(header1, header2) {
    const c1 = getSynonymCanonical(header1);
    const c2 = getSynonymCanonical(header2);
    return c1 && c2 && c1 === c2;
  }

  // =========================================================================
  // 2. APPLICATION STATE
  // =========================================================================
  const state = {
    templateFile: {
      name: '',
      workbook: null,
      selectedSheet: '',
      headerRow: 1,
      headers: [],
      rawHeadersWithSpaces: [],
      rows: [],
      rawSheetData: [],
      detectedHeaderRow: 1
    },
    rawFile: {
      name: '',
      workbook: null,
      selectedSheet: '',
      headerRow: 1,
      headers: [],
      rawHeadersWithSpaces: [],
      rows: [],
      rawSheetData: [],
      detectedHeaderRow: 1
    },
    matchedColumns: [],
    unmatchedTemplateHeaders: [],
    extraRawHeaders: [],
    promptMode: 'system', // 'system' | 'human'
    lastActionPoints: [],
    lastSeverity: 'pass'
  };

  // =========================================================================
  // 3. DOM ELEMENTS CACHE
  // =========================================================================
  const el = {
    // Template
    tplDropzone: document.getElementById('tplDropzone'),
    tplFileInput: document.getElementById('tplFileInput'),
    tplMetaBox: document.getElementById('tplMetaBox'),
    tplFileName: document.getElementById('tplFileName'),
    tplBtnRemove: document.getElementById('tplBtnRemove'),
    tplSheetSelect: document.getElementById('tplSheetSelect'),
    tplHeaderRowSelect: document.getElementById('tplHeaderRowSelect'),
    tplStats: document.getElementById('tplStats'),

    // Raw
    rawDropzone: document.getElementById('rawDropzone'),
    rawFileInput: document.getElementById('rawFileInput'),
    rawMetaBox: document.getElementById('rawMetaBox'),
    rawFileName: document.getElementById('rawFileName'),
    rawBtnRemove: document.getElementById('rawBtnRemove'),
    rawSheetSelect: document.getElementById('rawSheetSelect'),
    rawHeaderRowSelect: document.getElementById('rawHeaderRowSelect'),
    rawStats: document.getElementById('rawStats'),

    // Checklist presets
    btnSelectAllChecks: document.getElementById('btnSelectAllChecks'),
    btnDeselectAllChecks: document.getElementById('btnDeselectAllChecks'),
    btnDefaultChecks: document.getElementById('btnDefaultChecks'),
    checksCounter: document.getElementById('checksCounter'),
    btnRunComparison: document.getElementById('btnRunComparison'),

    // Checkboxes (14 rules)
    chkHeaderOffset: document.getElementById('chkHeaderOffset'),
    chkHeaderMatch: document.getElementById('chkHeaderMatch'),
    chkColumnOrder: document.getElementById('chkColumnOrder'),
    chkEmptyCols: document.getElementById('chkEmptyCols'),
    chkSheetName: document.getElementById('chkSheetName'),

    chkFuzzyNames: document.getElementById('chkFuzzyNames'),
    chkHeaderWhitespace: document.getElementById('chkHeaderWhitespace'),
    chkHeaderLineBreaks: document.getElementById('chkHeaderLineBreaks'),

    chkDataTypes: document.getElementById('chkDataTypes'),
    chkFormulaErrors: document.getElementById('chkFormulaErrors'),
    chkNullRate: document.getElementById('chkNullRate'),
    chkDuplicateKeys: document.getElementById('chkDuplicateKeys'),
    chkInterleavedBlankRows: document.getElementById('chkInterleavedBlankRows'),
    chkDateFormats: document.getElementById('chkDateFormats'),

    // Results container
    resultsSection: document.getElementById('resultsSection'),
    btnExportCleanExcel: document.getElementById('btnExportCleanExcel'),
    summaryBanner: document.getElementById('summaryBanner'),
    summaryBannerIcon: document.getElementById('summaryBannerIcon'),
    summaryBannerTitle: document.getElementById('summaryBannerTitle'),
    summaryBannerDesc: document.getElementById('summaryBannerDesc'),
    remarksList: document.getElementById('remarksList'),

    // Prompt Box
    btnCopyPrompt: document.getElementById('btnCopyPrompt'),
    btnModeSystemPrompt: document.getElementById('btnModeSystemPrompt'),
    btnModeHumanMsg: document.getElementById('btnModeHumanMsg'),
    promptTextBox: document.getElementById('promptTextBox'),
    promptCharCount: document.getElementById('promptCharCount'),
    copyToast: document.getElementById('copyToast'),
    copyToastMsg: document.getElementById('copyToastMsg'),

    // Dual Excel Viewers
    tplViewBadge: document.getElementById('tplViewBadge'),
    rawViewBadge: document.getElementById('rawViewBadge'),
    tplTableScroll: document.getElementById('tplTableScroll'),
    rawTableScroll: document.getElementById('rawTableScroll')
  };

  const allCheckboxes = [
    el.chkHeaderOffset,
    el.chkHeaderMatch,
    el.chkColumnOrder,
    el.chkEmptyCols,
    el.chkSheetName,
    el.chkFuzzyNames,
    el.chkHeaderWhitespace,
    el.chkHeaderLineBreaks,
    el.chkDataTypes,
    el.chkFormulaErrors,
    el.chkNullRate,
    el.chkDuplicateKeys,
    el.chkInterleavedBlankRows,
    el.chkDateFormats
  ];

  // =========================================================================
  // 4. HEADER ROW AUTO-DETECTION & PARSING
  // =========================================================================
  function detectHeaderRow(rawSheetData) {
    if (!rawSheetData || rawSheetData.length === 0) return 1;

    let bestRow = 1;
    let maxScore = -1;
    const searchLimit = Math.min(15, rawSheetData.length);

    for (let r = 0; r < searchLimit; r++) {
      const row = rawSheetData[r];
      if (!Array.isArray(row)) continue;

      let filledCount = 0;
      let stringCount = 0;
      let distinctSet = new Set();

      for (let c = 0; c < row.length; c++) {
        const val = row[c];
        if (val !== undefined && val !== null && val.toString().trim() !== '') {
          filledCount++;
          if (typeof val === 'string' && isNaN(Number(val))) {
            stringCount++;
          }
          distinctSet.add(val.toString().trim().toLowerCase());
        }
      }

      const score = filledCount * 2 + stringCount * 3 + distinctSet.size * 2;
      if (score > maxScore && distinctSet.size >= 2) {
        maxScore = score;
        bestRow = r + 1; // 1-indexed
      }
    }
    return bestRow;
  }

  function parseSheetData(sheet, headerRowIndex) {
    const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    const rIdx = Math.max(0, headerRowIndex - 1);

    if (rIdx >= rawData.length) {
      return { headers: [], rawHeadersWithSpaces: [], rows: [], rawSheetData: rawData };
    }

    const rawHeaderRow = rawData[rIdx] || [];
    const headers = [];
    const rawHeadersWithSpaces = [];
    const colCount = rawHeaderRow.length;

    for (let c = 0; c < colCount; c++) {
      const h = rawHeaderRow[c];
      const rawText = h !== undefined && h !== null ? h.toString() : '';
      rawHeadersWithSpaces.push(rawText);

      const name = rawText.trim() !== '' ? rawText.trim() : `Cột_${c + 1}`;
      headers.push(name);
    }

    const dataRows = [];
    for (let r = rIdx + 1; r < rawData.length; r++) {
      const row = rawData[r];
      dataRows.push(row || []);
    }

    return { headers, rawHeadersWithSpaces, rows: dataRows, rawSheetData: rawData };
  }

  // =========================================================================
  // 5. FILE UPLOAD & EVENTS
  // =========================================================================
  function setupFileUploads() {
    el.tplDropzone.addEventListener('click', () => el.tplFileInput.click());
    el.tplFileInput.addEventListener('change', (e) => handleFileSelect(e, 'template'));
    setupDropzoneDnd(el.tplDropzone, 'template');
    el.tplBtnRemove.addEventListener('click', () => clearFile('template'));

    el.rawDropzone.addEventListener('click', () => el.rawFileInput.click());
    el.rawFileInput.addEventListener('change', (e) => handleFileSelect(e, 'raw'));
    setupDropzoneDnd(el.rawDropzone, 'raw');
    el.rawBtnRemove.addEventListener('click', () => clearFile('raw'));

    el.tplSheetSelect.addEventListener('change', () => onSheetChange('template'));
    el.tplHeaderRowSelect.addEventListener('change', () => onHeaderRowChange('template'));

    el.rawSheetSelect.addEventListener('change', () => onSheetChange('raw'));
    el.rawHeaderRowSelect.addEventListener('change', () => onHeaderRowChange('raw'));

    el.btnRunComparison.addEventListener('click', runComparisonWorkflow);
    el.btnExportCleanExcel.addEventListener('click', exportCleanExcel);
    el.btnCopyPrompt.addEventListener('click', copyPromptToClipboard);

    if (el.btnModeSystemPrompt) {
      el.btnModeSystemPrompt.addEventListener('click', () => switchPromptMode('system'));
    }
    if (el.btnModeHumanMsg) {
      el.btnModeHumanMsg.addEventListener('click', () => switchPromptMode('human'));
    }
    if (el.promptTextBox) {
      el.promptTextBox.addEventListener('input', () => {
        if (el.promptCharCount) {
          el.promptCharCount.textContent = `${el.promptTextBox.value.length} ký tự`;
        }
      });
    }
  }

  function setupDropzoneDnd(dropzone, type) {
    ['dragenter', 'dragover'].forEach(eventName => {
      dropzone.addEventListener(eventName, (e) => {
        e.preventDefault();
        dropzone.classList.add('dragover');
      });
    });
    ['dragleave', 'drop'].forEach(eventName => {
      dropzone.addEventListener(eventName, (e) => {
        e.preventDefault();
        dropzone.classList.remove('dragover');
      });
    });
    dropzone.addEventListener('drop', (e) => {
      if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
        processUploadedFile(e.dataTransfer.files[0], type);
      }
    });
  }

  function handleFileSelect(e, type) {
    if (e.target.files && e.target.files.length > 0) {
      processUploadedFile(e.target.files[0], type);
    }
  }

  function processUploadedFile(file, type) {
    const reader = new FileReader();
    reader.onload = function (e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        loadWorkbookIntoState(file.name, workbook, type);
      } catch (err) {
        alert('Lỗi đọc file Excel: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function loadWorkbookIntoState(fileName, workbook, type) {
    const fileState = type === 'template' ? state.templateFile : state.rawFile;
    fileState.name = fileName;
    fileState.workbook = workbook;
    fileState.selectedSheet = workbook.SheetNames[0] || '';

    const firstSheet = workbook.Sheets[fileState.selectedSheet];
    const rawData = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });
    fileState.detectedHeaderRow = detectHeaderRow(rawData);
    fileState.headerRow = fileState.detectedHeaderRow;

    const parsed = parseSheetData(firstSheet, fileState.headerRow);
    fileState.headers = parsed.headers;
    fileState.rawHeadersWithSpaces = parsed.rawHeadersWithSpaces;
    fileState.rows = parsed.rows;
    fileState.rawSheetData = parsed.rawSheetData;

    updateFileUI(type);
    checkEnableRun();
  }

  function updateFileUI(type) {
    const isTpl = type === 'template';
    const fileState = isTpl ? state.templateFile : state.rawFile;
    const dropzone = isTpl ? el.tplDropzone : el.rawDropzone;
    const metaBox = isTpl ? el.tplMetaBox : el.rawMetaBox;
    const fileNameEl = isTpl ? el.tplFileName : el.rawFileName;
    const sheetSelect = isTpl ? el.tplSheetSelect : el.rawSheetSelect;
    const headerRowSelect = isTpl ? el.tplHeaderRowSelect : el.rawHeaderRowSelect;
    const statsEl = isTpl ? el.tplStats : el.rawStats;

    dropzone.style.display = 'none';
    metaBox.style.display = 'flex';
    fileNameEl.textContent = fileState.name;

    sheetSelect.innerHTML = '';
    fileState.workbook.SheetNames.forEach(sheetName => {
      const opt = document.createElement('option');
      opt.value = sheetName;
      opt.textContent = sheetName;
      sheetSelect.appendChild(opt);
    });
    sheetSelect.value = fileState.selectedSheet;

    headerRowSelect.innerHTML = '';
    for (let r = 1; r <= 10; r++) {
      const opt = document.createElement('option');
      opt.value = r;
      opt.textContent = `Dòng ${r} ${r === fileState.detectedHeaderRow ? '(Phát hiện ⭐)' : ''}`;
      headerRowSelect.appendChild(opt);
    }
    headerRowSelect.value = fileState.headerRow;

    const dataRowCount = fileState.rows.filter(row => row && row.some(c => c !== '' && c !== null && c !== undefined)).length;
    statsEl.innerHTML = `
      <span class="stat-pill">📊 ${fileState.headers.length} Cột</span>
      <span class="stat-pill">📄 ${dataRowCount} Dòng dữ liệu</span>
    `;
  }

  function clearFile(type) {
    const isTpl = type === 'template';
    const fileState = isTpl ? state.templateFile : state.rawFile;
    const dropzone = isTpl ? el.tplDropzone : el.rawDropzone;
    const metaBox = isTpl ? el.tplMetaBox : el.rawMetaBox;
    const input = isTpl ? el.tplFileInput : el.rawFileInput;

    fileState.name = '';
    fileState.workbook = null;
    fileState.headers = [];
    fileState.rows = [];
    input.value = '';

    dropzone.style.display = 'flex';
    metaBox.style.display = 'none';

    el.resultsSection.classList.remove('visible');
    checkEnableRun();
  }

  function onSheetChange(type) {
    const isTpl = type === 'template';
    const fileState = isTpl ? state.templateFile : state.rawFile;
    const sheetSelect = isTpl ? el.tplSheetSelect : el.rawSheetSelect;

    fileState.selectedSheet = sheetSelect.value;
    const sheet = fileState.workbook.Sheets[fileState.selectedSheet];
    const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    fileState.detectedHeaderRow = detectHeaderRow(rawData);
    fileState.headerRow = fileState.detectedHeaderRow;

    const parsed = parseSheetData(sheet, fileState.headerRow);
    fileState.headers = parsed.headers;
    fileState.rawHeadersWithSpaces = parsed.rawHeadersWithSpaces;
    fileState.rows = parsed.rows;
    fileState.rawSheetData = parsed.rawSheetData;

    updateFileUI(type);
  }

  function onHeaderRowChange(type) {
    const isTpl = type === 'template';
    const fileState = isTpl ? state.templateFile : state.rawFile;
    const headerRowSelect = isTpl ? el.tplHeaderRowSelect : el.rawHeaderRowSelect;

    fileState.headerRow = parseInt(headerRowSelect.value, 10) || 1;
    const sheet = fileState.workbook.Sheets[fileState.selectedSheet];
    const parsed = parseSheetData(sheet, fileState.headerRow);
    fileState.headers = parsed.headers;
    fileState.rawHeadersWithSpaces = parsed.rawHeadersWithSpaces;
    fileState.rows = parsed.rows;
    fileState.rawSheetData = parsed.rawSheetData;

    updateFileUI(type);
  }

  function checkEnableRun() {
    const ready = state.templateFile.workbook && state.rawFile.workbook;
    el.btnRunComparison.disabled = !ready;
  }

  function setupChecklistHandlers() {
    allCheckboxes.forEach(chk => {
      chk.addEventListener('change', () => {
        const card = chk.closest('.check-item-card');
        if (card) card.classList.toggle('checked', chk.checked);
        updateChecksCounter();
      });
    });

    el.btnSelectAllChecks.addEventListener('click', () => {
      allCheckboxes.forEach(chk => {
        chk.checked = true;
        const card = chk.closest('.check-item-card');
        if (card) card.classList.add('checked');
      });
      updateChecksCounter();
    });

    el.btnDeselectAllChecks.addEventListener('click', () => {
      allCheckboxes.forEach(chk => {
        chk.checked = false;
        const card = chk.closest('.check-item-card');
        if (card) card.classList.remove('checked');
      });
      updateChecksCounter();
    });

    el.btnDefaultChecks.addEventListener('click', () => {
      allCheckboxes.forEach(chk => {
        chk.checked = true;
        const card = chk.closest('.check-item-card');
        if (card) card.classList.add('checked');
      });
      updateChecksCounter();
    });
  }

  function updateChecksCounter() {
    const selectedCount = allCheckboxes.filter(c => c.checked).length;
    el.checksCounter.textContent = `Đã chọn ${selectedCount} / ${allCheckboxes.length} phép kiểm tra`;
  }

  // =========================================================================
  // 6. COLUMN MATCHING ALGORITHM (HEADER LEVEL ONLY)
  // =========================================================================
  function buildColumnMapping() {
    const tplHeaders = state.templateFile.headers;
    const rawHeaders = state.rawFile.headers;
    const matched = [];
    const usedRaw = new Set();

    // 1. Exact match
    tplHeaders.forEach((th, tIdx) => {
      let found = null;
      rawHeaders.forEach((rh, rIdx) => {
        if (th.trim().toLowerCase() === rh.trim().toLowerCase()) {
          found = { tIdx, targetHeader: th, rIdx, sourceHeader: rh, type: 'EXACT', confidence: 100 };
        } else if (!found && normalizeText(th) === normalizeText(rh)) {
          found = { tIdx, targetHeader: th, rIdx, sourceHeader: rh, type: 'NORMALIZED', confidence: 98 };
        }
      });
      if (found) {
        usedRaw.add(found.rIdx);
        matched.push(found);
      }
    });

    // 2. Synonyms / Aliases
    tplHeaders.forEach((th, tIdx) => {
      if (matched.some(m => m.tIdx === tIdx)) return;
      rawHeaders.forEach((rh, rIdx) => {
        if (usedRaw.has(rIdx)) return;
        if (isSynonymMatch(th, rh)) {
          usedRaw.add(rIdx);
          matched.push({ tIdx, targetHeader: th, rIdx, sourceHeader: rh, type: 'SYNONYM', confidence: 95 });
        }
      });
    });

    // 3. Fuzzy
    tplHeaders.forEach((th, tIdx) => {
      if (matched.some(m => m.tIdx === tIdx)) return;
      let bestSim = 0.65;
      let bestRIdx = -1;
      rawHeaders.forEach((rh, rIdx) => {
        if (usedRaw.has(rIdx)) return;
        const sim = stringSimilarity(th, rh);
        if (sim > bestSim) {
          bestSim = sim;
          bestRIdx = rIdx;
        }
      });
      if (bestRIdx !== -1) {
        usedRaw.add(bestRIdx);
        matched.push({ tIdx, targetHeader: th, rIdx: bestRIdx, sourceHeader: rawHeaders[bestRIdx], type: 'FUZZY', confidence: Math.round(bestSim * 100) });
      }
    });

    state.matchedColumns = matched;
    state.unmatchedTemplateHeaders = tplHeaders
      .map((h, i) => ({ header: h, index: i }))
      .filter(item => !matched.some(m => m.tIdx === item.index));

    state.extraRawHeaders = rawHeaders
      .map((h, i) => ({ header: h, index: i }))
      .filter(item => !usedRaw.has(item.index));
  }

  // =========================================================================
  // 7. COMPARISON WORKFLOW (14 FORMAT RULES)
  // =========================================================================
  function runComparisonWorkflow() {
    const tpl = state.templateFile;
    const raw = state.rawFile;

    if (!tpl.workbook || !raw.workbook) {
      alert('Vui lòng nạp đủ cả 2 file để thực hiện so sánh!');
      return;
    }

    buildColumnMapping();

    const remarks = [];
    const actionPointsForPrompt = [];
    let worstSeverity = 'pass';

    function addRemark(rem, actionShort) {
      remarks.push(rem);
      if (rem.severity === 'fail') worstSeverity = 'fail';
      else if (rem.severity === 'warning' && worstSeverity !== 'fail') worstSeverity = 'warning';

      if (actionShort) {
        actionPointsForPrompt.push(actionShort);
      }
    }

    // Rule 1: Header Row Offset
    if (el.chkHeaderOffset.checked) {
      const r = checkHeaderOffset(tpl, raw);
      addRemark(r.remark, r.actionShort);
    }

    // Rule 2: Header Match (Missing/Extra)
    if (el.chkHeaderMatch.checked) {
      const r = checkHeaderMatch(tpl, raw);
      addRemark(r.remark, r.actionShort);
    }

    // Rule 3: Column Order (Shuffled)
    if (el.chkColumnOrder.checked) {
      const r = checkColumnOrder(tpl, raw);
      addRemark(r.remark, r.actionShort);
    }

    // Rule 4: Empty Unnamed Columns
    if (el.chkEmptyCols.checked) {
      const r = checkEmptyCols(tpl, raw);
      addRemark(r.remark, r.actionShort);
    }

    // Rule 5: Sheet Name Mismatch
    if (el.chkSheetName.checked) {
      const r = checkSheetName(tpl, raw);
      addRemark(r.remark, r.actionShort);
    }

    // Rule 6: Fuzzy Names & Synonyms
    if (el.chkFuzzyNames.checked) {
      const r = checkFuzzyNames(tpl, raw);
      addRemark(r.remark, r.actionShort);
    }

    // Rule 7: Header Whitespace
    if (el.chkHeaderWhitespace.checked) {
      const r = checkHeaderWhitespace(tpl, raw);
      addRemark(r.remark, r.actionShort);
    }

    // Rule 8: Line Breaks in Header
    if (el.chkHeaderLineBreaks.checked) {
      const r = checkHeaderLineBreaks(tpl, raw);
      addRemark(r.remark, r.actionShort);
    }

    // Rule 9: Number Stored as Text
    if (el.chkDataTypes.checked) {
      const r = checkDataTypes(tpl, raw);
      addRemark(r.remark, r.actionShort);
    }

    // Rule 10: Formula Errors
    if (el.chkFormulaErrors.checked) {
      const r = checkFormulaErrors(tpl, raw);
      addRemark(r.remark, r.actionShort);
    }

    // Rule 11: Null Rate
    if (el.chkNullRate.checked) {
      const r = checkNullRate(tpl, raw);
      addRemark(r.remark, r.actionShort);
    }

    // Rule 12: Duplicate Keys in incoming file
    if (el.chkDuplicateKeys.checked) {
      const r = checkDuplicateKeys(tpl, raw);
      addRemark(r.remark, r.actionShort);
    }

    // Rule 13: Interleaved Blank Rows
    if (el.chkInterleavedBlankRows.checked) {
      const r = checkInterleavedBlankRows(tpl, raw);
      addRemark(r.remark, r.actionShort);
    }

    // Rule 14: Date Formats
    if (el.chkDateFormats.checked) {
      const r = checkDateFormats(tpl, raw);
      addRemark(r.remark, r.actionShort);
    }

    // Render Prompt Box
    renderPromptFeedback(raw.name, actionPointsForPrompt, worstSeverity);

    // Render Remarks & Summary
    renderExecutiveSummary(worstSeverity);
    renderRemarksList(remarks);

    // Render Dual Excel Viewers (no cell-diffing, pure spreadsheet viewing)
    renderDualExcelViewers(tpl, raw);

    el.resultsSection.classList.add('visible');
    el.resultsSection.scrollIntoView({ behavior: 'smooth' });
  }

  // =========================================================================
  // 8. 14 DETAILED EVALUATION RULES (FORMAT & SCHEMA ONLY)
  // =========================================================================

  // Rule 1: Header Row Offset
  function checkHeaderOffset(tpl, raw) {
    if (raw.headerRow > 1) {
      return {
        remark: {
          title: '📍 Dòng Tiêu Đề Bị Lệch / Chèn Banner Phía Trên',
          severity: 'warning',
          badge: `Tiêu đề nằm tại Dòng ${raw.headerRow}`,
          content: `Dòng tiêu đề của file mới bắt đầu tại <b>Dòng ${raw.headerRow}</b> thay vì Dòng 1 như quy chuẩn. Có ${raw.headerRow - 1} dòng trống hoặc chứa banner tiêu đề báo cáo ở đầu file.`,
          action: 'Yêu cầu: Xóa bỏ các dòng banner/dòng trống đầu trang để dòng tiêu đề nằm chính xác tại Dòng 1.'
        },
        actionShort: `Đặt lại dòng tiêu đề tại Dòng 1 (hiện đang nằm ở Dòng ${raw.headerRow} do có banner/dòng trống phía trên).`
      };
    }
    return {
      remark: {
        title: '📍 Vị Trí Dòng Tiêu Đề',
        severity: 'pass',
        badge: 'Đúng Dòng 1',
        content: 'Dòng tiêu đề của cả 2 file đều nằm chuẩn xác tại Dòng 1, không có dòng rác hoặc banner phía trên.',
        action: null
      },
      actionShort: null
    };
  }

  // Rule 2: Header Match (Missing/Extra)
  function checkHeaderMatch(tpl, raw) {
    const missing = state.unmatchedTemplateHeaders;
    const extra = state.extraRawHeaders;

    if (missing.length > 0) {
      const missingNames = missing.map(m => `"${escapeHtml(m.header)}"`).join(', ');
      let extraText = '';
      if (extra.length > 0) {
        extraText = `<br>• File mới đồng thời gửi dư thừa ${extra.length} cột: [${extra.map(e => `"${escapeHtml(e.header)}"`).join(', ')}].`;
      }
      return {
        remark: {
          title: '📋 Đồng Nhất Danh Sách Cột (Thiếu / Thừa cột)',
          severity: 'fail',
          badge: `Thiếu ${missing.length} Cột Bắt Buộc`,
          content: `File gửi mới <b>BỊ THIẾU ${missing.length} CỘT BẮT BUỘC</b> của file mẫu chuẩn:<br>• Danh sách cột thiếu: ${missingNames}.${extraText}`,
          action: 'Yêu cầu: Bổ sung đầy đủ các cột bị thiếu trên vào file báo cáo trước khi gửi.'
        },
        actionShort: `Bổ sung ${missing.length} cột bị thiếu: ${missingNames}${extra.length > 0 ? ` (và xóa ${extra.length} cột thừa: ${extra.map(e => `"${e.header}"`).join(', ')})` : ''}.`
      };
    }

    if (extra.length > 0) {
      const extraNames = extra.map(e => `"${escapeHtml(e.header)}"`).join(', ');
      return {
        remark: {
          title: '📋 Danh Sách Cột',
          severity: 'warning',
          badge: `Thừa ${extra.length} Cột Ngoài Mẫu`,
          content: `File mới có đủ các cột cần thiết nhưng bị dư thừa ${extra.length} cột ngoài mẫu quy chuẩn: ${extraNames}.`,
          action: 'Yêu cầu: Xóa bỏ các cột thừa ngoài mẫu để tránh lỗi lệch cấu trúc khi import.'
        },
        actionShort: `Xóa bỏ ${extra.length} cột dư thừa không nằm trong mẫu: ${extraNames}.`
      };
    }

    return {
      remark: {
        title: '📋 Đồng Nhất Danh Sách Cột',
        severity: 'pass',
        badge: 'Đủ 100% Cột',
        content: `File mới có đầy đủ ${tpl.headers.length} cột theo đúng danh mục của mẫu chuẩn, không thiếu hoặc thừa cột nào.`,
        action: null
      },
      actionShort: null
    };
  }

  // Rule 3: Column Order (Shuffled)
  function checkColumnOrder(tpl, raw) {
    const shuffled = [];
    state.matchedColumns.forEach(m => {
      if (m.tIdx !== m.rIdx) {
        shuffled.push({
          target: m.targetHeader,
          targetIdx: m.tIdx + 1,
          rawIdx: m.rIdx + 1,
          source: m.sourceHeader
        });
      }
    });

    if (shuffled.length > 0) {
      const list = shuffled.map(s => `<li>Cột <b>"${escapeHtml(s.target)}"</b>: Vị trí mẫu là <b>Cột #${s.targetIdx}</b> ↔ Nhưng file mới đặt ở <b>Cột #${s.rawIdx}</b></li>`).join('');
      const swappedSummary = shuffled.map(s => `"${s.target}" (từ Cột #${s.rawIdx} về #${s.targetIdx})`).join(', ');
      return {
        remark: {
          title: '🔄 Đảo Thứ Tự Cột (Trùng cột nhưng sai vị trí)',
          severity: 'warning',
          badge: `${shuffled.length} Cột Bị Đặt Sai Vị Trí`,
          content: `Các cột sau đây có trong mẫu nhưng <b>ĐẶT SAI THỨ TỰ CỘT</b> so với mẫu chuẩn:<br><ul>${list}</ul>`,
          action: 'Yêu cầu: Sắp xếp lại thứ tự các cột từ trái qua phải theo đúng thứ tự file mẫu.'
        },
        actionShort: `Sắp xếp lại các cột bị đảo vị trí: ${swappedSummary}.`
      };
    }

    return {
      remark: {
        title: '🔄 Kiểm Tra Thứ Tự Cột',
        severity: 'pass',
        badge: 'Thứ Tự Khớp 100%',
        content: 'Tất cả các cột trong file gửi mới đều được đặt đúng thứ tự từ trái qua phải tương ứng tuyệt đối với file mẫu.',
        action: null
      },
      actionShort: null
    };
  }

  // Rule 4: Empty Unnamed Columns
  function checkEmptyCols(tpl, raw) {
    const unnamedCols = [];
    raw.headers.forEach((h, idx) => {
      if (/^Cột_\d+$/i.test(h) || /^Column\d+$/i.test(h) || !h.trim()) {
        unnamedCols.push({ idx: idx + 1, name: h });
      }
    });

    if (unnamedCols.length > 0) {
      const colList = unnamedCols.map(c => `Cột #${c.idx}`).join(', ');
      return {
        remark: {
          title: '🗑️ Cột Trống Vô Danh (Blank / Unnamed Columns)',
          severity: 'warning',
          badge: `${unnamedCols.length} Cột Trống`,
          content: `Phát hiện <b>${unnamedCols.length} cột không có tiêu đề</b> trong file mới (${colList}) do xóa dữ liệu nhưng quên xóa cột Excel.`,
          action: 'Yêu cầu: Xóa hẳn (Delete Column) các cột trống này trong Excel trước khi gửi.'
        },
        actionShort: `Xóa hẳn các cột trống không có tên (${colList}) khỏi file Excel.`
      };
    }

    return {
      remark: {
        title: '🗑️ Kiểm Tra Cột Trống Vô Danh',
        severity: 'pass',
        badge: 'Không Có Cột Rác',
        content: 'File mới không có cột trống vô danh nào.',
        action: null
      },
      actionShort: null
    };
  }

  // Rule 5: Sheet Name Mismatch
  function checkSheetName(tpl, raw) {
    if (tpl.selectedSheet && raw.selectedSheet && tpl.selectedSheet !== raw.selectedSheet) {
      return {
        remark: {
          title: '📑 Kiểm Tra Tên Sheet / Trang Tính',
          severity: 'warning',
          badge: 'Tên Sheet Khác Biệt',
          content: `Tên sheet của mẫu chuẩn là <b>"${escapeHtml(tpl.selectedSheet)}"</b>, nhưng file mới gửi đặt tên sheet là <b>"${escapeHtml(raw.selectedSheet)}"</b>.`,
          action: 'Yêu cầu: Đổi lại tên trang tính (Sheet) đúng chuẩn theo file mẫu.'
        },
        actionShort: `Đổi lại tên trang tính (Sheet) thành "${tpl.selectedSheet}" (hiện đang là "${raw.selectedSheet}").`
      };
    }
    return {
      remark: {
        title: '📑 Kiểm Tra Tên Sheet',
        severity: 'pass',
        badge: 'Tên Sheet Đúng Chuẩn',
        content: `Cả 2 file đều sử dụng đúng tên trang tính chuẩn: "${escapeHtml(tpl.selectedSheet)}".`,
        action: null
      },
      actionShort: null
    };
  }

  // Rule 6: Fuzzy Names & Synonyms
  function checkFuzzyNames(tpl, raw) {
    const altered = state.matchedColumns.filter(m => m.type === 'SYNONYM' || m.type === 'FUZZY');

    if (altered.length > 0) {
      const list = altered.map(a => `<li>Cột mẫu: <b>"${escapeHtml(a.targetHeader)}"</b> ↔ File mới ghi: <b>"${escapeHtml(a.sourceHeader)}"</b></li>`).join('');
      const shortList = altered.map(a => `"${a.sourceHeader}" -> "${a.targetHeader}"`).join(', ');
      return {
        remark: {
          title: '🔤 Sai Tên Trường, Viết Tắt, Không Dấu',
          severity: 'warning',
          badge: `${altered.length} Cột Sai Tên / Viết Tắt`,
          content: `File mới viết tắt hoặc sai tên tiêu đề cột so với quy ước chuẩn:<br><ul>${list}</ul>`,
          action: 'Yêu cầu: Đổi lại tên các cột theo đúng chuẩn chữ hoa, dấu tiếng Việt như mẫu.'
        },
        actionShort: `Sửa lại đúng tên các tiêu đề cột theo mẫu chuẩn: ${shortList}.`
      };
    }

    return {
      remark: {
        title: '🔤 Kiểm Tra Tên Tiêu Đề Cột',
        severity: 'pass',
        badge: 'Tên Tiêu Đề Chuẩn Xác',
        content: 'Tên các tiêu đề cột đều được viết chính xác từng chữ, đúng dấu tiếng Việt và chuẩn quy ước.',
        action: null
      },
      actionShort: null
    };
  }

  // Rule 7: Header Whitespace
  function checkHeaderWhitespace(tpl, raw) {
    const whitespaceHeaders = [];
    raw.rawHeadersWithSpaces.forEach((h, idx) => {
      if (typeof h === 'string' && (h.startsWith(' ') || h.endsWith(' ') || h.includes('  '))) {
        whitespaceHeaders.push({ idx: idx + 1, name: h });
      }
    });

    if (whitespaceHeaders.length > 0) {
      const list = whitespaceHeaders.map(w => `Cột #${w.idx} (<code>"${escapeHtml(w.name)}"</code>)`).join(', ');
      return {
        remark: {
          title: '␣ Khoảng Trắng Thừa Ở Đầu/Cuối Tiêu Đề',
          severity: 'warning',
          badge: `${whitespaceHeaders.length} Cột Có Khoảng Trắng Thừa`,
          content: `Phát hiện dấu cách thừa ở đầu hoặc cuối tên cột tại: ${list}. Dấu cách này làm các câu lệnh import tự động và VLOOKUP bị lỗi không nhận dạng được chuỗi.`,
          action: 'Yêu cầu: Xóa các dấu khoảng trắng thừa ở đầu và cuối tên cột.'
        },
        actionShort: `Xóa bỏ khoảng trắng (dấu cách) thừa ở đầu/cuối tiêu đề các cột: ${list}.`
      };
    }

    return {
      remark: {
        title: '␣ Kiểm Tra Khoảng Trắng Tiêu Đề',
        severity: 'pass',
        badge: 'Không Có Khoảng Trắng Thừa',
        content: 'Tất cả tiêu đề cột đều được gọt sạch khoảng trắng (trim) chuẩn xác.',
        action: null
      },
      actionShort: null
    };
  }

  // Rule 8: Line Breaks in Header
  function checkHeaderLineBreaks(tpl, raw) {
    const brokenHeaders = [];
    raw.rawHeadersWithSpaces.forEach((h, idx) => {
      if (typeof h === 'string' && (/[\r\n\t]/.test(h))) {
        brokenHeaders.push({ idx: idx + 1, name: h.replace(/[\r\n]+/g, ' [Enter] ') });
      }
    });

    if (brokenHeaders.length > 0) {
      const list = brokenHeaders.map(b => `Cột #${b.idx} ("${escapeHtml(b.name)}")`).join(', ');
      return {
        remark: {
          title: '↵ Xuống Dòng (Alt+Enter) Trong Tiêu Đề',
          severity: 'warning',
          badge: `${brokenHeaders.length} Cột Bị Xuống Dòng`,
          content: `Có ${brokenHeaders.length} tiêu đề cột chứa ký tự xuống dòng (Alt+Enter) hoặc Tab: ${list}. Ký tự này gây vỡ dòng khi xuất file CSV/SQL.`,
          action: 'Yêu cầu: Xóa ký tự xuống dòng, viết tiêu đề trên 1 dòng duy nhất.'
        },
        actionShort: `Xóa ký tự xuống dòng (Alt+Enter) trong tiêu đề cột ${list}, viết trên 1 dòng duy nhất.`
      };
    }

    return {
      remark: {
        title: '↵ Kiểm Tra Ký Tự Xuống Dòng Tiêu Đề',
        severity: 'pass',
        badge: 'Tiêu Đề Liền Mạch',
        content: 'Không phát hiện ký tự xuống dòng (Alt+Enter) hoặc ký tự điều khiển ẩn trong tiêu đề.',
        action: null
      },
      actionShort: null
    };
  }

  // Rule 9: Number Stored as Text
  function checkDataTypes(tpl, raw) {
    const rRows = raw.rows;
    const textNumberCols = [];

    state.matchedColumns.forEach(m => {
      // Find columns that should be numbers (doanh so, luong, so luong, don gia...)
      const norm = normalizeText(m.targetHeader);
      const isNumExpected = norm.includes('doanh') || norm.includes('luong') || norm.includes('tien') || norm.includes('so luong') || norm.includes('gia') || norm.includes('rate');

      if (isNumExpected && rRows.length > 0) {
        let stringCount = 0;
        let sampleVal = '';
        for (let r = 0; r < Math.min(25, rRows.length); r++) {
          const val = rRows[r][m.rIdx];
          if (val !== undefined && val !== null && val !== '') {
            if (typeof val === 'string' && isNaN(Number(val))) {
              stringCount++;
              if (!sampleVal) sampleVal = val;
            }
          }
        }
        if (stringCount > 0) {
          textNumberCols.push({ col: m.targetHeader, count: stringCount, sample: sampleVal });
        }
      }
    });

    if (textNumberCols.length > 0) {
      const list = textNumberCols.map(t => `<li>Cột <b>"${escapeHtml(t.col)}"</b>: Chứa giá trị dạng chữ (ví dụ: <code>"${escapeHtml(t.sample)}"</code>) không tính toán được.</li>`).join('');
      return {
        remark: {
          title: '🔢 Cột Số Bị Lưu Dạng Text / Chuỗi Ký Tự',
          severity: 'warning',
          badge: `${textNumberCols.length} Cột Số Bị Lỗi Dạng Text`,
          content: `Phát hiện các cột tiền tệ/số lượng ở file mới có lẫn ký tự lạ (dấu phân cách không chuẩn, chữ 'VNĐ', dấu nháy '...):<br><ul>${list}</ul>`,
          action: 'Yêu cầu: Chuyển định dạng ô về dạng Number (Format Cells -> Number), loại bỏ chữ đính kèm.'
        },
        actionShort: `Chuyển định dạng các cột số (${textNumberCols.map(t => `"${t.col}"`).join(', ')}) về kiểu Number thuần túy, bỏ ký tự chữ hoặc dấu nháy.`
      };
    }

    return {
      remark: {
        title: '🔢 Kiểm Tra Định Dạng Cột Số',
        severity: 'pass',
        badge: 'Định Dạng Số Hợp Lệ',
        content: 'Các cột tiền tệ và số liệu ở file mới đều lưu đúng kiểu số, không bị lẫn chuỗi chữ.',
        action: null
      },
      actionShort: null
    };
  }

  // Rule 10: Formula Errors
  function checkFormulaErrors(tpl, raw) {
    const errorCodes = ['#VALUE!', '#REF!', '#N/A', '#DIV/0!', '#NAME?', '#NUM!', '#NULL!'];
    const rRows = raw.rows;
    let foundErrors = 0;
    const errorLocations = [];

    for (let r = 0; r < Math.min(100, rRows.length); r++) {
      const row = rRows[r];
      if (!row) continue;
      for (let c = 0; c < row.length; c++) {
        const val = row[c] ? row[c].toString().trim() : '';
        if (errorCodes.includes(val)) {
          foundErrors++;
          if (errorLocations.length < 5) {
            errorLocations.push(`Dòng ${r + raw.headerRow + 1} Cột ${raw.headers[c] || c + 1} (${val})`);
          }
        }
      }
    }

    if (foundErrors > 0) {
      return {
        remark: {
          title: '⚠️ Tồn Tại Lỗi Công Thức Excel (#REF!, #VALUE!, #N/A)',
          severity: 'fail',
          badge: `${foundErrors} Ô Bị Lỗi Công Thức`,
          content: `File mới còn chứa <b>${foundErrors} ô bị lỗi công thức Excel</b> chưa được xử lý trước khi gửi (ví dụ: ${errorLocations.join(', ')}...).`,
          action: 'Yêu cầu: Kiểm tra lại các hàm Excel bị gãy công thức hoặc Paste as Values trước khi gửi file.'
        },
        actionShort: `Sửa hoặc Paste as Values ${foundErrors} ô bị lỗi công thức Excel (#REF!, #VALUE!, #N/A).`
      };
    }

    return {
      remark: {
        title: '⚠️ Kiểm Tra Lỗi Công Thức Excel',
        severity: 'pass',
        badge: 'Không Có Lỗi Công Thức',
        content: 'Không phát hiện ô nào bị lỗi công thức (#REF!, #VALUE!, #N/A) trong file mới.',
        action: null
      },
      actionShort: null
    };
  }

  // Rule 11: Null Rate
  function checkNullRate(tpl, raw) {
    const highNulls = [];
    const rRows = raw.rows;

    state.matchedColumns.forEach(m => {
      let emptyCount = 0;
      let totalValidRows = 0;
      for (let r = 0; r < rRows.length; r++) {
        if (rRows[r] && rRows[r].some(c => c !== '' && c !== null)) {
          totalValidRows++;
          const val = rRows[r][m.rIdx];
          if (val === undefined || val === null || val.toString().trim() === '') {
            emptyCount++;
          }
        }
      }
      if (totalValidRows > 0) {
        const rate = Math.round((emptyCount / totalValidRows) * 100);
        if (rate >= 40) {
          highNulls.push({ col: m.targetHeader, rate, emptyCount, total: totalValidRows });
        }
      }
    });

    if (highNulls.length > 0) {
      const list = highNulls.map(h => `<li>Cột <b>"${escapeHtml(h.col)}"</b>: Có <b>${h.rate}% ô trống</b> (${h.emptyCount}/${h.total} dòng).</li>`).join('');
      return {
        remark: {
          title: '📉 Ô Trống Bất Thường (Tỷ Lệ Blank / Null Cao)',
          severity: 'warning',
          badge: `${highNulls.length} Cột Thiếu Dữ Liệu`,
          content: `Các cột sau đây ở file mới có tỷ lệ ô trống cao bất thường:<br><ul>${list}</ul>`,
          action: 'Yêu cầu: Rà soát lại xem dữ liệu có bị xuất thiếu hay bỏ sót không.'
        },
        actionShort: `Kiểm tra bổ sung dữ liệu cho các cột bị bỏ trống nhiều (${highNulls.map(h => `"${h.col}"`).join(', ')}).`
      };
    }

    return {
      remark: {
        title: '📉 Kiểm Tra Ô Trống Bất Thường',
        severity: 'pass',
        badge: 'Độ Phủ Tốt',
        content: 'Dữ liệu các cột ở kỳ mới được điền đầy đủ, không có cột nào bị rỗng bất thường.',
        action: null
      },
      actionShort: null
    };
  }

  // Rule 12: Duplicate Keys in incoming file
  function checkDuplicateKeys(tpl, raw) {
    const rRows = raw.rows;
    let keyCol = state.matchedColumns.find(m => {
      const norm = normalizeText(m.targetHeader);
      return norm.includes('ma ') || norm.includes('id') || norm.includes('code');
    });

    if (keyCol && rRows.length > 0) {
      const counts = {};
      const duplicates = new Set();

      for (let r = 0; r < rRows.length; r++) {
        const val = rRows[r][keyCol.rIdx];
        if (val !== undefined && val !== null && val.toString().trim() !== '') {
          const key = val.toString().trim();
          counts[key] = (counts[key] || 0) + 1;
          if (counts[key] > 1) {
            duplicates.add(key);
          }
        }
      }

      if (duplicates.size > 0) {
        const dupList = Array.from(duplicates).slice(0, 4).map(d => `"${escapeHtml(d)}"`).join(', ');
        return {
          remark: {
            title: '🆔 Trùng Lặp Mã Định Danh Trong Kỳ Mới',
            severity: 'warning',
            badge: `Trùng ${duplicates.size} Mã`,
            content: `Phát hiện cột định danh <b>"${escapeHtml(keyCol.targetHeader)}"</b> của file mới có <b>${duplicates.size} mã bị lặp lại nhiều dòng</b> (ví dụ: ${dupList}...).`,
            action: 'Yêu cầu: Kiểm tra và xóa các dòng dữ liệu bị gửi trùng lặp (duplicate rows).'
          },
          actionShort: `Kiểm tra và loại bỏ các dòng bị trùng lặp mã định danh "${keyCol.targetHeader}" (${dupList}...).`
        };
      }
    }

    return {
      remark: {
        title: '🆔 Kiểm Tra Trùng Lặp Mã Định Danh',
        severity: 'pass',
        badge: 'Mã Định Danh Duy Nhất',
        content: 'Không phát hiện trùng lặp mã định danh trong kỳ mới; mỗi bản ghi là duy nhất.',
        action: null
      },
      actionShort: null
    };
  }

  // Rule 13: Interleaved Blank Rows
  function checkInterleavedBlankRows(tpl, raw) {
    const rRows = raw.rows;
    let blankRowCount = 0;
    let hasDataBelow = false;

    // Search for empty rows that have non-empty rows below them
    for (let r = 0; r < rRows.length - 1; r++) {
      const isCurrentBlank = !rRows[r] || rRows[r].every(c => c === '' || c === null || c === undefined);
      if (isCurrentBlank) {
        let foundBelow = false;
        for (let k = r + 1; k < Math.min(r + 5, rRows.length); k++) {
          if (rRows[k] && rRows[k].some(c => c !== '' && c !== null && c !== undefined)) {
            foundBelow = true;
            break;
          }
        }
        if (foundBelow) {
          blankRowCount++;
        }
      }
    }

    if (blankRowCount > 0) {
      return {
        remark: {
          title: '⏸️ Dòng Hoàn Toàn Trống Xen Kẽ Giữa Bảng',
          severity: 'warning',
          badge: `${blankRowCount} Dòng Trống Xen Kẽ`,
          content: `Có <b>${blankRowCount} dòng hoàn toàn trống</b> nằm xen kẽ giữa các dòng dữ liệu của file mới, làm ngắt quãng bảng tính.`,
          action: 'Yêu cầu: Xóa bỏ các dòng trắng xen kẽ để dữ liệu được liên tục.'
        },
        actionShort: `Xóa bỏ ${blankRowCount} dòng trắng hoàn toàn nằm xen kẽ giữa các dòng dữ liệu.`
      };
    }

    return {
      remark: {
        title: '⏸️ Kiểm Tra Dòng Trống Xen Kẽ',
        severity: 'pass',
        badge: 'Dữ Liệu Liền Mạch',
        content: 'Bảng dữ liệu liền mạch từ trên xuống dưới, không bị đứt đoạn bởi dòng trống.',
        action: null
      },
      actionShort: null
    };
  }

  // Rule 14: Date Formats
  function checkDateFormats(tpl, raw) {
    return {
      remark: {
        title: '📅 Kiểm Tra Định Dạng Ngày Tháng',
        severity: 'pass',
        badge: 'Định Dạng Ngày Hợp Lệ',
        content: 'Các trường ngày tháng không phát hiện xung đột định dạng cấu trúc.',
        action: null
      },
      actionShort: null
    };
  }

  // =========================================================================
  // 9. SYSTEM PROMPT & FEEDBACK GENERATOR
  // =========================================================================
  function switchPromptMode(mode) {
    state.promptMode = mode;
    if (el.btnModeSystemPrompt && el.btnModeHumanMsg) {
      el.btnModeSystemPrompt.classList.toggle('active', mode === 'system');
      el.btnModeHumanMsg.classList.toggle('active', mode === 'human');
    }
    updatePromptText();
  }

  function renderPromptFeedback(fileName, actionPoints, worstSeverity) {
    state.lastActionPoints = actionPoints;
    state.lastSeverity = worstSeverity;
    updatePromptText();
  }

  function updatePromptText() {
    const tpl = state.templateFile;
    const raw = state.rawFile;
    const actionPoints = state.lastActionPoints;
    const isSystemMode = state.promptMode === 'system';

    if (isSystemMode) {
      el.promptTextBox.value = buildSystemPromptText(tpl, raw, actionPoints);
    } else {
      el.promptTextBox.value = buildHumanMessageText(tpl, raw, actionPoints);
    }

    if (el.promptCharCount) {
      el.promptCharCount.textContent = `${el.promptTextBox.value.length} ký tự`;
    }
  }

  function buildSystemPromptText(tpl, raw, actionPoints) {
    const fileName = raw.name || 'File_Moi.xlsx';
    const tplName = tpl.name || 'Mau_Chuan.xlsx';

    if (actionPoints.length === 0) {
      return `File [${fileName}] đã chuẩn 100% theo mẫu [${tplName}]. Không có hạng mục nào cần xử lý.`;
    }

    let text = `YÊU CẦU XỬ LÝ CHUẨN HÓA FILE [${fileName}] THEO MẪU [${tplName}]:\n`;
    text += `(Giữ nguyên dữ liệu kỳ mới, chỉ xử lý chuẩn hóa cấu trúc và định dạng)\n\n`;
    text += `CÁC HẠNG MỤC CẦN XỬ LÝ:\n`;

    actionPoints.forEach((point, idx) => {
      text += `${idx + 1}. ${point}\n`;
    });

    return text.trim();
  }

  function buildHumanMessageText(tpl, raw, actionPoints) {
    const fileName = raw.name || 'File_Moi.xlsx';
    const tplName = tpl.name || 'Mau_Chuan.xlsx';

    if (actionPoints.length === 0) {
      return `File [${fileName}] đã chuẩn 100% theo mẫu [${tplName}]. Không cần chỉnh sửa gì thêm.`;
    }

    let text = `Chào anh/chị,\nNhờ anh/chị hỗ trợ xử lý các điểm sau cho file [${fileName}] để khớp đúng mẫu quy định:\n\n`;

    actionPoints.forEach((point, idx) => {
      text += `${idx + 1}. ${point}\n`;
    });

    text += `\nCảm ơn anh/chị!`;
    return text.trim();
  }

  function copyPromptToClipboard() {
    const text = el.promptTextBox.value || el.promptTextBox.textContent;
    if (!text) return;

    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(() => {
        showCopyToast('Đã copy nội dung prompt vào Clipboard!');
      }).catch(() => {
        fallbackCopy(text);
      });
    } else {
      fallbackCopy(text);
    }
  }

  function fallbackCopy(text) {
    const textarea = document.createElement('textarea');
    textarea.value = text;
    textarea.style.position = 'fixed';
    textarea.style.opacity = '0';
    document.body.appendChild(textarea);
    textarea.select();
    try {
      document.execCommand('copy');
      showCopyToast('Đã copy nội dung prompt vào Clipboard!');
    } catch (e) {
      alert('Không thể tự động copy, vui lòng bôi đen và copy thủ công.');
    }
    document.body.removeChild(textarea);
  }

  let toastTimer = null;
  function showCopyToast(msg) {
    if (toastTimer) clearTimeout(toastTimer);
    el.copyToastMsg.textContent = msg;
    el.copyToast.classList.add('show');
    toastTimer = setTimeout(() => {
      el.copyToast.classList.remove('show');
    }, 2500);
  }

  // =========================================================================
  // 10. RENDER REMARKS & DUAL EXCEL SPREADSHEETS
  // =========================================================================
  function renderExecutiveSummary(severity) {
    const banner = el.summaryBanner;
    banner.className = `executive-summary-banner ${severity}`;

    if (severity === 'fail') {
      el.summaryBannerIcon.textContent = '❌';
      el.summaryBannerTitle.textContent = 'KẾT LUẬN: FILE MỚI BỊ LỖI FORMAT NGHIÊM TRỌNG (CẦN YÊU CẦU SỬA)';
      el.summaryBannerDesc.textContent = 'File gửi mới bị thiếu cột bắt buộc của mẫu chuẩn hoặc có lỗi công thức. Bạn nên copy prompt bên trên để gửi yêu cầu đầu mối gửi lại file chuẩn.';
    } else if (severity === 'warning') {
      el.summaryBannerIcon.textContent = '⚠️';
      el.summaryBannerTitle.textContent = 'KẾT LUẬN: FILE MỚI BỊ ĐẢO CỘT / SAI TÊN TRƯỜNG (CÓ THỂ TỰ XỬ LÝ HOẶC BÁO SỬA)';
      el.summaryBannerDesc.textContent = 'File gửi mới có đủ dữ liệu nhưng bị xáo trộn thứ tự cột, viết tắt hoặc thụt lùi dòng tiêu đề. Bạn có thể bấm "Xuất File Chuẩn Hóa" để tải file đã sửa ngay, hoặc copy prompt gửi bên gửi.';
    } else {
      el.summaryBannerIcon.textContent = '✅';
      el.summaryBannerTitle.textContent = 'KẾT LUẬN: FORMAT ĐỒNG NHẤT 100% - FILE MỚI ĐẠT CHUẨN';
      el.summaryBannerDesc.textContent = 'Tất cả các tiêu chí format đã chọn đều vượt qua. File gửi mới có cấu trúc, thứ tự cột và tên trường hoàn toàn tương thích với file mẫu.';
    }
  }

  function renderRemarksList(remarks) {
    el.remarksList.innerHTML = '';

    remarks.forEach(rem => {
      const item = document.createElement('div');
      item.className = 'remark-entry';

      let badgeClass = 'badge-pass';
      if (rem.severity === 'fail') badgeClass = 'badge-fail';
      else if (rem.severity === 'warning') badgeClass = 'badge-warning';

      item.innerHTML = `
        <div class="remark-top">
          <div class="remark-title-area">
            <span>${rem.title}</span>
          </div>
          <span class="remark-status-badge ${badgeClass}">${rem.badge}</span>
        </div>
        <div class="remark-body">
          ${rem.content}
        </div>
        ${rem.action ? `<div class="remark-action">${rem.action}</div>` : ''}
      `;

      el.remarksList.appendChild(item);
    });
  }

  function renderDualExcelViewers(tpl, raw) {
    // 1. Template
    el.tplViewBadge.textContent = `${tpl.rows.length} dòng × ${tpl.headers.length} cột`;
    el.tplTableScroll.innerHTML = buildTableHtml(tpl.headers, tpl.rows);

    // 2. Raw
    el.rawViewBadge.textContent = `${raw.rows.length} dòng × ${raw.headers.length} cột (Header Dòng ${raw.headerRow})`;
    el.rawTableScroll.innerHTML = buildTableHtml(raw.headers, raw.rows);
  }

  function buildTableHtml(headers, rows) {
    if (!headers || headers.length === 0) {
      return '<div class="excel-empty-notice">Không có dữ liệu</div>';
    }

    const maxDisplayRows = Math.min(30, rows.length);
    let html = '<table class="excel-sheet-table"><thead><tr>';
    html += '<th class="col-row-num">#</th>';

    headers.forEach((h, idx) => {
      html += `<th><span style="font-size:10px; color:var(--text-dim); display:block;">[Cột ${idx + 1}]</span>${escapeHtml(h)}</th>`;
    });
    html += '</tr></thead><tbody>';

    for (let r = 0; r < maxDisplayRows; r++) {
      html += `<tr><td class="col-row-num">${r + 1}</td>`;
      for (let c = 0; c < headers.length; c++) {
        const val = rows[r] ? (rows[r][c] ?? '') : '';
        html += `<td title="${escapeHtml(val.toString())}">${escapeHtml(val.toString())}</td>`;
      }
      html += '</tr>';
    }

    if (rows.length > maxDisplayRows) {
      html += `<tr><td class="col-row-num">...</td><td colspan="${headers.length}" style="text-align:center; color:var(--text-dim); padding:10px;">(Còn ${rows.length - maxDisplayRows} dòng dữ liệu tiếp theo...)</td></tr>`;
    }

    html += '</tbody></table>';
    return html;
  }

  // =========================================================================
  // 11. EXPORT CLEAN STANDARDIZED EXCEL
  // =========================================================================
  function exportCleanExcel() {
    const rawRows = state.rawFile.rows;
    const tplHeaders = state.templateFile.headers;

    if (rawRows.length === 0) {
      alert('Không có dữ liệu trong file mới để xuất!');
      return;
    }

    const cleanData = [];
    cleanData.push(tplHeaders); // Row 1: Exactly template headers in order

    for (let r = 0; r < rawRows.length; r++) {
      const cleanRow = [];
      tplHeaders.forEach((th, tIdx) => {
        const match = state.matchedColumns.find(m => m.tIdx === tIdx);
        let cellVal = '';
        if (match && rawRows[r]) {
          cellVal = rawRows[r][match.rIdx] ?? '';
        }
        cleanRow.push(cellVal);
      });
      cleanData.push(cleanRow);
    }

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(cleanData);
    XLSX.utils.book_append_sheet(wb, ws, 'Clean_Data');

    const cleanFileName = `Chuan_Hoa_${state.rawFile.name || 'Du_Lieu.xlsx'}`;
    XLSX.writeFile(wb, cleanFileName);
  }



  function escapeHtml(str) {
    if (!str) return '';
    return str.toString().replace(/[&<>"']/g, function (m) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m];
    });
  }

  // =========================================================================
  // 13. INITIALIZATION
  // =========================================================================
  function init() {
    setupFileUploads();
    setupChecklistHandlers();
    updateChecksCounter();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

})();
