/**
 * =============================================================================
 * EXCEL SCHEMA NORMALIZER & COLUMN ALIGNMENT ENGINE
 * Chuyên xử lý:
 * - Chuẩn hóa tên cột bị lệch giữa các kỳ gửi dữ liệu
 * - Tự động phát hiện và ánh xạ (mapping) thông minh (Exact, Normalized, Từ điển nghiệp vụ, Fuzzy)
 * - Tái sắp xếp toàn bộ thứ tự cột theo đúng 100% thứ tự chuẩn của File Kỳ Gốc
 * - Giữ nguyên toàn bộ dữ liệu hợp lệ và xuất file Excel (.xlsx) sạch
 * =============================================================================
 */

(function () {
  'use strict';

  // State quản lý 2 file
  const state = {
    master: {
      file: null,
      workbook: null,
      selectedSheet: '',
      headerRowIndex: 0,
      headers: [],
      dataRows: [],
      rawRows: [],
      samples: {}
    },
    incoming: {
      file: null,
      workbook: null,
      selectedSheet: '',
      headerRowIndex: 0,
      headers: [],
      dataRows: [],
      rawRows: [],
      samples: {} // { colIndex: [sample1, sample2] }
    },
    // Ánh xạ: { masterColIndex: incomingColIndex | '__NONE__' | '__FIXED__' }
    mapping: {},
    fixedValues: {}, // { masterColIndex: 'custom value' }
    processedRows: [],
    schemaCompareOpen: true,
    rawViewMode: 'split'
  };

  // ===========================================================================
  // 1. TỪ ĐIỂN ĐỒNG NGHĨA NGHIỆP VỤ VIỆT NAM (BUSINESS SYNONYMS DICTIONARY)
  // ===========================================================================
  const SYNONYM_GROUPS = [
    {
      group: 'id',
      keywords: ['ma', 'code', 'id', 'manv', 'ma_nv', 'ma nhan vien', 'ma kh', 'makh', 'cust_id', 'so_hd', 'so hop dong', 'so_cmt', 'cccd', 'cmnd', 'dinh danh', 'ma so']
    },
    {
      group: 'name',
      keywords: ['ten', 'hoten', 'ho va ten', 'full_name', 'name', 'ten khach hang', 'ten_nv', 'ho ten nv', 'ho ten', 'ten day du', 'nguoi dai dien']
    },
    {
      group: 'phone',
      keywords: ['sdt', 'so dien thoai', 'phone', 'mobile', 'tel', 'lien he', 'so dt', 'telephone', 'dien thoai', 'so di dong', 'dt']
    },
    {
      group: 'email',
      keywords: ['email', 'thu dien tu', 'mail', 'e-mail', 'hop thu']
    },
    {
      group: 'address',
      keywords: ['dia chi', 'address', 'dia_chi', 'noi o', 'dia chi thuong tru', 'dia chi lien lac', 'tinh thanh', 'quan huyen', 'noi o hien tai']
    },
    {
      group: 'date',
      keywords: ['ngay', 'date', 'ngay sinh', 'ngay tao', 'ngay gd', 'ngay giao dich', 'created_date', 'dob', 'thoi gian', 'time', 'ngay thang', 'ngay lap']
    },
    {
      group: 'amount',
      keywords: ['doanh thu', 'doanh so', 'tien', 'so tien', 'amount', 'revenue', 'tong tien', 'thanh tien', 'price', 'don gia', 'gia', 'so tien thanh toan', 'gia tri', 'so du', 'du no']
    },
    {
      group: 'quantity',
      keywords: ['so luong', 'sl', 'qty', 'quantity', 'luot', 'khoi luong']
    },
    {
      group: 'dept',
      keywords: ['phong ban', 'bo phan', 'dept', 'department', 'chi nhanh', 'branch', 'don vi', 'khoi', 'trung tam', 'ban']
    },
    {
      group: 'role',
      keywords: ['chuc vu', 'position', 'role', 'vai tro', 'cap bac', 'title']
    },
    {
      group: 'gender',
      keywords: ['gioi tinh', 'gender', 'sex', 'phai']
    },
    {
      group: 'note',
      keywords: ['ghi chu', 'note', 'remark', 'mo ta', 'description', 'dien giai', 'y kien']
    },
    {
      group: 'status',
      keywords: ['trang thai', 'status', 'tinh trang', 'ket qua']
    }
  ];

  // ===========================================================================
  // 2. CHUẨN HÓA VĂN BẢN & ĐO ĐỘ TƯƠNG ĐỒNG (STRING NORMALIZATION & FUZZY)
  // ===========================================================================
  function stripVietnameseAccents(str) {
    if (!str) return '';
    return str
      .toString()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/đ/g, 'd')
      .replace(/Đ/g, 'D');
  }

  function cleanStringForMatch(str) {
    if (!str) return '';
    const withoutAccents = stripVietnameseAccents(str);
    return withoutAccents
      .toLowerCase()
      .replace(/[^a-z0-9]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  // Đo độ tương đồng chuỗi Levenshtein
  function calculateSimilarity(str1, str2) {
    const s1 = cleanStringForMatch(str1);
    const s2 = cleanStringForMatch(str2);

    if (s1 === s2) return 1.0;
    if (!s1 || !s2) return 0.0;

    // Kiểm tra từ điển đồng nghĩa
    for (const grp of SYNONYM_GROUPS) {
      const match1 = grp.keywords.some(k => s1.includes(k) || k.includes(s1));
      const match2 = grp.keywords.some(k => s2.includes(k) || k.includes(s2));
      if (match1 && match2) return 0.92;
    }

    // Levenshtein
    const track = Array(s2.length + 1)
      .fill(null)
      .map(() => Array(s1.length + 1).fill(null));

    for (let i = 0; i <= s1.length; i += 1) track[0][i] = i;
    for (let j = 0; j <= s2.length; j += 1) track[j][0] = j;

    for (let j = 1; j <= s2.length; j += 1) {
      for (let i = 1; i <= s1.length; i += 1) {
        const indicator = s1[i - 1] === s2[j - 1] ? 0 : 1;
        track[j][i] = Math.min(
          track[j][i - 1] + 1,
          track[j - 1][i] + 1,
          track[j - 1][i - 1] + indicator
        );
      }
    }

    const dist = track[s2.length][s1.length];
    const maxLen = Math.max(s1.length, s2.length);
    return Math.max(0, 1 - dist / maxLen);
  }

  function getColumnLetter(colIndex) {
    let temp, letter = '';
    while (colIndex >= 0) {
      temp = colIndex % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      colIndex = Math.floor(colIndex / 26) - 1;
    }
    return letter;
  }

  // ===========================================================================
  // 3. ĐỌC VÀ PHÂN TÍCH FILE EXCEL (SHEETJS ENGINE)
  // ===========================================================================
  function readFileAsync(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array', cellDates: true });
          resolve(workbook);
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  }

  function extractSheetData(workbook, sheetName, headerRowIdx = 0) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return { headers: [], dataRows: [], samples: {}, rawRows: [] };

    const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    if (!rawData || rawData.length === 0) return { headers: [], dataRows: [], samples: {}, rawRows: [] };

    const headerRow = rawData[headerRowIdx] || [];
    const headers = headerRow.map((h, i) => (h ? h.toString().trim() : `(Cột ${getColumnLetter(i)})`));
    const dataRows = rawData.slice(headerRowIdx + 1);

    // Thu thập mẫu dữ liệu 5 dòng đầu tiên cho mỗi cột
    const samples = {};
    headers.forEach((_, colIdx) => {
      const colSamples = [];
      for (let r = 0; r < Math.min(8, dataRows.length); r++) {
        const val = dataRows[r][colIdx];
        if (val !== undefined && val !== null && val.toString().trim() !== '') {
          colSamples.push(val.toString().trim());
        }
      }
      samples[colIdx] = colSamples;
    });

    return { headers, dataRows, samples, rawRows: rawData };
  }

  // Phỏng đoán kiểu dữ liệu của một cột dựa trên các giá trị mẫu
  function inferColumnType(samples = []) {
    if (!samples || samples.length === 0) {
      return { label: 'Trống / Rỗng', cls: 'type-empty' };
    }

    let numCount = 0;
    let dateCount = 0;

    samples.forEach(s => {
      const str = String(s).trim();
      if (!str) return;
      // Nhận diện ngày tháng
      if (/\d{4}[-/.]\d{1,2}[-/.]\d{1,2}/.test(str) || /\d{1,2}[-/.]\d{1,2}[-/.]\d{4}/.test(str)) {
        dateCount++;
      } else if (!isNaN(Number(str.replace(/,/g, '')))) {
        numCount++;
      }
    });

    if (dateCount > 0 && dateCount >= samples.length * 0.7) {
      return { label: 'Ngày (Date)', cls: 'type-date' };
    }
    if (numCount > 0 && numCount >= samples.length * 0.7) {
      return { label: 'Số (Number)', cls: 'type-number' };
    }
    if (numCount > 0 && numCount < samples.length) {
      return { label: 'Hỗn hợp', cls: 'type-string' };
    }
    return { label: 'Văn bản (Text)', cls: 'type-string' };
  }

  // ===========================================================================
  // 4. THUẬT TOÁN TỰ ĐỘNG ÁNH XẠ THÔNG MINH (AUTO-MAPPING)
  // ===========================================================================
  function runAutoMapping() {
    const masterHeaders = state.master.headers;
    const incomingHeaders = state.incoming.headers;
    const newMapping = {};

    const usedIncomingIndices = new Set();

    // Vòng 1: Khớp chính xác 100% (Exact match)
    masterHeaders.forEach((mHeader, mIdx) => {
      const cleanM = cleanStringForMatch(mHeader);
      incomingHeaders.forEach((iHeader, iIdx) => {
        if (usedIncomingIndices.has(iIdx)) return;
        const cleanI = cleanStringForMatch(iHeader);
        if (cleanM === cleanI && cleanM !== '') {
          newMapping[mIdx] = iIdx;
          usedIncomingIndices.add(iIdx);
        }
      });
    });

    // Vòng 2: Khớp độ tương đồng cao (Fuzzy & Synonyms >= 0.72)
    masterHeaders.forEach((mHeader, mIdx) => {
      if (newMapping[mIdx] !== undefined) return;

      let bestScore = 0;
      let bestIdx = -1;

      incomingHeaders.forEach((iHeader, iIdx) => {
        if (usedIncomingIndices.has(iIdx)) return;
        const score = calculateSimilarity(mHeader, iHeader);
        if (score > bestScore && score >= 0.65) {
          bestScore = score;
          bestIdx = iIdx;
        }
      });

      if (bestIdx !== -1) {
        newMapping[mIdx] = bestIdx;
        usedIncomingIndices.add(bestIdx);
      } else {
        newMapping[mIdx] = '__NONE__';
      }
    });

    state.mapping = newMapping;
    renderMappingTable();
    updateActionButtons();
  }

  // ===========================================================================
  // 5. RENDER BẢNG MAPPING TRỰC QUAN (INTERACTIVE MAPPING MATRIX)
  // ===========================================================================
  function renderMappingTable() {
    const tableBody = document.getElementById('mappingTableBody');
    const counterBadge = document.getElementById('mappingCounterBadge');
    if (!tableBody) return;

    tableBody.innerHTML = '';
    const masterHeaders = state.master.headers;
    const incomingHeaders = state.incoming.headers;

    let mappedCount = 0;

    masterHeaders.forEach((mHeader, mIdx) => {
      const mappedIncomingIdx = state.mapping[mIdx];
      const isMapped = mappedIncomingIdx !== undefined && mappedIncomingIdx !== '__NONE__';
      if (isMapped) mappedCount++;

      let similarityScore = 0;
      if (isMapped) {
        similarityScore = calculateSimilarity(mHeader, incomingHeaders[mappedIncomingIdx]);
      }

      // Sample preview
      let sampleText = '—';
      if (isMapped && state.incoming.samples[mappedIncomingIdx]) {
        sampleText = state.incoming.samples[mappedIncomingIdx].slice(0, 2).join(' | ') || '(Không có mẫu)';
      }

      const tr = document.createElement('tr');

      // Tùy chọn trong dropdown
      const optionsHtml = [
        `<option value="__NONE__" ${mappedIncomingIdx === '__NONE__' ? 'selected' : ''}>⚠️ [Bỏ trống / Điền rỗng]</option>`,
        ...incomingHeaders.map((iHeader, iIdx) => {
          const isSelected = mappedIncomingIdx === iIdx;
          const score = calculateSimilarity(mHeader, iHeader);
          const pct = Math.round(score * 100);
          return `<option value="${iIdx}" ${isSelected ? 'selected' : ''}>
            ${getColumnLetter(iIdx)}. ${escapeHtml(iHeader)} (${pct}% khớp)
          </option>`;
        })
      ].join('');

      let statusBadgeHtml = '';
      if (!isMapped) {
        statusBadgeHtml = '<span class="status-badge empty">✕ Chưa ánh xạ</span>';
      } else if (similarityScore >= 0.95) {
        statusBadgeHtml = '<span class="status-badge exact">✓ Khớp chuẩn 100%</span>';
      } else {
        const pct = Math.round(similarityScore * 100);
        statusBadgeHtml = `<span class="status-badge fuzzy">⚡ Tự động khớp ${pct}%</span>`;
      }

      tr.innerHTML = `
        <td style="text-align: center; width: 60px;">
          <span class="col-num-badge">${getColumnLetter(mIdx)}</span>
        </td>
        <td style="width: 260px;">
          <div class="target-col-name">
            <b>${escapeHtml(mHeader)}</b>
          </div>
        </td>
        <td style="text-align: center; width: 40px;">
          <span class="arrow-icon">➔</span>
        </td>
        <td style="width: 320px;">
          <select class="mapping-source-select ${isMapped ? 'mapped' : 'unmapped'}" data-master-idx="${mIdx}">
            ${optionsHtml}
          </select>
        </td>
        <td>
          <div class="sample-value-preview" title="Dữ liệu mẫu từ cột nguồn: ${escapeHtml(sampleText)}">
            ${escapeHtml(sampleText)}
          </div>
        </td>
        <td style="width: 170px;">
          ${statusBadgeHtml}
        </td>
      `;

      // Gắn sự kiện khi người dùng thay đổi lựa chọn dropdown
      const selectElem = tr.querySelector('.mapping-source-select');
      selectElem.addEventListener('change', (e) => {
        const val = e.target.value;
        state.mapping[mIdx] = val === '__NONE__' ? '__NONE__' : parseInt(val, 10);
        renderMappingTable();
        updateActionButtons();
      });

      tableBody.appendChild(tr);
    });

    if (counterBadge) {
      counterBadge.textContent = `Đã khớp: ${mappedCount} / ${masterHeaders.length} cột`;
    }
  }

  // ===========================================================================
  // 6. XỬ LÝ DỮ LIỆU & TÁI CẤU TRÚC THEO KỲ GỐC (NORMALIZATION & ALIGNMENT)
  // ===========================================================================
  function processAndAlignData() {
    const masterHeaders = state.master.headers;
    const incomingDataRows = state.incoming.dataRows;

    const optTrim = document.getElementById('optTrimSpaces')?.checked ?? true;
    const optRemoveBlank = document.getElementById('optRemoveBlankRows')?.checked ?? true;
    const optCleanNum = document.getElementById('optCleanNumbers')?.checked ?? true;

    const resultRows = [];
    // Hàng tiêu đề 100% chuẩn theo File Kỳ Gốc!
    resultRows.push([...masterHeaders]);

    for (let r = 0; r < incomingDataRows.length; r++) {
      const rawRow = incomingDataRows[r];

      // Kiểm tra dòng hoàn toàn rỗng
      const hasAnyValue = rawRow.some(val => val !== undefined && val !== null && val.toString().trim() !== '');
      if (optRemoveBlank && !hasAnyValue) {
        continue;
      }

      // Tạo dòng mới đã sắp xếp đúng thứ tự cột của Kỳ Gốc
      const newRow = masterHeaders.map((_, mIdx) => {
        const mappedIncomingIdx = state.mapping[mIdx];
        if (mappedIncomingIdx === undefined || mappedIncomingIdx === '__NONE__') {
          return ''; // Bỏ trống theo cấu hình
        }

        let cellVal = rawRow[mappedIncomingIdx];
        if (cellVal === undefined || cellVal === null) return '';

        if (optTrim && typeof cellVal === 'string') {
          cellVal = cellVal.trim();
        }

        // Chuẩn hóa số lưu dạng text (loại bỏ dấu phân cách nghìn)
        if (optCleanNum && typeof cellVal === 'string') {
          const numCandidate = cellVal.replace(/,/g, '');
          if (!isNaN(numCandidate) && numCandidate.trim() !== '') {
            cellVal = Number(numCandidate);
          }
        }

        return cellVal;
      });

      resultRows.push(newRow);
    }

    state.processedRows = resultRows;
    return resultRows;
  }

  // ===========================================================================
  // 7. XUẤT FILE EXCEL ĐÃ CHUẨN HÓA (EXPORT FILE .XLSX)
  // ===========================================================================
  function exportCleanExcel() {
    const rows = processAndAlignData();
    if (!rows || rows.length <= 1) {
      alert('Không có dữ liệu hợp lệ để xuất file.');
      return;
    }

    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();

    const sheetName = state.master.selectedSheet || 'Data_Chuan_Hoa';
    XLSX.utils.book_append_sheet(wb, ws, sheetName);

    const baseName = state.incoming.file ? state.incoming.file.name.replace(/\.[^/.]+$/, '') : 'Du_Lieu';
    const exportFileName = `${baseName}_Chuan_Hoa_Theo_Ky_Goc.xlsx`;

    XLSX.writeFile(wb, exportFileName);

    if (window.showToast) {
      window.showToast(`✅ Đã xuất thành công: ${exportFileName}`);
    } else {
      alert(`Đã xuất thành công file Excel chuẩn hóa: ${exportFileName}`);
    }

    // Hiển thị live preview
    renderOutputPreview();
  }

  // ===========================================================================
  // 8. LIVE PREVIEW TABLE (XEM TRƯỚC BẢNG KẾT QUẢ SẮP XẾP)
  // ===========================================================================
  function renderOutputPreview() {
    const rows = state.processedRows.length > 0 ? state.processedRows : processAndAlignData();
    const previewSection = document.getElementById('outputPreviewSection');
    const tableHead = document.getElementById('outputTableHead');
    const tableBody = document.getElementById('outputTableBody');
    const statsBadge = document.getElementById('outputStatsBadge');

    if (!previewSection || !tableHead || !tableBody) return;

    previewSection.style.display = 'block';
    tableHead.innerHTML = '';
    tableBody.innerHTML = '';

    const headers = rows[0] || [];
    const dataRows = rows.slice(1, 101); // Preview 100 dòng đầu tiên

    if (statsBadge) {
      statsBadge.textContent = `${rows.length - 1} dòng x ${headers.length} cột`;
    }

    // Header row
    const trHead = document.createElement('tr');
    trHead.innerHTML = `
      <th style="width: 50px; text-align: center;">STT</th>
      ${headers.map((h, i) => `<th>${getColumnLetter(i)}. ${escapeHtml(h)}</th>`).join('')}
    `;
    tableHead.appendChild(trHead);

    // Data rows preview
    dataRows.forEach((row, rIdx) => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td style="text-align: center; color: var(--text-dim);">${rIdx + 1}</td>
        ${row.map(cell => `<td>${escapeHtml(cell !== undefined && cell !== null ? cell.toString() : '')}</td>`).join('')}
      `;
      tableBody.appendChild(tr);
    });

    previewSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  // ===========================================================================
  // 9. CẬP NHẬT TRẠNG THÁI NÚT VÀ GIAO DIỆN
  // ===========================================================================
  function updateActionButtons() {
    const btnRun = document.getElementById('btnRunNormalize');
    const btnPreview = document.getElementById('btnPreviewOutput');
    const canRun = state.master.headers.length > 0 && state.incoming.headers.length > 0;

    if (btnRun) btnRun.disabled = !canRun;
    if (btnPreview) btnPreview.disabled = !canRun;
  }

  function updateFileCardUI(type) {
    const fileData = state[type];
    const isMaster = type === 'master';
    const dropzone = document.getElementById(isMaster ? 'masterDropzone' : 'incomingDropzone');
    const metaBox = document.getElementById(isMaster ? 'masterMetaBox' : 'incomingMetaBox');
    const fileNameDisplay = document.getElementById(isMaster ? 'masterFileName' : 'incomingFileName');
    const sheetSelect = document.getElementById(isMaster ? 'masterSheetSelect' : 'incomingSheetSelect');
    const headerRowSelect = document.getElementById(isMaster ? 'masterHeaderRowSelect' : 'incomingHeaderRowSelect');
    const statsContainer = document.getElementById(isMaster ? 'masterStats' : 'incomingStats');

    if (!fileData.file) {
      if (dropzone) dropzone.style.display = 'flex';
      if (metaBox) metaBox.style.display = 'none';
      return;
    }

    if (dropzone) dropzone.style.display = 'none';
    if (metaBox) metaBox.style.display = 'flex';
    if (fileNameDisplay) fileNameDisplay.textContent = fileData.file.name;

    // Fill sheet options
    if (sheetSelect && fileData.workbook) {
      sheetSelect.innerHTML = fileData.workbook.SheetNames.map(
        s => `<option value="${escapeHtml(s)}" ${s === fileData.selectedSheet ? 'selected' : ''}>${escapeHtml(s)}</option>`
      ).join('');
    }

    // Fill header row options (Dòng 1 đến dòng 10)
    if (headerRowSelect) {
      headerRowSelect.innerHTML = [0, 1, 2, 3, 4].map(
        idx => `<option value="${idx}" ${idx === fileData.headerRowIndex ? 'selected' : ''}>Dòng ${idx + 1}</option>`
      ).join('');
    }

    if (statsContainer) {
      statsContainer.innerHTML = `
        <span class="stat-pill">📊 ${fileData.headers.length} cột</span>
        <span class="stat-pill">📋 ${fileData.dataRows.length} dòng</span>
      `;
    }

    // Kiểm tra xem cả 2 file đã nạp chưa -> hiển thị mapping matrix
    const mappingSection = document.getElementById('mappingSectionCard');
    if (state.master.headers.length > 0 && state.incoming.headers.length > 0) {
      if (mappingSection) mappingSection.style.display = 'block';
      runAutoMapping();
    } else {
      if (mappingSection) mappingSection.style.display = 'none';
    }

    renderSchemaComparison();
    updateActionButtons();
  }

  // ===========================================================================
  // 10. SO SÁNH CẤU TRÚC 2 FILE MẪU (SCHEMA INSPECTOR & RAW DATA PREVIEW)
  // ===========================================================================
  function renderSchemaComparison() {
    const triggerBar = document.getElementById('compareTriggerBar');
    const compareSection = document.getElementById('schemaCompareSection');
    const summaryElem = document.getElementById('compareTriggerSummary');
    const toggleBtnText = document.getElementById('btnToggleCompareText');

    const hasMaster = state.master.headers.length > 0;
    const hasIncoming = state.incoming.headers.length > 0;

    if (!hasMaster && !hasIncoming) {
      if (triggerBar) triggerBar.style.display = 'none';
      if (compareSection) compareSection.style.display = 'none';
      return;
    }

    if (triggerBar) triggerBar.style.display = 'flex';
    if (compareSection) {
      compareSection.style.display = state.schemaCompareOpen ? 'block' : 'none';
    }
    if (toggleBtnText) {
      toggleBtnText.textContent = state.schemaCompareOpen ? 'Ẩn Cấu Trúc File ▲' : 'Xem Chi Tiết Cấu Trúc File ▼';
    }

    const mLen = state.master.headers.length;
    const iLen = state.incoming.headers.length;
    let diffStr = '';
    if (hasMaster && hasIncoming) {
      if (iLen > mLen) {
        diffStr = ` • Kỳ Mới thừa ${iLen - mLen} cột`;
      } else if (iLen < mLen) {
        diffStr = ` • Kỳ Mới thiếu ${mLen - iLen} cột`;
      } else {
        diffStr = ` • Cùng số lượng (${mLen} cột)`;
      }
    }
    if (summaryElem) {
      summaryElem.textContent = `Kỳ Gốc: ${mLen} cột • Kỳ Mới: ${iLen} cột${diffStr}`;
    }

    // -------------------------------------------------------------------------
    // RENDER TAB 1: SO SÁNH DANH SÁCH CỘT SONG SONG
    // -------------------------------------------------------------------------
    const masterTbody = document.getElementById('schemaMasterTbody');
    const incomingTbody = document.getElementById('schemaIncomingTbody');
    const masterName = document.getElementById('schemaMasterName');
    const incomingName = document.getElementById('schemaIncomingName');
    const masterMeta = document.getElementById('schemaMasterMeta');
    const incomingMeta = document.getElementById('schemaIncomingMeta');

    // Cột bên trái: Kỳ Gốc
    if (masterName) masterName.textContent = state.master.file ? state.master.file.name : '—';
    if (masterMeta) masterMeta.textContent = `${mLen} cột • ${state.master.dataRows.length} dòng dữ liệu`;

    if (masterTbody) {
      if (!hasMaster) {
        masterTbody.innerHTML = '<tr><td colspan="4" class="empty-cell" style="text-align:center; padding: 24px; color: var(--text-dim);">Chưa nạp file kỳ gốc</td></tr>';
      } else {
        masterTbody.innerHTML = state.master.headers.map((h, idx) => {
          const typeInfo = inferColumnType(state.master.samples[idx]);
          const sampleText = state.master.samples[idx] && state.master.samples[idx].length > 0 
            ? state.master.samples[idx].slice(0, 2).join(' | ') 
            : '—';
          return `
            <tr>
              <td style="text-align: center; font-weight: 700; color: var(--primary);">${getColumnLetter(idx)}</td>
              <td><b>${escapeHtml(h)}</b></td>
              <td><span class="schema-type-badge ${typeInfo.cls}">${typeInfo.label}</span></td>
              <td><div class="schema-sample-text" title="${escapeHtml(sampleText)}">${escapeHtml(sampleText)}</div></td>
            </tr>
          `;
        }).join('');
      }
    }

    // Cột bên phải: Kỳ Mới
    if (incomingName) incomingName.textContent = state.incoming.file ? state.incoming.file.name : '—';
    if (incomingMeta) incomingMeta.textContent = `${iLen} cột • ${state.incoming.dataRows.length} dòng dữ liệu`;

    if (incomingTbody) {
      if (!hasIncoming) {
        incomingTbody.innerHTML = '<tr><td colspan="5" class="empty-cell" style="text-align:center; padding: 24px; color: var(--text-dim);">Chưa nạp file kỳ mới</td></tr>';
      } else {
        incomingTbody.innerHTML = state.incoming.headers.map((h, idx) => {
          const typeInfo = inferColumnType(state.incoming.samples[idx]);
          const sampleText = state.incoming.samples[idx] && state.incoming.samples[idx].length > 0 
            ? state.incoming.samples[idx].slice(0, 2).join(' | ') 
            : '—';

          // Kiểm tra tình trạng đối soát với kỳ gốc
          let matchStatusHtml = '<span class="match-status-pill extra">🆕 Cột ngoài mẫu</span>';
          if (hasMaster) {
            let maxSim = 0;
            let matchedHeader = '';
            state.master.headers.forEach(mH => {
              const sim = calculateSimilarity(mH, h);
              if (sim > maxSim) {
                maxSim = sim;
                matchedHeader = mH;
              }
            });

            if (maxSim >= 0.95) {
              matchStatusHtml = '<span class="match-status-pill exact">✓ Trùng tên gốc</span>';
            } else if (maxSim >= 0.65) {
              matchStatusHtml = `<span class="match-status-pill fuzzy" title="Tương đồng với: ${escapeHtml(matchedHeader)}">⚡ Gần giống (${Math.round(maxSim * 100)}%)</span>`;
            }
          }

          return `
            <tr>
              <td style="text-align: center; font-weight: 700; color: #f59e0b;">${getColumnLetter(idx)}</td>
              <td><b>${escapeHtml(h)}</b></td>
              <td><span class="schema-type-badge ${typeInfo.cls}">${typeInfo.label}</span></td>
              <td><div class="schema-sample-text" title="${escapeHtml(sampleText)}">${escapeHtml(sampleText)}</div></td>
              <td>${matchStatusHtml}</td>
            </tr>
          `;
        }).join('');
      }
    }

    // -------------------------------------------------------------------------
    // RENDER TAB 2: XEM BẢNG DỮ LIỆU THÔ CỦA 2 FILE
    // -------------------------------------------------------------------------
    renderRawDataTables();
  }

  function renderRawDataTables() {
    const rawMasterTbody = document.getElementById('rawMasterTbody');
    const rawIncomingTbody = document.getElementById('rawIncomingTbody');
    const rawMasterSheetLabel = document.getElementById('rawMasterSheetLabel');
    const rawIncomingSheetLabel = document.getElementById('rawIncomingSheetLabel');
    const rawMasterHeaderIndexLabel = document.getElementById('rawMasterHeaderIndexLabel');
    const rawIncomingHeaderIndexLabel = document.getElementById('rawIncomingHeaderIndexLabel');

    if (rawMasterSheetLabel) rawMasterSheetLabel.textContent = state.master.selectedSheet || 'Sheet1';
    if (rawMasterHeaderIndexLabel) rawMasterHeaderIndexLabel.textContent = `Header ở dòng ${state.master.headerRowIndex + 1}`;

    if (rawIncomingSheetLabel) rawIncomingSheetLabel.textContent = state.incoming.selectedSheet || 'Sheet1';
    if (rawIncomingHeaderIndexLabel) rawIncomingHeaderIndexLabel.textContent = `Header ở dòng ${state.incoming.headerRowIndex + 1}`;

    // Render Raw Master
    if (rawMasterTbody) {
      if (!state.master.rawRows || state.master.rawRows.length === 0) {
        rawMasterTbody.innerHTML = '<tr><td style="text-align:center; padding: 20px; color: var(--text-dim);">Chưa nạp file kỳ gốc</td></tr>';
      } else {
        const previewRows = state.master.rawRows.slice(0, 10);
        rawMasterTbody.innerHTML = previewRows.map((row, rIdx) => {
          const isHeader = rIdx === state.master.headerRowIndex;
          const trClass = isHeader ? 'class="header-row"' : '';
          const rowTag = isHeader ? ' <span style="font-size:10px; padding:1px 4px; border-radius:3px; background:var(--primary); color:#fff; margin-left:4px;">TIÊU ĐỀ</span>' : '';
          return `
            <tr ${trClass}>
              <td style="text-align:center; font-weight:700; width:65px; background:var(--bg-card);">${rIdx + 1}${rowTag}</td>
              ${row.map(c => `<td>${escapeHtml(c !== undefined && c !== null ? c.toString() : '')}</td>`).join('')}
            </tr>
          `;
        }).join('');
      }
    }

    // Render Raw Incoming
    if (rawIncomingTbody) {
      if (!state.incoming.rawRows || state.incoming.rawRows.length === 0) {
        rawIncomingTbody.innerHTML = '<tr><td style="text-align:center; padding: 20px; color: var(--text-dim);">Chưa nạp file kỳ mới</td></tr>';
      } else {
        const previewRows = state.incoming.rawRows.slice(0, 10);
        rawIncomingTbody.innerHTML = previewRows.map((row, rIdx) => {
          const isHeader = rIdx === state.incoming.headerRowIndex;
          const trClass = isHeader ? 'class="header-row"' : '';
          const rowTag = isHeader ? ' <span style="font-size:10px; padding:1px 4px; border-radius:3px; background:#f59e0b; color:#fff; margin-left:4px;">TIÊU ĐỀ</span>' : '';
          return `
            <tr ${trClass}>
              <td style="text-align:center; font-weight:700; width:65px; background:var(--bg-card);">${rIdx + 1}${rowTag}</td>
              ${row.map(c => `<td>${escapeHtml(c !== undefined && c !== null ? c.toString() : '')}</td>`).join('')}
            </tr>
          `;
        }).join('');
      }
    }
  }

  // ===========================================================================
  // 11. GẮN KẾT SỰ KIỆN KÉO THẢ & NẠP FILE
  // ===========================================================================
  function setupFileHandlers(type) {
    const isMaster = type === 'master';
    const dropzone = document.getElementById(isMaster ? 'masterDropzone' : 'incomingDropzone');
    const fileInput = document.getElementById(isMaster ? 'masterFileInput' : 'incomingFileInput');
    const btnRemove = document.getElementById(isMaster ? 'masterBtnRemove' : 'incomingBtnRemove');
    const sheetSelect = document.getElementById(isMaster ? 'masterSheetSelect' : 'incomingSheetSelect');
    const headerRowSelect = document.getElementById(isMaster ? 'masterHeaderRowSelect' : 'incomingHeaderRowSelect');

    if (!dropzone || !fileInput) return;

    dropzone.addEventListener('click', () => fileInput.click());

    dropzone.addEventListener('dragover', (e) => {
      e.preventDefault();
      dropzone.classList.add('dragover');
    });

    dropzone.addEventListener('dragleave', () => {
      dropzone.classList.remove('dragover');
    });

    dropzone.addEventListener('drop', (e) => {
      e.preventDefault();
      dropzone.classList.remove('dragover');
      if (e.dataTransfer.files && e.dataTransfer.files[0]) {
        handleFileSelect(type, e.dataTransfer.files[0]);
      }
    });

    fileInput.addEventListener('change', (e) => {
      if (e.target.files && e.target.files[0]) {
        handleFileSelect(type, e.target.files[0]);
      }
    });

    if (btnRemove) {
      btnRemove.addEventListener('click', () => {
        state[type].file = null;
        state[type].workbook = null;
        state[type].headers = [];
        state[type].dataRows = [];
        state[type].rawRows = [];
        state[type].samples = {};
        fileInput.value = '';
        updateFileCardUI(type);
      });
    }

    if (sheetSelect) {
      sheetSelect.addEventListener('change', (e) => {
        state[type].selectedSheet = e.target.value;
        const res = extractSheetData(state[type].workbook, state[type].selectedSheet, state[type].headerRowIndex);
        state[type].headers = res.headers;
        state[type].dataRows = res.dataRows;
        state[type].rawRows = res.rawRows;
        state[type].samples = res.samples;
        updateFileCardUI(type);
      });
    }

    if (headerRowSelect) {
      headerRowSelect.addEventListener('change', (e) => {
        state[type].headerRowIndex = parseInt(e.target.value, 10);
        const res = extractSheetData(state[type].workbook, state[type].selectedSheet, state[type].headerRowIndex);
        state[type].headers = res.headers;
        state[type].dataRows = res.dataRows;
        state[type].rawRows = res.rawRows;
        state[type].samples = res.samples;
        updateFileCardUI(type);
      });
    }
  }

  async function handleFileSelect(type, file) {
    try {
      const workbook = await readFileAsync(file);
      state[type].file = file;
      state[type].workbook = workbook;
      state[type].selectedSheet = workbook.SheetNames[0] || '';
      state[type].headerRowIndex = 0;

      const res = extractSheetData(workbook, state[type].selectedSheet, 0);
      state[type].headers = res.headers;
      state[type].dataRows = res.dataRows;
      state[type].rawRows = res.rawRows;
      state[type].samples = res.samples;

      updateFileCardUI(type);
    } catch (err) {
      console.error('Lỗi khi đọc file Excel:', err);
      alert('Không thể mở file Excel này. Vui lòng kiểm tra định dạng .xlsx, .xls hoặc .csv!');
    }
  }

  function escapeHtml(str) {
    if (!str) return '';
    return str
      .toString()
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  // ===========================================================================
  // 12. KHỞI TẠO BAN ĐẦU
  // ===========================================================================
  function init() {
    setupFileHandlers('master');
    setupFileHandlers('incoming');

    // Nút toggle đóng mở phần so sánh cấu trúc
    const btnToggleCompare = document.getElementById('btnToggleCompare');
    if (btnToggleCompare) {
      btnToggleCompare.addEventListener('click', () => {
        state.schemaCompareOpen = !state.schemaCompareOpen;
        const compareSection = document.getElementById('schemaCompareSection');
        const toggleBtnText = document.getElementById('btnToggleCompareText');
        if (compareSection) {
          compareSection.style.display = state.schemaCompareOpen ? 'block' : 'none';
        }
        if (toggleBtnText) {
          toggleBtnText.textContent = state.schemaCompareOpen ? 'Ẩn Cấu Trúc File ▲' : 'Xem Chi Tiết Cấu Trúc File ▼';
        }
      });
    }

    // Switch Tabs trong Schema Compare (Tab 1: Cột, Tab 2: Dữ liệu thô)
    const tabBtnColumns = document.getElementById('tabBtnColumns');
    const tabBtnRaw = document.getElementById('tabBtnRaw');
    const tabContentSplitColumns = document.getElementById('tabContentSplitColumns');
    const tabContentRawData = document.getElementById('tabContentRawData');

    if (tabBtnColumns && tabBtnRaw) {
      tabBtnColumns.addEventListener('click', () => {
        tabBtnColumns.classList.add('active');
        tabBtnRaw.classList.remove('active');
        if (tabContentSplitColumns) tabContentSplitColumns.style.display = 'block';
        if (tabContentRawData) tabContentRawData.style.display = 'none';
      });

      tabBtnRaw.addEventListener('click', () => {
        tabBtnRaw.classList.add('active');
        tabBtnColumns.classList.remove('active');
        if (tabContentRawData) tabContentRawData.style.display = 'block';
        if (tabContentSplitColumns) tabContentSplitColumns.style.display = 'none';
        renderRawDataTables();
      });
    }

    // Nút đổi chế độ xem dữ liệu thô (Song song / Kỳ Gốc / Kỳ Mới)
    const rawModeBtns = document.querySelectorAll('.raw-view-btn');
    const rawDataSheetsWrapper = document.getElementById('rawDataSheetsWrapper');
    rawModeBtns.forEach(btn => {
      btn.addEventListener('click', () => {
        rawModeBtns.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        const mode = btn.dataset.rawMode || 'split';
        state.rawViewMode = mode;
        if (rawDataSheetsWrapper) {
          rawDataSheetsWrapper.className = `raw-data-sheets-wrapper mode-${mode}`;
        }
      });
    });

    // Nút tự động khớp lại
    const btnAutoMatch = document.getElementById('btnAutoMatch');
    if (btnAutoMatch) {
      btnAutoMatch.addEventListener('click', runAutoMapping);
    }

    // Nút xóa toàn bộ ánh xạ
    const btnClearMapping = document.getElementById('btnClearMapping');
    if (btnClearMapping) {
      btnClearMapping.addEventListener('click', () => {
        state.master.headers.forEach((_, idx) => {
          state.mapping[idx] = '__NONE__';
        });
        renderMappingTable();
        updateActionButtons();
      });
    }

    // Nút Thực thi xuất file
    const btnRun = document.getElementById('btnRunNormalize');
    if (btnRun) {
      btnRun.addEventListener('click', exportCleanExcel);
    }

    // Nút xem trước Live Preview
    const btnPreview = document.getElementById('btnPreviewOutput');
    if (btnPreview) {
      btnPreview.addEventListener('click', renderOutputPreview);
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

})();
