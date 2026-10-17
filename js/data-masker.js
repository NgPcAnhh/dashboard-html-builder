/**
 * =============================================================================
 * DATA MASKER & ANONYMIZER ENGINE (EXCEL & CSV)
 * - Che mờ dữ liệu nhạy cảm theo quy tắc:
 *   + Ô dữ liệu < 5 ký tự: che toàn bộ bằng '*****'
 *   + Ô dữ liệu > 5 ký tự: giữ nguyên 3 ký tự đầu, các ký tự sau che mờ hết (ví dụ: abc****)
 *   + Xử lý ô đúng 5 ký tự: linh hoạt theo cấu hình (mặc định: '*****')
 *   + Option chọn dòng đầu làm Header (giữ nguyên không che tiêu đề cột)
 *   + Hỗ trợ đa sheet Excel (.xlsx, .xls) và file CSV
 *   + Bảng xem trước so sánh Before / After và xuất file an toàn
 * =============================================================================
 */

(function () {
  'use strict';

  // State quản lý của ứng dụng
  const state = {
    file: null,
    fileName: '',
    fileType: '', // 'xlsx', 'xls', 'csv'
    workbook: null,
    sheetNames: [],
    selectedSheets: new Set(),
    activePreviewSheet: '',
    // Dữ liệu từng sheet: { [sheetName]: { rawRows: [[]], maskedRows: [[]], headers: [], stats: {} } }
    sheetData: {},
    options: {
      hasHeader: true,           // Dòng đầu tiên là Header (giữ nguyên không che)
      headerRowIndex: 0,         // Dòng số mấy là header (0-indexed)
      maskLessThan5: '*****',    // Ký tự che mờ cho chuỗi < 5 ký tự
      fiveCharsRule: 'mask_all', // 'mask_all' (*****) hoặc 'keep_3' (abc**)
      starPaddingMode: 'exact',  // 'exact' (* lặp theo số ký tự còn lại) hoặc 'fixed4' (cố định 4 sao)
      maskChar: '*',             // Ký tự che mờ
      maskAllColumns: true,      // Che tất cả các cột
      selectedColumns: {}        // { [sheetName]: Set([colIndex1, colIndex2]) }
    },
    preview: {
      viewMode: 'masked', // 'masked', 'original'
      searchTerm: '',
      currentPage: 1,
      pageSize: 20
    },
    totalStats: {
      processedSheets: 0,
      totalRows: 0,
      totalCells: 0,
      maskedCells: 0,
      executionTimeMs: 0
    }
  };

  // DOM Elements
  const dom = {};

  function initDOM() {
    dom.fileInput = document.getElementById('fileInput');
    dom.dropzone = document.getElementById('dropzone');
    dom.fileInfoBar = document.getElementById('fileInfoBar');
    dom.fileNameDisplay = document.getElementById('fileNameDisplay');
    dom.fileDetailsDisplay = document.getElementById('fileDetailsDisplay');
    dom.btnRemoveFile = document.getElementById('btnRemoveFile');

    // Options elements
    dom.chkHasHeader = document.getElementById('chkHasHeader');
    dom.selectHeaderRow = document.getElementById('selectHeaderRow');
    dom.radioFiveAll = document.getElementById('radioFiveAll');
    dom.radioFiveKeep3 = document.getElementById('radioFiveKeep3');
    dom.radioStarExact = document.getElementById('radioStarExact');
    dom.radioStarFixed4 = document.getElementById('radioStarFixed4');
    dom.sheetSelectionBox = document.getElementById('sheetSelectionBox');
    dom.sheetChipsContainer = document.getElementById('sheetChipsContainer');
    dom.btnSelectAllSheets = document.getElementById('btnSelectAllSheets');

    // Playground elements
    dom.playgroundInput = document.getElementById('playgroundInput');
    dom.playgroundResult = document.getElementById('playgroundResult');
    dom.samplePills = document.querySelectorAll('.sample-pill');

    // Action buttons
    dom.btnExecuteMask = document.getElementById('btnExecuteMask');
    dom.btnReset = document.getElementById('btnReset');

    // Stats elements
    dom.statsGrid = document.getElementById('statsGrid');
    dom.statSheets = document.getElementById('statSheets');
    dom.statRows = document.getElementById('statRows');
    dom.statCells = document.getElementById('statCells');
    dom.statMasked = document.getElementById('statMasked');
    dom.statRatio = document.getElementById('statRatio');
    dom.statTime = document.getElementById('statTime');

    // Preview elements
    dom.previewCard = document.getElementById('previewCard');
    dom.sheetTabsBar = document.getElementById('sheetTabsBar');
    dom.tabMasked = document.getElementById('tabMasked');
    dom.tabOriginal = document.getElementById('tabOriginal');
    dom.previewSearchInput = document.getElementById('previewSearchInput');
    dom.tableScrollContainer = document.getElementById('tableScrollContainer');
    dom.tableHead = document.getElementById('tableHead');
    dom.tableBody = document.getElementById('tableBody');
    dom.paginationInfo = document.getElementById('paginationInfo');
    dom.btnPrevPage = document.getElementById('btnPrevPage');
    dom.btnNextPage = document.getElementById('btnNextPage');
    dom.selectPageSize = document.getElementById('selectPageSize');

    // Export elements
    dom.exportCard = document.getElementById('exportCard');
    dom.btnExportExcel = document.getElementById('btnExportExcel');
    dom.btnExportCsv = document.getElementById('btnExportCsv');
    dom.btnCopyClipboard = document.getElementById('btnCopyClipboard');
    dom.toast = document.getElementById('maskerToast');
  }

  // ===========================================================================
  // 1. THUẬT TOÁN CHE MỜ DỮ LIỆU (DATA MASKING LOGIC)
  // ===========================================================================
  /**
   * Che mờ giá trị của một ô dữ liệu
   * @param {any} value Giá trị ô
   * @param {boolean} isHeader Có phải ô header không
   * @returns {{ value: string, isMasked: boolean }}
   */
  function maskValue(value, isHeader = false) {
    // Nếu là ô Header và tùy chọn hasHeader = true: Giữ nguyên tuyệt đối
    if (isHeader && state.options.hasHeader) {
      return { value: value !== null && value !== undefined ? String(value) : '', isMasked: false };
    }

    if (value === null || value === undefined) {
      return { value: '', isMasked: false };
    }

    const str = String(value);

    // Chuỗi rỗng hoặc chỉ toàn khoảng trắng: giữ nguyên
    if (str.trim() === '') {
      return { value: str, isMasked: false };
    }

    const len = str.length;
    const starChar = state.options.maskChar || '*';

    // Điều kiện 1: Dưới 5 ký tự (< 5 ký tự) -> Che mờ toàn bộ bằng *****
    if (len < 5) {
      return { value: '*****', isMasked: true };
    }

    // Điều kiện 2: Đúng 5 ký tự
    if (len === 5) {
      if (state.options.fiveCharsRule === 'mask_all') {
        return { value: '*****', isMasked: true };
      } else {
        // Giữ 3 ký tự đầu + 2 sao
        return { value: str.slice(0, 3) + starChar.repeat(2), isMasked: true };
      }
    }

    // Điều kiện 3: Trên 5 ký tự (> 5 ký tự)
    // Giữ nguyên 3 ký tự đầu còn các ký tự sau sẽ che mờ hết đi ví dụ: abc****
    const prefix = str.slice(0, 3);
    let maskPart = '';

    if (state.options.starPaddingMode === 'fixed4') {
      maskPart = starChar.repeat(4);
    } else {
      // Số dấu sao bằng số ký tự còn lại (len - 3)
      const remainingCount = len - 3;
      maskPart = starChar.repeat(remainingCount);
    }

    return { value: prefix + maskPart, isMasked: true };
  }

  // ===========================================================================
  // 2. PLAYGROUND TESTER
  // ===========================================================================
  function updatePlayground() {
    if (!dom.playgroundInput || !dom.playgroundResult) return;
    const val = dom.playgroundInput.value;
    const res = maskValue(val, false);
    dom.playgroundResult.textContent = res.value || '(Trống)';
  }

  // ===========================================================================
  // 3. ĐỌC FILE (XLSX, XLS, CSV)
  // ===========================================================================
  function handleFileSelected(file) {
    if (!file) return;

    state.file = file;
    state.fileName = file.name;
    const ext = file.name.split('.').pop().toLowerCase();
    state.fileType = ext;

    dom.fileNameDisplay.textContent = file.name;
    dom.fileDetailsDisplay.textContent = `Dung lượng: ${(file.size / 1024).toFixed(1)} KB | Định dạng: .${ext.toUpperCase()}`;
    dom.fileInfoBar.style.display = 'flex';
    dom.dropzone.style.display = 'none';

    // Đọc file
    const reader = new FileReader();
    reader.onload = function (e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: false, raw: false });
        processLoadedWorkbook(workbook);
      } catch (err) {
        showToast('❌ Lỗi khi đọc file: ' + err.message);
        console.error(err);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function processLoadedWorkbook(workbook) {
    state.workbook = workbook;
    state.sheetNames = workbook.SheetNames || [];
    state.selectedSheets.clear();
    state.sheetData = {};

    // Mặc định chọn tất cả các sheet
    state.sheetNames.forEach(name => state.selectedSheets.add(name));
    state.activePreviewSheet = state.sheetNames[0] || '';

    // Render danh sách sheet cho người dùng chọn
    renderSheetSelection();

    // Tự động phân tích raw data ban đầu
    state.sheetNames.forEach(sheetName => {
      const sheet = workbook.Sheets[sheetName];
      const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      state.sheetData[sheetName] = {
        rawRows: rawData || [],
        maskedRows: [],
        headers: [],
        stats: {}
      };
    });

    // Kích hoạt nút thực thi
    dom.btnExecuteMask.disabled = false;
    showToast(`✅ Đã nạp thành công file (${state.sheetNames.length} sheet)`);

    // Tự động chạy che mờ luôn để người dùng thấy kết quả tức thì
    executeDataMasking();
  }

  function renderSheetSelection() {
    if (!dom.sheetChipsContainer) return;
    dom.sheetChipsContainer.innerHTML = '';

    if (state.sheetNames.length <= 1) {
      dom.sheetSelectionBox.style.display = 'none';
      return;
    }

    dom.sheetSelectionBox.style.display = 'block';

    state.sheetNames.forEach(sheetName => {
      const sheet = state.workbook.Sheets[sheetName];
      const range = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : null;
      const rowCount = range ? (range.e.r - range.s.r + 1) : 0;
      const isSelected = state.selectedSheets.has(sheetName);

      const chip = document.createElement('div');
      chip.className = `sheet-chip ${isSelected ? 'selected' : ''}`;
      chip.innerHTML = `
        <input type="checkbox" ${isSelected ? 'checked' : ''} style="cursor:pointer;" />
        <span>${escapeHtml(sheetName)}</span>
        <span style="font-size:10px; opacity:0.75;">(${rowCount} dòng)</span>
      `;

      chip.addEventListener('click', (e) => {
        if (e.target.tagName !== 'INPUT') {
          const chk = chip.querySelector('input');
          chk.checked = !chk.checked;
        }
        const isChecked = chip.querySelector('input').checked;
        if (isChecked) {
          state.selectedSheets.add(sheetName);
          chip.classList.add('selected');
        } else {
          state.selectedSheets.delete(sheetName);
          chip.classList.remove('selected');
        }
      });

      dom.sheetChipsContainer.appendChild(chip);
    });
  }

  // ===========================================================================
  // 4. TIẾN HÀNH CHE MỜ DỮ LIỆU TOÀN DIỆN (EXECUTE DATA MASKING)
  // ===========================================================================
  function executeDataMasking() {
    if (!state.workbook) {
      showToast('⚠️ Vui lòng tải file lên trước khi thực hiện!');
      return;
    }

    const startTime = performance.now();
    let totalRowsCount = 0;
    let totalCellsCount = 0;
    let totalMaskedCount = 0;
    let processedSheetsCount = 0;

    const hasHeader = state.options.hasHeader;
    const headerRowIdx = parseInt(state.options.headerRowIndex, 10) || 0;

    state.sheetNames.forEach(sheetName => {
      // Chỉ xử lý những sheet được chọn
      if (!state.selectedSheets.has(sheetName)) return;

      processedSheetsCount++;
      const currentSheetObj = state.sheetData[sheetName];
      const rawRows = currentSheetObj.rawRows;

      if (!rawRows || rawRows.length === 0) {
        currentSheetObj.maskedRows = [];
        return;
      }

      totalRowsCount += rawRows.length;

      const maskedRows = [];
      let sheetMaskedCells = 0;
      let sheetTotalCells = 0;

      // Tìm số cột lớn nhất
      let maxCols = 0;
      rawRows.forEach(row => {
        if (row.length > maxCols) maxCols = row.length;
      });

      // Duyệt từng dòng
      for (let r = 0; r < rawRows.length; r++) {
        const rawRow = rawRows[r] || [];
        const maskedRow = [];
        const isHeaderRow = (hasHeader && r === headerRowIdx);

        for (let c = 0; c < maxCols; c++) {
          const cellVal = rawRow[c] !== undefined ? rawRow[c] : '';
          sheetTotalCells++;

          const maskResult = maskValue(cellVal, isHeaderRow);
          maskedRow.push({
            value: maskResult.value,
            originalValue: cellVal,
            isMasked: maskResult.isMasked
          });

          if (maskResult.isMasked) {
            sheetMaskedCells++;
          }
        }

        maskedRows.push(maskedRow);
      }

      // Headers
      if (hasHeader && rawRows[headerRowIdx]) {
        currentSheetObj.headers = rawRows[headerRowIdx].map((h, idx) => (h !== '' && h !== null && h !== undefined) ? String(h) : `Cột ${getColumnLetter(idx)}`);
      } else {
        currentSheetObj.headers = Array.from({ length: maxCols }, (_, idx) => `Cột ${getColumnLetter(idx)}`);
      }

      currentSheetObj.maskedRows = maskedRows;
      currentSheetObj.stats = {
        totalRows: rawRows.length,
        totalCells: sheetTotalCells,
        maskedCells: sheetMaskedCells
      };

      totalCellsCount += sheetTotalCells;
      totalMaskedCount += sheetMaskedCells;
    });

    const endTime = performance.now();
    const execTime = Math.round(endTime - startTime);

    // Cập nhật State thống kê
    state.totalStats = {
      processedSheets: processedSheetsCount,
      totalRows: totalRowsCount,
      totalCells: totalCellsCount,
      maskedCells: totalMaskedCount,
      executionTimeMs: execTime
    };

    // Render giao diện kết quả
    updateStatsUI();
    renderSheetTabs();
    renderPreviewTable();

    // Hiển thị các khối giao diện
    dom.statsGrid.style.display = 'grid';
    dom.previewCard.style.display = 'block';
    dom.exportCard.style.display = 'flex';
    dom.btnExportExcel.disabled = false;
    dom.btnExportCsv.disabled = false;
    dom.btnCopyClipboard.disabled = false;

    showToast(`🎉 Đã che mờ xong ${totalMaskedCount.toLocaleString('vi-VN')} ô dữ liệu (${execTime}ms)!`);
  }

  // ===========================================================================
  // 5. HIỂN THỊ THỐNG KÊ (STATS UI)
  // ===========================================================================
  function updateStatsUI() {
    const s = state.totalStats;
    dom.statSheets.textContent = s.processedSheets;
    dom.statRows.textContent = s.totalRows.toLocaleString('vi-VN');
    dom.statCells.textContent = s.totalCells.toLocaleString('vi-VN');
    dom.statMasked.textContent = s.maskedCells.toLocaleString('vi-VN');

    const ratio = s.totalCells > 0 ? ((s.maskedCells / s.totalCells) * 100).toFixed(1) : 0;
    dom.statRatio.textContent = ratio + '%';
    dom.statTime.textContent = s.executionTimeMs + 'ms';
  }

  // ===========================================================================
  // 6. RENDER TABS SHEET VÀ BẢNG XEM TRƯỚC (PREVIEW TABLE)
  // ===========================================================================
  function renderSheetTabs() {
    if (!dom.sheetTabsBar) return;
    dom.sheetTabsBar.innerHTML = '';

    const validSheets = state.sheetNames.filter(name => state.selectedSheets.has(name));

    if (validSheets.length <= 1) {
      dom.sheetTabsBar.style.display = 'none';
      if (validSheets[0]) state.activePreviewSheet = validSheets[0];
      return;
    }

    dom.sheetTabsBar.style.display = 'flex';

    if (!state.selectedSheets.has(state.activePreviewSheet)) {
      state.activePreviewSheet = validSheets[0] || '';
    }

    validSheets.forEach(sheetName => {
      const tab = document.createElement('div');
      tab.className = `sheet-tab-item ${sheetName === state.activePreviewSheet ? 'active' : ''}`;
      tab.innerHTML = `<span>📑 ${escapeHtml(sheetName)}</span>`;
      tab.addEventListener('click', () => {
        state.activePreviewSheet = sheetName;
        state.preview.currentPage = 1;
        renderSheetTabs();
        renderPreviewTable();
      });
      dom.sheetTabsBar.appendChild(tab);
    });
  }

  function renderPreviewTable() {
    const sheetName = state.activePreviewSheet;
    const currentSheet = state.sheetData[sheetName];

    if (!currentSheet || !currentSheet.maskedRows || currentSheet.maskedRows.length === 0) {
      dom.tableHead.innerHTML = '';
      dom.tableBody.innerHTML = `<tr><td colspan="10" style="text-align:center; padding:30px; color:var(--text-muted);">Không có dữ liệu cho sheet này.</td></tr>`;
      dom.paginationInfo.textContent = 'Dòng 0 / 0';
      dom.btnPrevPage.disabled = true;
      dom.btnNextPage.disabled = true;
      return;
    }

    const headers = currentSheet.headers || [];
    const rows = currentSheet.maskedRows;
    const hasHeader = state.options.hasHeader;
    const headerRowIdx = parseInt(state.options.headerRowIndex, 10) || 0;
    const isMaskedView = state.preview.viewMode === 'masked';
    const searchTerm = (state.preview.searchTerm || '').toLowerCase();

    // Lọc theo tìm kiếm
    let filteredRows = [];
    rows.forEach((row, originalIndex) => {
      if (hasHeader && originalIndex === headerRowIdx) return; // Dòng header đã ở thead

      if (!searchTerm) {
        filteredRows.push({ row, originalIndex });
      } else {
        const matches = row.some(cell => {
          const val = isMaskedView ? cell.value : cell.originalValue;
          return String(val).toLowerCase().includes(searchTerm);
        });
        if (matches) {
          filteredRows.push({ row, originalIndex });
        }
      }
    });

    // Phân trang
    const totalFiltered = filteredRows.length;
    const pageSize = state.preview.pageSize;
    const totalPages = Math.ceil(totalFiltered / pageSize) || 1;

    if (state.preview.currentPage > totalPages) {
      state.preview.currentPage = totalPages;
    }
    const currentPage = state.preview.currentPage;
    const startIdx = (currentPage - 1) * pageSize;
    const endIdx = Math.min(startIdx + pageSize, totalFiltered);
    const pageRows = filteredRows.slice(startIdx, endIdx);

    // 1. Render Header (thead)
    let theadHtml = `<tr><th class="row-index-col">#</th>`;
    headers.forEach((h, colIdx) => {
      theadHtml += `<th title="Cột ${getColumnLetter(colIdx)}">${escapeHtml(h)}</th>`;
    });
    theadHtml += `</tr>`;
    dom.tableHead.innerHTML = theadHtml;

    // 2. Render Body (tbody)
    let tbodyHtml = '';
    pageRows.forEach((item, rIdx) => {
      const rowData = item.row;
      const actualRowNumber = item.originalIndex + 1;

      tbodyHtml += `<tr>`;
      tbodyHtml += `<td class="row-index-cell">${actualRowNumber}</td>`;

      rowData.forEach(cell => {
        const displayVal = isMaskedView ? cell.value : cell.originalValue;
        const isMaskedClass = (isMaskedView && cell.isMasked) ? 'cell-masked' : '';

        tbodyHtml += `<td class="${isMaskedClass}" title="${escapeHtml(String(displayVal))}">${escapeHtml(String(displayVal))}</td>`;
      });

      tbodyHtml += `</tr>`;
    });

    if (pageRows.length === 0) {
      tbodyHtml = `<tr><td colspan="${headers.length + 1}" style="text-align:center; padding:30px; color:var(--text-muted);">Không tìm thấy dòng nào khớp với từ khóa tìm kiếm.</td></tr>`;
    }

    dom.tableBody.innerHTML = tbodyHtml;

    // Cập nhật phân trang
    dom.paginationInfo.textContent = `Hiển thị ${totalFiltered === 0 ? 0 : startIdx + 1} - ${endIdx} của ${totalFiltered.toLocaleString('vi-VN')} dòng (Trang ${currentPage}/${totalPages})`;
    dom.btnPrevPage.disabled = (currentPage <= 1);
    dom.btnNextPage.disabled = (currentPage >= totalPages);
  }

  // ===========================================================================
  // 7. XUẤT FILE ĐÃ CHE MỜ (EXCEL, CSV & CLIPBOARD)
  // ===========================================================================
  function exportExcelFile() {
    if (!state.workbook) return;

    try {
      const newWb = XLSX.utils.book_new();
      let hasExportedSheet = false;

      state.sheetNames.forEach(sheetName => {
        if (!state.selectedSheets.has(sheetName)) return;

        const currentSheet = state.sheetData[sheetName];
        if (!currentSheet || !currentSheet.maskedRows) return;

        // Trích xuất mảng 2 chiều giá trị sau khi che mờ
        const aoa = currentSheet.maskedRows.map(row => {
          return row.map(cell => cell.value);
        });

        const newWs = XLSX.utils.aoa_to_sheet(aoa);
        XLSX.utils.book_append_sheet(newWb, newWs, sheetName);
        hasExportedSheet = true;
      });

      if (!hasExportedSheet) {
        showToast('⚠️ Không có sheet nào được chọn để xuất file!');
        return;
      }

      // Tạo tên file xuất
      const baseName = state.fileName.replace(/\.[^/.]+$/, '');
      const outName = `${baseName}_masked.xlsx`;

      XLSX.writeFile(newWb, outName);
      showToast(`💾 Đã xuất file thành công: ${outName}`);
    } catch (err) {
      showToast('❌ Lỗi khi xuất file Excel: ' + err.message);
      console.error(err);
    }
  }

  function exportCsvFile() {
    const sheetName = state.activePreviewSheet;
    const currentSheet = state.sheetData[sheetName];

    if (!currentSheet || !currentSheet.maskedRows || currentSheet.maskedRows.length === 0) {
      showToast('⚠️ Sheet hiện tại không có dữ liệu để xuất CSV!');
      return;
    }

    try {
      const aoa = currentSheet.maskedRows.map(row => {
        return row.map(cell => cell.value);
      });

      const ws = XLSX.utils.aoa_to_sheet(aoa);
      const csvContent = XLSX.utils.sheet_to_csv(ws);

      // Thêm BOM \uFEFF để Excel mở file CSV tiếng Việt không bị lỗi font
      const blob = new Blob(['\uFEFF' + csvContent], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');

      const baseName = state.fileName.replace(/\.[^/.]+$/, '');
      link.href = url;
      link.download = `${baseName}_${sheetName}_masked.csv`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);

      showToast(`💾 Đã xuất file CSV sheet "${sheetName}" thành công!`);
    } catch (err) {
      showToast('❌ Lỗi khi xuất CSV: ' + err.message);
      console.error(err);
    }
  }

  function copyToClipboard() {
    const sheetName = state.activePreviewSheet;
    const currentSheet = state.sheetData[sheetName];

    if (!currentSheet || !currentSheet.maskedRows || currentSheet.maskedRows.length === 0) {
      showToast('⚠️ Không có dữ liệu để sao chép!');
      return;
    }

    try {
      const tsvRows = currentSheet.maskedRows.map(row => {
        return row.map(cell => cell.value).join('\t');
      });
      const tsvText = tsvRows.join('\n');

      navigator.clipboard.writeText(tsvText).then(() => {
        showToast('📋 Đã sao chép toàn bộ bảng dữ liệu vào Clipboard (dán trực tiếp vào Excel)!');
      }).catch(err => {
        // Fallback textarea
        const ta = document.createElement('textarea');
        ta.value = tsvText;
        document.body.appendChild(ta);
        ta.select();
        document.execCommand('copy');
        document.body.removeChild(ta);
        showToast('📋 Đã sao chép vào Clipboard!');
      });
    } catch (err) {
      showToast('❌ Không thể sao chép: ' + err.message);
    }
  }

  // ===========================================================================
  // 8. RESET & TIỆN ÍCH HỖ TRỢ
  // ===========================================================================
  function resetAll() {
    state.file = null;
    state.fileName = '';
    state.workbook = null;
    state.sheetNames = [];
    state.selectedSheets.clear();
    state.sheetData = {};
    state.preview.currentPage = 1;
    state.preview.searchTerm = '';

    dom.fileInput.value = '';
    dom.fileInfoBar.style.display = 'none';
    dom.dropzone.style.display = 'block';
    dom.sheetSelectionBox.style.display = 'none';
    dom.statsGrid.style.display = 'none';
    dom.previewCard.style.display = 'none';
    dom.exportCard.style.display = 'none';
    dom.btnExecuteMask.disabled = true;

    showToast('🔄 Đã làm mới toàn bộ dữ liệu.');
  }

  function getColumnLetter(colIndex) {
    let temp, letter = '';
    let idx = colIndex;
    while (idx >= 0) {
      temp = idx % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      idx = Math.floor(idx / 26) - 1;
    }
    return letter;
  }

  function escapeHtml(str) {
    if (str === null || str === undefined) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function showToast(msg) {
    if (!dom.toast) return;
    dom.toast.textContent = msg;
    dom.toast.classList.add('show');
    clearTimeout(dom.toastTimer);
    dom.toastTimer = setTimeout(() => {
      dom.toast.classList.remove('show');
    }, 3500);
  }

  // ===========================================================================
  // 9. EVENT LISTENERS SETUP
  // ===========================================================================
  function setupEventListeners() {
    // 1. Dropzone & File Input
    dom.dropzone.addEventListener('click', () => dom.fileInput.click());
    dom.fileInput.addEventListener('change', (e) => {
      if (e.target.files && e.target.files[0]) {
        handleFileSelected(e.target.files[0]);
      }
    });

    ['dragenter', 'dragover'].forEach(eventName => {
      dom.dropzone.addEventListener(eventName, (e) => {
        e.preventDefault();
        dom.dropzone.classList.add('drag-over');
      });
    });

    ['dragleave', 'drop'].forEach(eventName => {
      dom.dropzone.addEventListener(eventName, (e) => {
        e.preventDefault();
        dom.dropzone.classList.remove('drag-over');
      });
    });

    dom.dropzone.addEventListener('drop', (e) => {
      if (e.dataTransfer.files && e.dataTransfer.files[0]) {
        handleFileSelected(e.dataTransfer.files[0]);
      }
    });

    dom.btnRemoveFile.addEventListener('click', resetAll);

    // 2. Options changes
    dom.chkHasHeader.addEventListener('change', (e) => {
      state.options.hasHeader = e.target.checked;
      if (state.workbook) executeDataMasking();
    });

    dom.selectHeaderRow.addEventListener('change', (e) => {
      state.options.headerRowIndex = parseInt(e.target.value, 10);
      if (state.workbook) executeDataMasking();
    });

    dom.radioFiveAll.addEventListener('change', () => {
      state.options.fiveCharsRule = 'mask_all';
      updatePlayground();
      if (state.workbook) executeDataMasking();
    });

    dom.radioFiveKeep3.addEventListener('change', () => {
      state.options.fiveCharsRule = 'keep_3';
      updatePlayground();
      if (state.workbook) executeDataMasking();
    });

    dom.radioStarExact.addEventListener('change', () => {
      state.options.starPaddingMode = 'exact';
      updatePlayground();
      if (state.workbook) executeDataMasking();
    });

    dom.radioStarFixed4.addEventListener('change', () => {
      state.options.starPaddingMode = 'fixed4';
      updatePlayground();
      if (state.workbook) executeDataMasking();
    });

    dom.btnSelectAllSheets.addEventListener('click', () => {
      const allSelected = state.selectedSheets.size === state.sheetNames.length;
      if (allSelected) {
        state.selectedSheets.clear();
      } else {
        state.sheetNames.forEach(n => state.selectedSheets.add(n));
      }
      renderSheetSelection();
      if (state.workbook) executeDataMasking();
    });

    // 3. Playground
    dom.playgroundInput.addEventListener('input', updatePlayground);
    dom.samplePills.forEach(pill => {
      pill.addEventListener('click', () => {
        dom.playgroundInput.value = pill.dataset.val || pill.textContent;
        updatePlayground();
      });
    });

    // 4. Action Buttons
    dom.btnExecuteMask.addEventListener('click', executeDataMasking);
    dom.btnReset.addEventListener('click', resetAll);

    // 5. Preview Tabs & Search
    dom.tabMasked.addEventListener('click', () => {
      state.preview.viewMode = 'masked';
      dom.tabMasked.classList.add('active');
      dom.tabOriginal.classList.remove('active');
      renderPreviewTable();
    });

    dom.tabOriginal.addEventListener('click', () => {
      state.preview.viewMode = 'original';
      dom.tabOriginal.classList.add('active');
      dom.tabMasked.classList.remove('active');
      renderPreviewTable();
    });

    dom.previewSearchInput.addEventListener('input', (e) => {
      state.preview.searchTerm = e.target.value.trim();
      state.preview.currentPage = 1;
      renderPreviewTable();
    });

    dom.btnPrevPage.addEventListener('click', () => {
      if (state.preview.currentPage > 1) {
        state.preview.currentPage--;
        renderPreviewTable();
      }
    });

    dom.btnNextPage.addEventListener('click', () => {
      state.preview.currentPage++;
      renderPreviewTable();
    });

    dom.selectPageSize.addEventListener('change', (e) => {
      state.preview.pageSize = parseInt(e.target.value, 10);
      state.preview.currentPage = 1;
      renderPreviewTable();
    });

    // 6. Export Buttons
    dom.btnExportExcel.addEventListener('click', exportExcelFile);
    dom.btnExportCsv.addEventListener('click', exportCsvFile);
    dom.btnCopyClipboard.addEventListener('click', copyToClipboard);
  }

  // Khởi động khi DOM sẵn sàng
  document.addEventListener('DOMContentLoaded', () => {
    initDOM();
    setupEventListeners();
    updatePlayground();
  });

})();
