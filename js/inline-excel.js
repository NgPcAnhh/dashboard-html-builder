/**
 * inline-excel.js - Bảng tính Mô phỏng Excel & Chế độ "Dùng data này"
 * Hỗ trợ:
 * - Scale động không giới hạn số cột (A, B, C... Z, AA, AB, AC... BA, BB...) và số dòng (1, 2, 3...)
 * - Dán (Paste) thông minh: Dán toàn bảng hoặc Dán bắt đầu từ ô đang chọn
 * - Điều hướng bàn phím như Excel (Tab, Enter, Phím mũi tên)
 * - Tự động đồng bộ vào rawRows & SheetJS Workbook
 */

const inlineFileNameInput = document.getElementById('inlineFileNameInput');
const inlineSheetNameInput = document.getElementById('inlineSheetNameInput');
const inlineStatsText = document.getElementById('inlineStatsText');
const btnInlinePaste = document.getElementById('btnInlinePaste');
const btnInlineAddRow = document.getElementById('btnInlineAddRow');
const btnInlineAddCol = document.getElementById('btnInlineAddCol');
const btnInlineDelCol = document.getElementById('btnInlineDelCol');
const btnToggleFirstRowHeader = document.getElementById('btnToggleFirstRowHeader');
const btnInlineClear = document.getElementById('btnInlineClear');
const inlineExcelViewport = document.getElementById('inlineExcelViewport');
const inlineExcelHead = document.getElementById('inlineExcelHead');
const inlineExcelBody = document.getElementById('inlineExcelBody');
const btnUseMockData = document.getElementById('btnUseMockData');

// Lưu trữ tọa độ ô đang được chọn { r: number, c: number }
let selectedCell = null;

// ========================================================
// 1. CHUYỂN CHỈ SỐ CỘT SANG CHỮ CÁI EXCEL
// 0 -> A, 1 -> B, ..., 25 -> Z, 26 -> AA, 27 -> AB... 51 -> AZ, 52 -> BA...
// ========================================================
function colIndexToExcelLetter(idx) {
  let letter = '';
  let n = idx + 1;
  while (n > 0) {
    let rem = (n - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter;
}

// Kiểm tra xem header hiện tại có phải là chuỗi mặc định A, B, C... hay không
function isDefaultAlphabetHeaders(headers) {
  if (!headers || !headers.length) return true;
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] !== colIndexToExcelLetter(i)) return false;
  }
  return true;
}

// ========================================================
// 2. RENDER BẢNG TÍNH MÔ PHỎNG EXCEL (INLINE VIEWPORT)
// ========================================================
function renderInlineExcel() {
  if (!inlineExcelHead || !inlineExcelBody) return;
  inlineExcelHead.innerHTML = '';
  inlineExcelBody.innerHTML = '';

  const { headers, rows } = inlineExcelState;
  const colCount = Math.max(headers.length, 1);

  // Chuẩn hóa mọi dòng dữ liệu có đủ colCount phần tử
  rows.forEach(r => {
    while (r.length < colCount) r.push('');
  });

  // --- THEAD: Hàng tiêu đề cột A, B, C... AA, AB... ---
  const trHead = document.createElement('tr');
  const thCorner = document.createElement('th');
  thCorner.className = 'inline-excel-corner-th';
  thCorner.textContent = '#';
  thCorner.title = 'Bảng tính Excel (Scale tự động theo dữ liệu)';
  trHead.appendChild(thCorner);

  for (let c = 0; c < colCount; c++) {
    const th = document.createElement('th');
    th.className = 'inline-excel-col-th';
    th.contentEditable = 'true';
    th.spellcheck = false;
    const colName = headers[c] !== undefined && headers[c] !== '' ? headers[c] : colIndexToExcelLetter(c);
    th.textContent = colName;
    th.dataset.c = c;
    th.title = `Cột ${colIndexToExcelLetter(c)}: Nhấp vào để sửa tên cột (Hiển thị trong code qua rawRows[i].${colName})`;

    th.oninput = () => {
      inlineExcelState.headers[c] = th.textContent.trim();
      updateInlineStats();
      onInlineDataChanged();
    };

    trHead.appendChild(th);
  }
  inlineExcelHead.appendChild(trHead);

  // --- TBODY: Dữ liệu từng dòng (1, 2, 3...) ---
  rows.forEach((row, r) => {
    const tr = document.createElement('tr');

    // Cột số thứ tự dòng 1, 2, 3...
    const tdRowIdx = document.createElement('td');
    tdRowIdx.className = 'inline-excel-row-td';
    tdRowIdx.innerHTML = `<span>${r + 1}</span> <button type="button" class="btn-inline-del-row" title="Xóa dòng ${r + 1}">&times;</button>`;

    const btnDel = tdRowIdx.querySelector('.btn-inline-del-row');
    btnDel.onclick = (e) => {
      e.stopPropagation();
      inlineExcelState.rows.splice(r, 1);
      if (inlineExcelState.rows.length === 0) {
        inlineExcelState.rows.push(new Array(colCount).fill(''));
      }
      renderInlineExcel();
      onInlineDataChanged();
    };
    tr.appendChild(tdRowIdx);

    // Các ô dữ liệu
    for (let c = 0; c < colCount; c++) {
      const td = document.createElement('td');
      td.contentEditable = 'true';
      td.spellcheck = false;
      const cellVal = row[c] !== undefined ? row[c] : '';
      td.textContent = cellVal;
      td.dataset.r = r;
      td.dataset.c = c;
      if (cellVal) td.title = cellVal;

      // Đánh dấu ô đang được chọn
      if (selectedCell && selectedCell.r === r && selectedCell.c === c) {
        td.classList.add('cell-selected');
      }

      // Xử lý Focus & Click
      td.onfocus = () => {
        highlightSelectedCell(r, c);
      };
      td.onclick = () => {
        highlightSelectedCell(r, c);
      };

      // Xử lý nhập liệu
      td.oninput = () => {
        if (!inlineExcelState.rows[r]) inlineExcelState.rows[r] = [];
        inlineExcelState.rows[r][c] = td.textContent;
        td.title = td.textContent;
        updateInlineStats();
        onInlineDataChanged();
      };

      // Bắt phím điều hướng (Tab, Enter, Mũi tên)
      td.onkeydown = (e) => {
        handleCellKeyNavigation(e, r, c, colCount, rows.length);
      };

      tr.appendChild(td);
    }

    inlineExcelBody.appendChild(tr);
  });

  updateInlineStats();
}

// Làm nổi bật ô đang chọn
function highlightSelectedCell(r, c) {
  selectedCell = { r, c };
  if (!inlineExcelBody) return;
  const prev = inlineExcelBody.querySelector('td.cell-selected');
  if (prev) prev.classList.remove('cell-selected');

  const curr = inlineExcelBody.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
  if (curr) curr.classList.add('cell-selected');
}

// Điều hướng bàn phím như Excel
function handleCellKeyNavigation(e, r, c, colCount, rowCount) {
  if (e.key === 'Tab') {
    e.preventDefault();
    if (e.shiftKey) {
      // Tab lùi
      if (c > 0) focusCell(r, c - 1);
      else if (r > 0) focusCell(r - 1, colCount - 1);
    } else {
      // Tab tới
      if (c < colCount - 1) focusCell(r, c + 1);
      else if (r < rowCount - 1) focusCell(r + 1, 0);
      else {
        addInlineRow();
        focusCell(rowCount, 0);
      }
    }
  } else if (e.key === 'Enter' && !e.shiftKey) {
    e.preventDefault();
    if (r < rowCount - 1) {
      focusCell(r + 1, c);
    } else {
      addInlineRow();
      focusCell(rowCount, c);
    }
  }
}

function focusCell(r, c) {
  if (!inlineExcelBody) return;
  const target = inlineExcelBody.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
  if (target) {
    target.focus();
    highlightSelectedCell(r, c);
  }
}

function updateInlineStats() {
  if (!inlineStatsText) return;
  const rCount = inlineExcelState.rows.length;
  const hCount = inlineExcelState.headers.length;
  inlineStatsText.textContent = `${rCount} dòng, ${hCount} cột (${colIndexToExcelLetter(0)} → ${colIndexToExcelLetter(Math.max(0, hCount - 1))})`;
}

// ========================================================
// 3. THAO TÁC THÊM / XÓA DÒNG & CỘT DỘNG
// ========================================================

// Thêm dòng mới (tự động tăng dần 1, 2, 3...)
function addInlineRow() {
  const colCount = Math.max(inlineExcelState.headers.length, 1);
  const newRow = new Array(colCount).fill('');
  inlineExcelState.rows.push(newRow);
  renderInlineExcel();
  onInlineDataChanged();

  // Tự động focus vào ô đầu tiên của dòng vừa tạo
  focusCell(inlineExcelState.rows.length - 1, 0);
}

// Thêm cột mới (tự động tăng dần A, B, C... Z, AA, AB, AC...)
function addInlineColumn() {
  const nextLetter = colIndexToExcelLetter(inlineExcelState.headers.length);
  inlineExcelState.headers.push(nextLetter);
  inlineExcelState.rows.forEach(r => r.push(''));
  renderInlineExcel();
  onInlineDataChanged();
  if (typeof printToTerminal === 'function') {
    printToTerminal(`[Bảng tính] Đã thêm cột mới: "${nextLetter}" (Tổng: ${inlineExcelState.headers.length} cột)`, "log-info");
  }
}

// Xóa cột cuối cùng
function removeInlineColumn() {
  if (inlineExcelState.headers.length <= 1) {
    alert("Bảng tính phải có ít nhất 1 cột!");
    return;
  }
  const removedName = inlineExcelState.headers.pop();
  inlineExcelState.rows.forEach(r => r.pop());
  renderInlineExcel();
  onInlineDataChanged();
  if (typeof printToTerminal === 'function') {
    printToTerminal(`[Bảng tính] Đã xóa cột: "${removedName}" (Còn lại: ${inlineExcelState.headers.length} cột)`, "log-warn");
  }
}

// Chuyển đổi qua lại giữa việc dùng dòng 1 làm Header hoặc đặt lại Header là A, B, C...
function toggleFirstRowHeader() {
  if (!inlineExcelState.rows.length) return;

  const currentIsAlphabet = isDefaultAlphabetHeaders(inlineExcelState.headers);

  if (currentIsAlphabet) {
    // Chuyển dòng 1 thành Header
    const firstRow = inlineExcelState.rows.shift();
    inlineExcelState.headers = firstRow.map((val, idx) => (val && val.trim()) ? val.trim() : colIndexToExcelLetter(idx));
    if (inlineExcelState.rows.length === 0) {
      inlineExcelState.rows.push(new Array(inlineExcelState.headers.length).fill(''));
    }
    if (typeof printToTerminal === 'function') {
      printToTerminal(`[Bảng tính] Đã chuyển dòng đầu tiên thành tiêu đề cột: [${inlineExcelState.headers.join(', ')}]`, "log-info");
    }
  } else {
    // Đưa Header xuống thành dòng 1 và đặt lại Header A, B, C...
    const oldHeadersAsRow = [...inlineExcelState.headers];
    inlineExcelState.rows.unshift(oldHeadersAsRow);
    inlineExcelState.headers = inlineExcelState.headers.map((_, idx) => colIndexToExcelLetter(idx));
    if (typeof printToTerminal === 'function') {
      printToTerminal(`[Bảng tính] Đã khôi phục tiêu đề cột chuẩn Excel: A, B, C... và chuyển tên cũ vào dòng 1.`, "log-info");
    }
  }

  renderInlineExcel();
  onInlineDataChanged();
}

// Xóa trắng bảng tính
function clearInlineTable() {
  if (confirm("Xóa trắng bảng tính mô phỏng để nhập lại từ đầu?")) {
    inlineExcelState.headers = ["A", "B", "C", "D"];
    inlineExcelState.rows = [["", "", "", ""]];
    selectedCell = null;
    renderInlineExcel();
    onInlineDataChanged();
    if (typeof printToTerminal === 'function') {
      printToTerminal(`[Bảng tính] Đã xóa trắng bảng tính về mặc định 4 cột (A, B, C, D).`, "log-warn");
    }
  }
}

if (btnInlineAddRow) btnInlineAddRow.onclick = addInlineRow;
if (btnInlineAddCol) btnInlineAddCol.onclick = addInlineColumn;
if (btnInlineDelCol) btnInlineDelCol.onclick = removeInlineColumn;
if (btnToggleFirstRowHeader) btnToggleFirstRowHeader.onclick = toggleFirstRowHeader;
if (btnInlineClear) btnInlineClear.onclick = clearInlineTable;

// ========================================================
// 4. PARSER CLIPBOARD THÔNG MINH (TSV TỪ EXCEL, CSV, MULTILINE)
// ========================================================
function parseClipboardGrid(rawText) {
  if (!rawText || !rawText.trim()) return [];

  // Nhận diện dấu phân cách chủ đạo
  let delim = '\t'; // Mặc định khi copy từ Excel/Google Sheets luôn là Tab
  const firstNewlineIdx = rawText.indexOf('\n');
  const sample = firstNewlineIdx !== -1 ? rawText.substring(0, firstNewlineIdx) : rawText;
  const tabCount = (sample.match(/\t/g) || []).length;
  const commaCount = (sample.match(/,/g) || []).length;
  const semiCount = (sample.match(/;/g) || []).length;

  if (tabCount > 0) {
    delim = '\t';
  } else if (semiCount > commaCount && semiCount > 0) {
    delim = ';';
  } else if (commaCount > 0) {
    delim = ',';
  }

  const lines = [];
  let curLine = [];
  let curVal = '';
  let inQuotes = false;

  for (let i = 0; i < rawText.length; i++) {
    const ch = rawText[i];
    const nextCh = rawText[i + 1];

    if (ch === '"') {
      if (inQuotes && nextCh === '"') {
        curVal += '"';
        i++; // Bỏ qua nháy kép escape
      } else {
        inQuotes = !inQuotes;
      }
    } else if (ch === delim && !inQuotes) {
      curLine.push(curVal.trim());
      curVal = '';
    } else if ((ch === '\r' || ch === '\n') && !inQuotes) {
      if (ch === '\r' && nextCh === '\n') i++; // Skip \r\n
      curLine.push(curVal.trim());
      curVal = '';
      if (curLine.some(c => c.length > 0)) {
        lines.push(curLine);
      }
      curLine = [];
    } else {
      curVal += ch;
    }
  }

  // Thu nạp phần tử cuối
  if (curVal.length > 0 || curLine.length > 0) {
    curLine.push(curVal.trim());
    if (curLine.some(c => c.length > 0)) {
      lines.push(curLine);
    }
  }

  return lines;
}

// ========================================================
// 5. CHỨC NĂNG DÁN THÔNG MINH & SCALE BẢNG TỰ ĐỘNG
// ========================================================
function pasteToInlineExcel(clipText, targetCell = null) {
  const grid = parseClipboardGrid(clipText);
  if (!grid.length) {
    alert("Clipboard hiện không có dữ liệu bảng tính hợp lệ!");
    return;
  }

  const pasteRows = grid.length;
  let pasteCols = 0;
  grid.forEach(row => {
    if (row.length > pasteCols) pasteCols = row.length;
  });
  pasteCols = Math.max(pasteCols, 1);

  // KỊCH BẢN 1: Dán bắt đầu từ một ô cụ thể mà người dùng đang chọn
  if (targetCell && typeof targetCell.r === 'number' && typeof targetCell.c === 'number') {
    const startR = targetCell.r;
    const startC = targetCell.c;

    const neededCols = startC + pasteCols;
    const neededRows = startR + pasteRows;

    // Tự động mở rộng số lượng cột nếu dữ liệu paste tràn ra ngoài bên phải
    while (inlineExcelState.headers.length < neededCols) {
      inlineExcelState.headers.push(colIndexToExcelLetter(inlineExcelState.headers.length));
    }

    // Tự động mở rộng số lượng dòng nếu dữ liệu paste tràn xuống dưới
    while (inlineExcelState.rows.length < neededRows) {
      inlineExcelState.rows.push(new Array(inlineExcelState.headers.length).fill(''));
    }

    // Đảm bảo tất cả các dòng hiện có đủ số cột
    inlineExcelState.rows.forEach(r => {
      while (r.length < inlineExcelState.headers.length) r.push('');
    });

    // Ghi đè dữ liệu dán vào đúng vị trí
    for (let i = 0; i < pasteRows; i++) {
      const r = startR + i;
      for (let j = 0; j < grid[i].length; j++) {
        const c = startC + j;
        inlineExcelState.rows[r][c] = grid[i][j];
      }
    }

    renderInlineExcel();
    onInlineDataChanged();
    highlightSelectedCell(startR, startC);

    const startCellName = `${colIndexToExcelLetter(startC)}${startR + 1}`;
    const endCellName = `${colIndexToExcelLetter(neededCols - 1)}${neededRows}`;
    if (typeof printToTerminal === 'function') {
      printToTerminal(`\n📋 [Dán Thông Minh] Đã dán thành công ${pasteRows} dòng x ${pasteCols} cột vào vùng [${startCellName} : ${endCellName}]!\n- Bảng tự động scale lên: ${inlineExcelState.rows.length} dòng x ${inlineExcelState.headers.length} cột.`, "log-info");
    }
    return;
  }

  // KỊCH BẢN 2: Dán toàn bảng (Full Table Paste) - Scale hoàn toàn theo kích thước dữ liệu
  // Tạo lại headers theo đúng số cột tối đa của dữ liệu dán vào: A, B, C, D... AA, AB...
  const newHeaders = [];
  for (let i = 0; i < pasteCols; i++) {
    newHeaders.push(colIndexToExcelLetter(i));
  }
  inlineExcelState.headers = newHeaders;

  // Chuẩn hóa mọi dòng có đúng pasteCols phần tử
  inlineExcelState.rows = grid.map(row => {
    const padded = [...row];
    while (padded.length < pasteCols) padded.push('');
    return padded;
  });

  renderInlineExcel();
  onInlineDataChanged();

  const colRange = `${colIndexToExcelLetter(0)} → ${colIndexToExcelLetter(pasteCols - 1)}`;
  if (typeof printToTerminal === 'function') {
    printToTerminal(`\n📋 [Dán Dữ Liệu Excel/CSV] ĐÃ TỰ ĐỘNG SCALE BẢNG THÀNH CÔNG!\n- Kích thước: ${pasteRows.toLocaleString('vi-VN')} dòng, ${pasteCols} cột (${colRange})\n- Tiêu đề cột: Chuẩn Excel tăng dần tự động (A, B, C... AA, AB...)\n- Bạn có thể nhấp vào tiêu đề để sửa tên cột hoặc bấm "🏷️ Dòng 1 ⇄ Header" nếu dòng 1 là tiêu đề.`, "log-info");
  }
}

// Bắt sự kiện Paste trên toàn bộ vùng bảng tính
if (inlineExcelViewport) {
  inlineExcelViewport.addEventListener('paste', (e) => {
    const clip = (e.clipboardData || window.clipboardData).getData('text');
    if (clip && clip.trim()) {
      e.preventDefault();

      // Kiểm tra xem người dùng có đang focus vào 1 ô dữ liệu td cụ thể không
      let target = null;
      const activeEl = document.activeElement;
      if (activeEl && activeEl.tagName === 'TD' && activeEl.dataset.r !== undefined) {
        target = {
          r: parseInt(activeEl.dataset.r, 10),
          c: parseInt(activeEl.dataset.c, 10)
        };
      } else if (selectedCell) {
        target = selectedCell;
      }

      pasteToInlineExcel(clip, target);
    }
  });
}

// Nút "📋 Dán từ Excel / CSV"
if (btnInlinePaste) {
  btnInlinePaste.onclick = async () => {
    try {
      const text = await navigator.clipboard.readText();
      if (text && text.trim()) {
        // Dán toàn bảng hoặc từ ô đang chọn
        pasteToInlineExcel(text, selectedCell);
      } else {
        alert("Clipboard của bạn hiện không có dữ liệu văn bản! Hãy copy bảng từ Excel trước rồi bấm nút này.");
      }
    } catch (err) {
      alert("Trình duyệt yêu cầu quyền truy cập Clipboard. Bạn hãy nhấp chuột vào ô cần dán trong bảng rồi nhấn phím tắt Ctrl + V!");
    }
  };
}

if (inlineFileNameInput) {
  inlineFileNameInput.addEventListener('input', () => {
    inlineExcelState.fileName = inlineFileNameInput.value.trim() || "mock_data.xlsx";
    onInlineDataChanged();
    if (typeof buildPrompt === 'function') buildPrompt();
  });
}

if (inlineSheetNameInput) {
  inlineSheetNameInput.addEventListener('input', () => {
    inlineExcelState.sheetName = inlineSheetNameInput.value.trim() || "Sheet1";
    onInlineDataChanged();
    if (typeof buildPrompt === 'function') buildPrompt();
  });
}

// ========================================================
// 6. ĐỒNG BỘ BẢNG MÔ PHỎNG VÀO DỮ LIỆU XỬ LÝ (rawRows, files...)
// ========================================================
function syncInlineExcelToDataState(showLog = true) {
  const { headers, rows } = inlineExcelState;
  let fileName = (inlineFileNameInput ? inlineFileNameInput.value.trim() : "") || inlineExcelState.fileName || "mock_data.xlsx";
  if (!fileName.toLowerCase().endsWith('.xlsx') && !fileName.toLowerCase().endsWith('.csv') && !fileName.toLowerCase().endsWith('.xls')) {
    fileName += ".xlsx";
  }
  const sheetName = (inlineSheetNameInput ? inlineSheetNameInput.value.trim() : "") || inlineExcelState.sheetName || "Sheet1";

  const validHeaders = headers.length ? headers : ["A"];
  const rowObjects = rows.map(row => {
    const obj = {};
    validHeaders.forEach((h, idx) => {
      obj[h] = row[idx] !== undefined ? row[idx] : "";
    });
    return obj;
  });

  // Tạo SheetJS Workbook để xử lý như file Excel thật
  const ws = XLSX.utils.json_to_sheet(rowObjects.length ? rowObjects : [{}]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);

  const existingIdx = loadedFiles.findIndex(f => f.name === fileName || f.isMockFile);
  const sheetsData = {};
  sheetsData[sheetName] = {
    name: sheetName,
    rows: rowObjects,
    columns: validHeaders,
    count: rowObjects.length,
    selected: true
  };

  const fileRowsArray = [...rowObjects];
  fileRowsArray.sheets = {};
  fileRowsArray[sheetName] = rowObjects;
  fileRowsArray.sheets[sheetName] = rowObjects;

  const fileItem = {
    id: existingIdx >= 0 ? loadedFiles[existingIdx].id : 'mock_file_' + Date.now(),
    name: fileName,
    sizeStr: (new Blob([JSON.stringify(rowObjects)]).size / 1024).toFixed(1) + " KB",
    workbook: wb,
    sheets: [sheetName],
    sheetsData: sheetsData,
    selectedSheets: [sheetName],
    currentSheet: sheetName,
    rows: fileRowsArray,
    columns: validHeaders,
    isMockFile: true
  };

  if (existingIdx >= 0) {
    loadedFiles[existingIdx] = fileItem;
  } else {
    loadedFiles.unshift(fileItem);
  }

  activeFileId = fileItem.id;
  if (typeof updateDataState === 'function') updateDataState();

  if (showLog && typeof printToTerminal === 'function') {
    printToTerminal(`\n⚡ [Dữ liệu mô phỏng] ĐÃ ĐỒNG BỘ DỮ LIỆU BẢNG TÍNH!\n- File: "${fileName}" (Sheet: "${sheetName}")\n- Dữ liệu: ${rowObjects.length.toLocaleString('vi-VN')} dòng x ${validHeaders.length} cột\n- Trạng thái: Code trỏ trực tiếp vào bảng này qua 'rawRows' và 'files["${fileName}"]'.\n👉 Bấm "▶ Chạy script" để bắt đầu xử lý!`, "log-info");
  }
}

function setUseMockDataState(enable, showLog = true) {
  isUsingMockData = enable;
  if (!btnUseMockData) return;

  if (isUsingMockData) {
    btnUseMockData.classList.add('active');
    btnUseMockData.innerHTML = '✅ Đang dùng data này';
    btnUseMockData.title = 'Đang dùng dữ liệu từ bảng mô phỏng. Bấm lại để tắt.';
    syncInlineExcelToDataState(showLog);
  } else {
    btnUseMockData.classList.remove('active');
    btnUseMockData.innerHTML = '⚡ Dùng data này';
    btnUseMockData.title = 'Bật tính năng này để code tự động trỏ vào bảng tính mô phỏng và xử lý như file Excel thông thường';
    // Loại bỏ mock file khỏi loadedFiles nếu người dùng tắt
    const mockIdx = loadedFiles.findIndex(f => f.isMockFile);
    if (mockIdx >= 0) {
      loadedFiles.splice(mockIdx, 1);
      if (loadedFiles.length > 0) {
        activeFileId = loadedFiles[0].id;
      } else {
        activeFileId = null;
      }
      if (typeof updateDataState === 'function') updateDataState();
    }
    if (showLog && typeof printToTerminal === 'function') {
      printToTerminal(`\n[Mô phỏng Excel] Đã tắt chế độ dùng data mô phỏng.`, "log-warn");
    }
  }
  if (typeof buildPrompt === 'function') buildPrompt();
}

if (btnUseMockData) {
  btnUseMockData.onclick = () => {
    setUseMockDataState(!isUsingMockData, true);
  };
}
