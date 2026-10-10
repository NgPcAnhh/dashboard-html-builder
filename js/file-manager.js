/**
 * file-manager.js - Quản lý Tải và Xử lý Đa File Excel / CSV & Đa Sheet Toàn Diện
 * Tính năng:
 * - Mặc định nạp và chọn TẤT CẢ các sheet từ mỗi file Excel (không chỉ 1 sheet).
 * - Cho phép chọn/bỏ chọn từng sheet hoặc chọn tất cả linh hoạt theo nhu cầu.
 * - Hỗ trợ nạp 2 hoặc nhiều file Excel/CSV đồng thời, code có thể gọi bất kỳ file nào qua files['tên_file'] hoặc getFile().
 * - Cung cấp hàm getSheet('tên_file', 'tên_sheet') hoặc getSheet('tên_sheet') tiện lợi.
 */

const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const badgeLoaded = document.getElementById('badgeLoaded');
const filesListWrapper = document.getElementById('filesListWrapper');
const filesList = document.getElementById('filesList');
const totalStats = document.getElementById('totalStats');
const btnClearFiles = document.getElementById('btnClearFiles');

// Chặn hành vi mở file mặc định của trình duyệt
window.addEventListener('dragover', (e) => e.preventDefault());
window.addEventListener('drop', (e) => e.preventDefault());

if (dropZone && fileInput) {
  dropZone.addEventListener('click', () => {
    fileInput.value = '';
    fileInput.click();
  });

  ['dragenter', 'dragover'].forEach(name => {
    dropZone.addEventListener(name, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.add('dragover');
    });
  });

  ['dragleave', 'drop'].forEach(name => {
    dropZone.addEventListener(name, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.remove('dragover');
    });
  });

  dropZone.addEventListener('drop', (e) => {
    const droppedFiles = Array.from(e.dataTransfer.files);
    if (droppedFiles.length > 0) {
      handleFilesSelected(droppedFiles);
    }
  });

  fileInput.addEventListener('change', (e) => {
    const selectedFiles = Array.from(e.target.files);
    if (selectedFiles.length > 0) {
      handleFilesSelected(selectedFiles);
    }
  });
}

async function handleFilesSelected(fileArr) {
  const validFiles = fileArr.filter(f => {
    const name = f.name.toLowerCase();
    return name.endsWith('.xlsx') || name.endsWith('.xls') || name.endsWith('.csv');
  });

  if (validFiles.length === 0) {
    alert("Vui lòng chỉ chọn các file bảng tính có định dạng .xlsx, .xls hoặc .csv!");
    return;
  }

  printToTerminal(`\n⏳ Đang nạp và đọc ${validFiles.length} file (tự động nạp 100% các sheet)... Vui lòng đợi.`, "log-info");

  for (const f of validFiles) {
    await readSpreadsheetFile(f);
  }

  updateDataState();
}

/**
 * Đọc file bảng tính và nạp TẤT CẢ các sheet vào bộ nhớ
 */
function readSpreadsheetFile(file) {
  return new Promise((resolve) => {
    const fileName = file.name;
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetNames = workbook.SheetNames || [];

        const sheetsData = {};
        const allColumnsSet = new Set();

        // 1. Phân tích TẤT CẢ các sheet từ file Excel đó
        sheetNames.forEach(shName => {
          const worksheet = workbook.Sheets[shName];
          const parsedRows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
          const cols = parsedRows.length > 0 ? Object.keys(parsedRows[0]) : [];
          cols.forEach(c => allColumnsSet.add(c));

          sheetsData[shName] = {
            name: shName,
            rows: parsedRows,
            columns: cols,
            count: parsedRows.length,
            selected: true // MẶC ĐỊNH CHỌN TẤT CẢ CÁC SHEET
          };
        });

        const existingIdx = loadedFiles.findIndex(f => f.name === fileName);
        const fileItem = {
          id: existingIdx >= 0 ? loadedFiles[existingIdx].id : Date.now() + Math.floor(Math.random() * 1000),
          name: fileName,
          size: file.size,
          sizeStr: (file.size / 1024).toFixed(1) + " KB",
          workbook: workbook,
          sheets: sheetNames,
          sheetsData: sheetsData,
          selectedSheets: [...sheetNames], // Mặc định tất cả sheet đều được chọn
          rows: [],
          columns: Array.from(allColumnsSet)
        };

        // Tính toán rows tổng hợp từ các sheet đã chọn
        recomputeFileRows(fileItem);

        if (existingIdx >= 0) {
          loadedFiles[existingIdx] = fileItem;
        } else {
          loadedFiles.push(fileItem);
        }

        if (!activeFileId) {
          activeFileId = fileItem.id;
        }

        const msg = `✓ Đã nạp thành công "${fileName}" [${sheetNames.length} sheet: ${sheetNames.join(', ')}] - Tổng ${fileItem.rows.length.toLocaleString('vi-VN')} dòng`;
        console.log(msg);
        printToTerminal(msg, "log-info");
      } catch (err) {
        const errMsg = `[Lỗi nạp ${fileName}]: ${err.message}`;
        console.error(errMsg);
        printToTerminal(errMsg, "log-error");
      }
      resolve();
    };

    reader.onerror = () => {
      console.error(`[Lỗi đọc file]: Không thể đọc ${fileName}`);
      resolve();
    };

    reader.readAsArrayBuffer(file);
  });
}

/**
 * Tái tổng hợp dữ liệu mảng rows của file dựa trên các sheet đang được chọn
 */
function recomputeFileRows(fileItem) {
  if (!fileItem || !fileItem.sheetsData) return;

  const selectedSheets = fileItem.sheets.filter(sh => fileItem.sheetsData[sh] && fileItem.sheetsData[sh].selected);
  fileItem.selectedSheets = selectedSheets;

  let combined = [];
  selectedSheets.forEach(sh => {
    const rawList = fileItem.sheetsData[sh].rows || [];
    // Gắn thông tin _sheet và _file để code người dùng phân biệt nguồn gốc dòng nếu cần
    const tagged = rawList.map(r => ({ ...r, _sheet: sh, _file: fileItem.name }));
    combined = combined.concat(tagged);
  });

  // Tạo mảng vừa là mảng các dòng gộp, vừa cho phép truy cập trực tiếp theo tên sheet:
  // fileItem.rows['Sheet1'] = rows
  const fileRowsArray = [...combined];
  fileRowsArray.sheets = {};
  fileItem.sheets.forEach(sh => {
    const shRows = fileItem.sheetsData[sh] ? fileItem.sheetsData[sh].rows : [];
    fileRowsArray[sh] = shRows;
    fileRowsArray.sheets[sh] = shRows;
  });

  fileItem.rows = fileRowsArray;
}

/**
 * Bật / tắt chọn một sheet cụ thể trong file
 */
function toggleSheetSelection(fileId, sheetName, isSelected) {
  const file = loadedFiles.find(f => f.id === fileId);
  if (!file || !file.sheetsData || !file.sheetsData[sheetName]) return;

  file.sheetsData[sheetName].selected = isSelected;
  recomputeFileRows(file);
  updateDataState();

  const count = file.sheetsData[sheetName].count;
  const statusStr = isSelected ? "Đã bật" : "Đã tắt";
  printToTerminal(`[Sheet] ${statusStr} sheet "${sheetName}" của "${file.name}" (${count.toLocaleString('vi-VN')} dòng)`, "log-info");
}

/**
 * Chọn tất cả hoặc bỏ chọn tất cả các sheet của một file
 */
function toggleAllSheets(fileId, selectAll) {
  const file = loadedFiles.find(f => f.id === fileId);
  if (!file || !file.sheetsData) return;

  file.sheets.forEach(sh => {
    if (file.sheetsData[sh]) {
      file.sheetsData[sh].selected = selectAll;
    }
  });

  recomputeFileRows(file);
  updateDataState();

  const statusStr = selectAll ? "Chọn tất cả" : "Bỏ chọn tất cả";
  printToTerminal(`[Sheet] ${statusStr} (${file.sheets.length} sheet) trong file "${file.name}"`, "log-info");
}

/**
 * Cập nhật toàn bộ trạng thái dữ liệu (files, fileSheets, rawRows)
 */
function updateDataState() {
  files = {};
  fileSheets = {};
  let totalRows = 0;
  let totalSheets = 0;

  loadedFiles.forEach(f => {
    files[f.name] = f.rows;
    fileSheets[f.name] = {};

    if (f.sheetsData) {
      f.sheets.forEach(sh => {
        fileSheets[f.name][sh] = f.sheetsData[sh].rows;
      });
      totalSheets += (f.selectedSheets ? f.selectedSheets.length : 0);
    } else {
      const sh = f.currentSheet || (f.sheets && f.sheets[0]) || "Sheet1";
      fileSheets[f.name][sh] = f.rows;
      if (f.rows) {
        f.rows[sh] = f.rows;
        if (!f.rows.sheets) f.rows.sheets = {};
        f.rows.sheets[sh] = f.rows;
      }
      totalSheets += 1;
    }
    totalRows += (f.rows ? f.rows.length : 0);
  });

  const activeFile = loadedFiles.find(f => f.id === activeFileId) || loadedFiles[0];
  if (activeFile) {
    activeFileId = activeFile.id;
    rawRows = activeFile.rows;
    dataColumns = activeFile.columns;
    sampleRecords = rawRows.slice(0, 2);
  } else {
    activeFileId = null;
    rawRows = [];
    dataColumns = [];
    sampleRecords = [];
  }

  renderFilesList();

  if (loadedFiles.length > 0) {
    if (badgeLoaded) {
      badgeLoaded.style.display = 'inline-block';
      badgeLoaded.textContent = `${loadedFiles.length} file • ${totalSheets} sheets (${totalRows.toLocaleString('vi-VN')} dòng)`;
    }
    if (btnClearFiles) btnClearFiles.style.display = 'inline-block';
    if (filesListWrapper) filesListWrapper.style.display = 'flex';
    if (totalStats) totalStats.textContent = `Tổng: ${loadedFiles.length} file, ${totalSheets} sheet đã chọn, ${totalRows.toLocaleString('vi-VN')} dòng`;
  } else {
    if (badgeLoaded) badgeLoaded.style.display = 'none';
    if (btnClearFiles) btnClearFiles.style.display = 'none';
    if (filesListWrapper) filesListWrapper.style.display = 'none';
  }

  if (typeof buildPrompt === 'function') {
    buildPrompt();
  }
}

/**
 * Render danh sách file và danh sách sheet trực quan
 */
function renderFilesList() {
  if (!filesList) return;
  filesList.innerHTML = '';

  loadedFiles.forEach((f, idx) => {
    const card = document.createElement('div');
    card.className = `file-card ${f.id === activeFileId ? 'active' : ''}`;

    // 1. Dòng Header: Số thứ tự file, Tên file, Nút copy mã gọi, Nút xóa
    const headerRow = document.createElement('div');
    headerRow.className = 'file-card-header';

    const titleGroup = document.createElement('div');
    titleGroup.className = 'file-card-title-group';

    const badgeIdx = document.createElement('span');
    badgeIdx.className = 'file-badge-index';
    badgeIdx.textContent = `File ${idx + 1}`;

    const nameSpan = document.createElement('span');
    nameSpan.className = 'file-card-name';
    nameSpan.textContent = f.name;
    nameSpan.title = f.name;

    const copyCodeChip = document.createElement('button');
    copyCodeChip.type = 'button';
    copyCodeChip.className = 'file-code-chip';
    copyCodeChip.textContent = `files['${f.name}'] 📋`;
    copyCodeChip.title = `Nhấp để copy mã gọi file này vào clipboard`;
    copyCodeChip.onclick = (e) => {
      e.stopPropagation();
      navigator.clipboard.writeText(`files['${f.name}']`).then(() => {
        copyCodeChip.textContent = `✓ Đã copy!`;
        setTimeout(() => {
          copyCodeChip.textContent = `files['${f.name}'] 📋`;
        }, 1500);
      });
    };

    titleGroup.appendChild(badgeIdx);
    titleGroup.appendChild(nameSpan);
    titleGroup.appendChild(copyCodeChip);

    const metaDiv = document.createElement('div');
    metaDiv.className = 'file-card-meta';
    const selCount = f.selectedSheets ? f.selectedSheets.length : 0;
    const totalCount = f.sheets ? f.sheets.length : 1;
    metaDiv.innerHTML = `<span><b>${selCount}/${totalCount}</b> sheet đã chọn</span> • <span>${f.rows.length.toLocaleString('vi-VN')} dòng</span> • <span>${f.sizeStr || ''}</span>`;

    const removeBtn = document.createElement('button');
    removeBtn.className = 'file-remove-btn';
    removeBtn.innerHTML = '&times;';
    removeBtn.title = 'Xóa file này';
    removeBtn.onclick = (e) => {
      e.stopPropagation();
      removeFile(f.id);
    };

    headerRow.appendChild(titleGroup);
    headerRow.appendChild(metaDiv);
    headerRow.appendChild(removeBtn);
    card.appendChild(headerRow);

    // 2. Khu vực danh sách các Sheet (Mặc định chọn tất cả, người dùng có thể tùy ý chọn)
    if (f.sheets && f.sheets.length > 0) {
      const sheetSec = document.createElement('div');
      sheetSec.className = 'sheet-section';

      const secHeader = document.createElement('div');
      secHeader.className = 'sheet-section-header';
      secHeader.innerHTML = `<span>📑 <b>Danh Sách Sheet:</b> (Mặc định chọn tất cả, code truy cập trực tiếp từng sheet)</span>`;

      const quickActions = document.createElement('div');
      quickActions.className = 'sheet-quick-actions';

      const btnSelectAll = document.createElement('button');
      btnSelectAll.type = 'button';
      btnSelectAll.className = 'btn-sheet-quick';
      btnSelectAll.textContent = '✓ Chọn tất cả';
      btnSelectAll.title = 'Chọn tất cả các sheet của file này';
      btnSelectAll.onclick = (e) => {
        e.stopPropagation();
        toggleAllSheets(f.id, true);
      };

      const btnDeselectAll = document.createElement('button');
      btnDeselectAll.type = 'button';
      btnDeselectAll.className = 'btn-sheet-quick';
      btnDeselectAll.textContent = 'Bỏ chọn';
      btnDeselectAll.title = 'Bỏ chọn tất cả các sheet của file này';
      btnDeselectAll.onclick = (e) => {
        e.stopPropagation();
        toggleAllSheets(f.id, false);
      };

      quickActions.appendChild(btnSelectAll);
      quickActions.appendChild(btnDeselectAll);
      secHeader.appendChild(quickActions);
      sheetSec.appendChild(secHeader);

      const chipsList = document.createElement('div');
      chipsList.className = 'sheet-chips-list';

      f.sheets.forEach(shName => {
        const shData = f.sheetsData[shName];
        if (!shData) return;

        const isSelected = !!shData.selected;
        const chip = document.createElement('label');
        chip.className = `sheet-chip ${isSelected ? 'selected' : ''}`;
        chip.title = `Tích chọn để nạp sheet "${shName}" (${shData.count} dòng). Code gọi: files['${f.name}']['${shName}']`;

        const chk = document.createElement('input');
        chk.type = 'checkbox';
        chk.checked = isSelected;
        chk.onclick = (e) => {
          e.stopPropagation();
        };
        chk.onchange = (e) => {
          e.stopPropagation();
          toggleSheetSelection(f.id, shName, e.target.checked);
        };

        const shTitle = document.createElement('span');
        shTitle.textContent = shName;

        const shCount = document.createElement('span');
        shCount.className = 'sheet-chip-count';
        shCount.textContent = `(${shData.count.toLocaleString('vi-VN')})`;

        chip.appendChild(chk);
        chip.appendChild(shTitle);
        chip.appendChild(shCount);
        chipsList.appendChild(chip);
      });

      sheetSec.appendChild(chipsList);
      card.appendChild(sheetSec);
    }

    // Nhấp vào card để kích hoạt làm active file (nếu muốn dùng rawRows trỏ tới nó)
    card.onclick = () => {
      activeFileId = f.id;
      updateDataState();
      printToTerminal(`\n[Active File] rawRows hiện trỏ tới: "${f.name}" (${f.rows.length.toLocaleString('vi-VN')} dòng)`, "log-info");
    };

    filesList.appendChild(card);
  });
}

function removeFile(fileId) {
  loadedFiles = loadedFiles.filter(f => f.id !== fileId);
  if (activeFileId === fileId) {
    activeFileId = loadedFiles.length > 0 ? loadedFiles[0].id : null;
  }
  updateDataState();
  console.log("[Hệ thống] Đã xóa file khỏi danh sách.");
  printToTerminal("[Hệ thống] Đã xóa file khỏi danh sách.", "log-info");
}

if (btnClearFiles) {
  btnClearFiles.onclick = () => {
    if (confirm("Bạn có chắc chắn muốn xóa toàn bộ danh sách file đã nạp?")) {
      loadedFiles = [];
      activeFileId = null;
      updateDataState();
      console.log("[Hệ thống] Đã xóa tất cả các file.");
      printToTerminal("[Hệ thống] Đã xóa tất cả các file.", "log-info");
    }
  };
}

// ========================================================
// CÁC HÀM TIỆN ÍCH TRUY CẬP ĐA FILE & ĐA SHEET TOÀN CỤC CHO CODE
// ========================================================

/**
 * Lấy dữ liệu file theo chỉ số (0, 1, 2...) hoặc theo tên / từ khóa linh hoạt
 */
function getFile(query) {
  if (typeof query === 'number') {
    const f = loadedFiles[query];
    return f ? f.rows : null;
  }
  if (!query || typeof query !== 'string') return null;

  // Khớp chính xác
  if (files[query]) return files[query];

  // Khớp linh hoạt (bỏ đuôi .xlsx/.xls/.csv, không phân biệt hoa thường)
  const qClean = query.toLowerCase().trim().replace(/\.(xlsx|xls|csv)$/, '');
  const found = loadedFiles.find(f => {
    const fClean = f.name.toLowerCase().trim().replace(/\.(xlsx|xls|csv)$/, '');
    return f.name.toLowerCase() === query.toLowerCase() || fClean === qClean || fClean.includes(qClean);
  });

  return found ? found.rows : null;
}

/**
 * Lấy dữ liệu của một sheet cụ thể:
 * - getSheet('file.xlsx', 'Sheet1'): Lấy Sheet1 của file cụ thể
 * - getSheet('Sheet1'): Tìm Sheet1 trên bất kỳ file nào có chứa nó
 */
function getSheet(arg1, arg2) {
  if (arg2) {
    const file = getFile(arg1);
    if (!file) return null;
    if (file[arg2]) return file[arg2];
    if (file.sheets && file.sheets[arg2]) return file.sheets[arg2];
    return null;
  }

  // Nếu chỉ truyền 1 tham số (tên sheet)
  for (const f of loadedFiles) {
    if (f.sheetsData && f.sheetsData[arg1]) {
      return f.sheetsData[arg1].rows;
    }
  }
  return null;
}

window.getFile = getFile;
window.getSheet = getSheet;

/**
 * Hàm tiện ích tự động gộp tất cả các dòng từ tất cả các file đã chọn
 * @param {Array<string>} [fileNames] Tùy chọn lọc danh sách file cần gộp
 */
function mergeAllRows(fileNames) {
  let combined = [];
  const targetFiles = Array.isArray(fileNames) && fileNames.length > 0
    ? loadedFiles.filter(f => fileNames.includes(f.name))
    : loadedFiles;

  targetFiles.forEach(f => {
    const tagged = f.rows.map(r => ({ ...r, _sourceFile: f.name }));
    combined = combined.concat(tagged);
  });
  return combined;
}
window.mergeAllRows = mergeAllRows;

/**
 * Hàm xuất file Excel toàn cục
 */
window.exportExcel = function (dataArray, fileName = "output.xlsx") {
  if (!dataArray || !dataArray.length) {
    console.error("Không có dữ liệu mảng để xuất file!");
    printToTerminal("Lỗi xuất file: Mảng dữ liệu rỗng!", "log-error");
    return;
  }
  const ws = XLSX.utils.json_to_sheet(dataArray);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "KetQua");
  XLSX.writeFile(wb, fileName);
  console.log(`[Hệ thống] Đã xuất và tải xuống file: ${fileName}`);
  printToTerminal(`✓ Đã xuất và tải xuống file: ${fileName} (${dataArray.length.toLocaleString('vi-VN')} dòng)`, "log-success");
};
