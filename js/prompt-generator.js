/**
 * prompt-generator.js - Trình Tạo Prompt AI Cho Xử Lý Dữ Liệu
 * Hỗ trợ chọn linh hoạt nguồn truyền dữ liệu mẫu (File tải lên, Bảng mô phỏng Excel),
 * hỗ trợ đa file, và chỉ hiển thị dữ liệu mẫu khi người dùng kích hoạt.
 */

const promptModal = document.getElementById('promptModal');
const inputDataDesc = document.getElementById('inputDataDesc');
const businessLogicDesc = document.getElementById('businessLogicDesc');
const outputDesc = document.getElementById('outputDesc');
const fullPromptOutput = document.getElementById('fullPromptOutput');
const btnOpenPrompt = document.getElementById('btnOpenPrompt');
const btnCloseModal = document.getElementById('btnCloseModal');
const btnCloseModalX = document.getElementById('btnCloseModalX');
const btnCopyPrompt = document.getElementById('btnCopyPrompt');
const toggleMockPreview = document.getElementById('toggleMockPreview');
const mockPreviewContent = document.getElementById('mockPreviewContent');
const mockPreviewBadge = document.getElementById('mockPreviewBadge');
const chkIncludeSampleData = document.getElementById('chkIncludeSampleData');
const promptSourcePicker = document.getElementById('promptSourcePicker');
const promptSourcesList = document.getElementById('promptSourcesList');
const btnSelectAllPromptSources = document.getElementById('btnSelectAllPromptSources');
const btnUnselectAllPromptSources = document.getElementById('btnUnselectAllPromptSources');
const previewFilesCount = document.getElementById('previewFilesCount');

// Tập hợp các ID nguồn file được chọn để truyền dữ liệu mẫu vào Prompt AI
let selectedPromptSourceIds = new Set();

// ========================================================
// TIỆN ÍCH CHE MỜ 75% DỮ LIỆU & FORMAT BẢNG CHO PROMPT
// ========================================================

// Hàm che mờ 75% dữ liệu của từng ô để bảo vệ thông tin khi cung cấp mẫu cho AI
function mask75PercentValue(val) {
  if (val === null || val === undefined) return '';
  const str = String(val).trim();
  if (str === '') return '';
  const len = str.length;
  if (len === 1) return '*';
  if (len === 2) return str[0] + '*';
  if (len === 3) return str[0] + '**';
  const visibleLen = Math.max(1, Math.round(len * 0.25));
  const maskLen = len - visibleLen;
  return str.substring(0, visibleLen) + '*'.repeat(maskLen);
}

// Chuyển mảng object sang bảng Markdown thân thiện với AI
function formatMarkdownTable(cols, rows) {
  if (!cols || cols.length === 0 || !rows || rows.length === 0) return '';
  const head = '| ' + cols.join(' | ') + ' |';
  const sep = '| ' + cols.map(() => '---').join(' | ') + ' |';
  const body = rows.map(r => {
    return '| ' + cols.map(c => {
      let cell = r[c] !== undefined ? String(r[c]) : '';
      return cell.replace(/\|/g, '\\|').replace(/\r?\n/g, ' ');
    }).join(' | ') + ' |';
  }).join('\n');
  return `${head}\n${sep}\n${body}`;
}

// ========================================================
// THU THẬP TẤT CẢ NGUỒN FILE KHẢ DỤNG (MÔ PHỎNG + FILE TẢI LÊN)
// ========================================================

function getAvailablePromptSources(shouldMask75 = true) {
  const sources = [];

  // 1. Nguồn Bảng mô phỏng Excel
  const mockFileName = (inlineFileNameInput ? inlineFileNameInput.value.trim() : "") || (inlineExcelState && inlineExcelState.fileName) || "mock_data.xlsx";
  const mockSheetName = (inlineSheetNameInput ? inlineSheetNameInput.value.trim() : "") || (inlineExcelState && inlineExcelState.sheetName) || "Sheet1";
  const mockCols = (inlineExcelState && inlineExcelState.headers && inlineExcelState.headers.length > 0)
    ? [...inlineExcelState.headers]
    : ["A", "B", "C", "D"];
  const mockRowsRaw = (inlineExcelState && inlineExcelState.rows) ? inlineExcelState.rows : [];
  
  // Lọc lấy các dòng có ít nhất một ô có dữ liệu
  const nonEmptyMockRows = mockRowsRaw.filter(r => r && r.some(c => c !== null && c !== undefined && String(c).trim() !== ""));
  const slice5Mock = (nonEmptyMockRows.length > 0 ? nonEmptyMockRows : mockRowsRaw).slice(0, 5);
  const maskedMockRows = slice5Mock.map(r => {
    const rowObj = {};
    mockCols.forEach((c, idx) => {
      const val = r ? r[idx] : '';
      rowObj[c] = shouldMask75 ? mask75PercentValue(val) : (val !== undefined ? val : '');
    });
    return rowObj;
  });

  sources.push({
    id: 'source_mock_excel',
    name: mockFileName,
    sheet: mockSheetName,
    sheets: [mockSheetName],
    totalRows: nonEmptyMockRows.length,
    columns: mockCols,
    rows: maskedMockRows,
    type: 'mock',
    badge: '⚡ Bảng mô phỏng Excel'
  });

  // 2. Nguồn Các file tải lên từ máy tính (loại bỏ file mock nếu nó đã sync vào loadedFiles để tránh trùng lặp)
  if (typeof loadedFiles !== 'undefined' && loadedFiles && loadedFiles.length > 0) {
    loadedFiles.forEach(f => {
      if (f.isMockFile) return; // Đã đại diện bởi source_mock_excel
      const cols = f.columns && f.columns.length > 0 ? f.columns : (f.rows && f.rows[0] ? Object.keys(f.rows[0]) : []);
      const slice5 = (f.rows || []).slice(0, 5);
      const maskedRows = slice5.map(r => {
        const rowObj = {};
        cols.forEach(c => {
          const val = r ? r[c] : '';
          rowObj[c] = shouldMask75 ? mask75PercentValue(val) : (val !== undefined ? val : '');
        });
        return rowObj;
      });

      sources.push({
        id: f.id,
        name: f.name,
        sheet: f.currentSheet || (f.sheets && f.sheets[0]) || "Sheet1",
        sheets: f.sheets || [],
        totalRows: (f.rows || []).length,
        columns: cols,
        rows: maskedRows,
        type: 'loaded',
        badge: '📁 File từ máy tính'
      });
    });
  }

  return sources;
}

// ========================================================
// KHỞI TẠO VÀ RENDER DANH SÁCH CHỌN FILE TRUYỀN DATA MẪU
// ========================================================

function initPromptSources() {
  const sources = getAvailablePromptSources();
  const btnTabMockMode = document.getElementById('btnTabMockMode');
  const isMockTab = btnTabMockMode && btnTabMockMode.classList.contains('active');
  const usingMock = typeof isUsingMockData !== 'undefined' && isUsingMockData;

  // Nếu chưa có nguồn nào được chọn, thiết lập các lựa chọn mặc định thông minh
  if (selectedPromptSourceIds.size === 0) {
    if (usingMock || isMockTab) {
      selectedPromptSourceIds.add('source_mock_excel');
    }
    // Tự động chọn các file tải lên từ máy nếu có
    sources.forEach(s => {
      if (s.type === 'loaded') {
        selectedPromptSourceIds.add(s.id);
      }
    });
    // Nếu vẫn chưa chọn gì (chưa có file tải lên và chưa ở mock mode), chọn mock_excel
    if (selectedPromptSourceIds.size === 0) {
      selectedPromptSourceIds.add('source_mock_excel');
    }
  } else {
    // Lọc bỏ những id không còn tồn tại trong sources hiện tại
    const validIds = new Set(sources.map(s => s.id));
    selectedPromptSourceIds = new Set([...selectedPromptSourceIds].filter(id => validIds.has(id)));
    if (selectedPromptSourceIds.size === 0 && sources.length > 0) {
      if (usingMock || isMockTab) {
        selectedPromptSourceIds.add('source_mock_excel');
      } else {
        selectedPromptSourceIds.add(sources[0].id);
      }
    }
  }

  renderPromptSourcePicker();
}

function renderPromptSourcePicker() {
  if (!promptSourcesList) return;
  const sources = getAvailablePromptSources();

  if (sources.length === 0) {
    promptSourcesList.innerHTML = `<div style="color: #888; font-size: 11px; padding: 4px;">Chưa có nguồn dữ liệu nào khả dụng.</div>`;
    return;
  }

  let html = '';
  sources.forEach(src => {
    const isChecked = selectedPromptSourceIds.has(src.id);
    const isMock = src.type === 'mock';
    const icon = isMock ? '⚡' : '📊';
    const badgeColor = isMock ? '#dcdcaa' : '#4ec9b0';

    html += `
      <div class="prompt-source-item ${isChecked ? 'active' : ''}">
        <label class="prompt-source-label">
          <input type="checkbox" class="prompt-source-chk" data-source-id="${src.id}" ${isChecked ? 'checked' : ''} />
          <span>${icon}</span>
          <span style="font-weight: 500;">${src.name}</span>
          <span style="color: #888; font-size: 11px;">(${src.sheet})</span>
          <span style="color: ${badgeColor}; font-size: 10px; background: ${badgeColor}18; border: 1px solid ${badgeColor}40; padding: 1px 5px; border-radius: 3px;">${src.badge}</span>
        </label>
        <span class="prompt-source-meta">${src.totalRows.toLocaleString('vi-VN')} dòng · ${src.columns.length} cột</span>
      </div>
    `;
  });

  promptSourcesList.innerHTML = html;

  // Gắn sự kiện khi click vào checkbox từng nguồn
  const checkboxes = promptSourcesList.querySelectorAll('.prompt-source-chk');
  checkboxes.forEach(chk => {
    chk.onchange = (e) => {
      const id = e.target.getAttribute('data-source-id');
      const itemRow = e.target.closest('.prompt-source-item');
      if (e.target.checked) {
        selectedPromptSourceIds.add(id);
        if (itemRow) itemRow.classList.add('active');
      } else {
        selectedPromptSourceIds.delete(id);
        if (itemRow) itemRow.classList.remove('active');
      }
      buildPrompt();
    };
  });
}

// Xử lý nút "Chọn tất cả" nguồn
if (btnSelectAllPromptSources) {
  btnSelectAllPromptSources.onclick = () => {
    const sources = getAvailablePromptSources();
    sources.forEach(s => selectedPromptSourceIds.add(s.id));
    renderPromptSourcePicker();
    buildPrompt();
  };
}

// Xử lý nút "Bỏ chọn hết" nguồn
if (btnUnselectAllPromptSources) {
  btnUnselectAllPromptSources.onclick = () => {
    selectedPromptSourceIds.clear();
    renderPromptSourcePicker();
    buildPrompt();
  };
}

// Xử lý checkbox tổng: Bật/Tắt đính kèm dữ liệu mẫu
if (chkIncludeSampleData) {
  chkIncludeSampleData.onchange = () => {
    const isInclude = chkIncludeSampleData.checked;
    if (promptSourcePicker) {
      promptSourcePicker.style.opacity = isInclude ? '1' : '0.4';
      promptSourcePicker.style.pointerEvents = isInclude ? 'auto' : 'none';
    }
    buildPrompt();
  };
}

// Hiển thị trực tiếp bảng xem trước Mock Data đã che mờ trong Modal Prompt
function renderMockPreview(sources) {
  if (!mockPreviewContent) return;
  if (!sources || sources.length === 0) {
    mockPreviewContent.innerHTML = '<div style="color: #888; padding: 10px; font-style: italic; text-align: center;">🚫 Đã tắt đính kèm dữ liệu mẫu hoặc chưa chọn file nào.<br>Prompt AI sẽ KHÔNG chứa bất kỳ bảng dữ liệu mẫu nào.</div>';
    return;
  }
  let html = '';
  sources.forEach((s, idx) => {
    html += `
    <div style="padding: 6px; ${idx > 0 ? 'border-top: 1px solid #333; margin-top: 6px;' : ''}">
      <div style="color: #4ec9b0; font-weight: 600; margin-bottom: 4px; display: flex; justify-content: space-between; align-items: center;">
        <span>📁 ${s.name} (${s.rows.length} dòng mẫu, 75% ký tự ô bị che mờ '*')</span>
        <span style="color: #888; font-size: 10px;">Sheet: ${s.sheet}</span>
      </div>
      <div style="overflow-x: auto;">
        <table style="width: 100%; border-collapse: collapse; font-size: 11px; background: #141416;">
          <thead>
            <tr style="background: #252528; color: #9cdcfe;">
              ${s.columns.map(c => `<th style="border: 1px solid #3c3c3c; padding: 3px 6px; text-align: left; white-space: nowrap;">${c}</th>`).join('')}
            </tr>
          </thead>
          <tbody>
            ${s.rows.map(r => `
              <tr>
                ${s.columns.map(c => `<td style="border: 1px solid #2a2a2d; padding: 3px 6px; font-family: Consolas, monospace; color: #ce9178; white-space: nowrap;">${r[c] !== undefined ? r[c] : ''}</td>`).join('')}
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    </div>`;
  });
  mockPreviewContent.innerHTML = html;
}

if (toggleMockPreview && mockPreviewContent) {
  toggleMockPreview.onclick = () => {
    const isHidden = mockPreviewContent.style.display === 'none';
    mockPreviewContent.style.display = isHidden ? 'block' : 'none';
    if (mockPreviewBadge) mockPreviewBadge.textContent = isHidden ? '▲ Thu gọn' : '▼ Mở bảng xem';
  };
}

// ========================================================
// HÀM BUILD PROMPT CHÍNH
// ========================================================

function buildPrompt() {
  if (!inputDataDesc || !businessLogicDesc || !outputDesc || !fullPromptOutput) return;

  const inputDesc = inputDataDesc.value.trim();
  const logicDesc = businessLogicDesc.value.trim() || "[Chưa nhập yêu cầu cụ thể - Viết script lọc, phân tích và xử lý dữ liệu theo nghiệp vụ]";
  const outDesc = outputDesc.value.trim() || "[Xuất file Excel kết quả tính toán và in báo cáo tóm tắt ra console]";

  const chkMask75 = document.getElementById('chkMask75') ? document.getElementById('chkMask75').checked : true;
  const chkCleanNum = document.getElementById('chkCleanNum') ? document.getElementById('chkCleanNum').checked : true;
  const chkExport = document.getElementById('chkExport') ? document.getElementById('chkExport').checked : true;
  const chkSummaryTable = document.getElementById('chkSummaryTable') ? document.getElementById('chkSummaryTable').checked : true;
  const chkJoinOpt = document.getElementById('chkJoinOpt') ? document.getElementById('chkJoinOpt').checked : true;

  const isIncludeSample = chkIncludeSampleData ? chkIncludeSampleData.checked : true;
  const allSources = getAvailablePromptSources(chkMask75);

  // Danh sách các nguồn thực sự được chọn để đưa dữ liệu mẫu vào Prompt
  const activeSources = isIncludeSample ? allSources.filter(s => selectedPromptSourceIds.has(s.id)) : [];

  // Cập nhật số file xem trước trên giao diện
  if (previewFilesCount) {
    previewFilesCount.textContent = activeSources.length;
  }

  // Render bảng xem trước
  renderMockPreview(activeSources);

  // 1. Xác định file nguồn và biến xử lý trong code
  let targetFilesDirective = "";
  const hasMockSelected = activeSources.some(s => s.type === 'mock');
  const hasLoadedSelected = activeSources.some(s => s.type === 'loaded');

  if (activeSources.length === 0) {
    // Không có file nào được chọn truyền mẫu hoặc đã tắt đính kèm mẫu
    targetFilesDirective = `### 🎯 CHỈ ĐỊNH FILE NGUỒN VÀ BIẾN XỬ LÝ:
- **Tùy chọn dữ liệu mẫu**: Người dùng đã tắt đính kèm dữ liệu mẫu trong prompt này.
- **YÊU CẦU ĐỐI VỚI CODE CỦA BẠN**:
  Code của bạn xử lý dữ liệu thông qua biến toàn cục **\`rawRows\`** (mảng các object JSON đại diện cho file đang chọn) hoặc object **\`files\`**:
\`\`\`javascript
// Trỏ vào dữ liệu file đang kích hoạt:
const data = rawRows; 
\`\`\`
- Nếu xử lý nhiều file, truy cập thông qua object \`files["tên_file.xlsx"]\`.`;
  } else if (activeSources.length === 1) {
    const s = activeSources[0];
    if (s.type === 'mock') {
      targetFilesDirective = `### 🎯 CHỈ ĐỊNH FILE NGUỒN VÀ BIẾN XỬ LÝ (BẢNG MÔ PHỎNG EXCEL):
- **Nguồn dữ liệu**: BẢNG MÔ PHỎNG EXCEL (\`${s.name}\`, Sheet: "${s.sheet}").
- **YÊU CẦU BẮT BUỘC ĐỐI VỚI CODE CỦA BẠN**:
  Code của bạn PHẢI trỏ trực tiếp vào dữ liệu bảng mô phỏng này thông qua biến **\`rawRows\`**:
\`\`\`javascript
// Trỏ trực tiếp vào bảng dữ liệu mô phỏng Excel
const data = rawRows; 
// (hoặc qua object files: const data = files["${s.name}"];)
\`\`\`
- Toàn bộ tính toán, lọc, gộp hãy thực hiện trực tiếp trên biến \`data\` này.`;
    } else {
      targetFilesDirective = `### 🎯 CHỈ ĐỊNH FILE NGUỒN VÀ BIẾN XỬ LÝ (FILE TẢI LÊN TỪ MÁY):
- **Nguồn dữ liệu**: File \`${s.name}\` (Sheet: "${s.sheet}", tổng ${s.totalRows.toLocaleString('vi-VN')} dòng).
- **YÊU CẦU BẮT BUỘC ĐỐI VỚI CODE CỦA BẠN**:
  Code của bạn PHẢI trỏ vào file này thông qua biến **\`rawRows\`** hoặc **\`files["${s.name}"]\`**:
\`\`\`javascript
// Trỏ mặc định vào file Excel đã nạp
const data = rawRows; // (hoặc: const data = files["${s.name}"];)
\`\`\`
- Toàn bộ tính toán, lọc, gộp hãy thực hiện trực tiếp trên biến \`data\` này.`;
    }
  } else {
    // Nhiều file được chọn đồng thời
    targetFilesDirective = `### 🎯 CHỈ ĐỊNH TẤT CẢ CÁC FILE VÀ BIẾN XỬ LÝ (CHẾ ĐỘ ${activeSources.length} FILE ĐƯỢC CHỌN):
- **Trạng thái hiện tại**: Người dùng đã chọn **${activeSources.length} nguồn dữ liệu Excel** để xử lý.
- **YÊU CẦU BẮT BUỘC ĐỐI VỚI CODE CỦA BẠN**:
  Code của bạn CẦN TRỎ VÀO **TẤT CẢ ${activeSources.length} NGUỒN FILE** tương ứng thông qua object **\`files\`** hoặc biến **\`rawRows\`** với tên file chính xác như sau:
\`\`\`javascript
// Khai báo trỏ chính xác vào từng file theo tên trong object files:
${activeSources.map((f, i) => `const dataFile${i + 1} = files["${f.name}"]; // (${f.badge} - ${f.totalRows.toLocaleString('vi-VN')} dòng, Sheet: "${f.sheet}")`).join('\n')}
\`\`\`
- **NẾU BÀI TOÁN YÊU CẦU GỘP TOÀN BỘ DỮ LIỆU CỦA CÁC FILE LẠI**:
\`\`\`javascript
const allRows = mergeAllRows(); // Tự động gộp dữ liệu của tất cả file (mỗi dòng có thêm trường _sourceFile)
\`\`\`
- **CHÚ Ý QUAN TRỌNG**: TUYỆT ĐỐI KHÔNG tự ý bịa tên file khác. Bắt buộc dùng đúng các key: ${activeSources.map(f => `\`files["${f.name}"]\``).join(', ')}.`;
  }

  // 2. Mô tả chi tiết dữ liệu đầu vào & Mock Data
  let filesDetailedInfo = "";
  if (!isIncludeSample || activeSources.length === 0) {
    filesDetailedInfo = `*(Không đính kèm dữ liệu mẫu trong prompt này theo thiết lập của người dùng. Code của bạn hãy dựa vào mô tả nghiệp vụ và cấu trúc biến chuẩn.)*\n`;
  } else {
    activeSources.forEach((src, idx) => {
      let callSyntax = "";
      if (src.type === 'mock') {
        callSyntax = `Biến \`rawRows\` *(hoặc \`files["${src.name}"]\`)*`;
      } else {
        const isLoadedActive = (typeof activeFileId !== 'undefined' && src.id === activeFileId);
        callSyntax = `\`files["${src.name}"]\` ${isLoadedActive ? '*(hoặc biến `rawRows` vì đây là file đang active)*' : ''}`;
      }

      const mockSection = `
- **Mô phỏng 5 dòng Mock Data mẫu (${chkMask75 ? 'Đã che mờ 75% dữ liệu từng ô để bảo mật' : 'Dữ liệu mẫu 5 dòng đầu'}):**
  > *Lưu ý cho AI: Dữ liệu bên dưới được trích xuất từ 5 dòng đầu tiên của file dữ liệu thực tế. Các ký tự \`*\` thể hiện 75% phần đuôi của dữ liệu đã được làm mờ để bảo mật thông tin. Hãy quan sát 25% ký tự đầu và tên cột để hiểu chính xác kiểu dữ liệu và quy luật dữ liệu đầu vào.*

  **Bảng mô phỏng cấu trúc (Markdown Table):**
${formatMarkdownTable(src.columns, src.rows)}

  **Mẫu JSON 5 dòng mô phỏng (Đại diện cho cấu trúc từng phần tử \`row\` trong runtime):**
\`\`\`json
${JSON.stringify(src.rows, null, 2)}
\`\`\`
`;

      filesDetailedInfo += `
### NGUỒN DỮ LIỆU #${idx + 1}: \`${src.name}\` (${src.badge})
- **Cách gọi trong code**: ${callSyntax}
- **Sheet**: "${src.sheet}" ${src.sheets && src.sheets.length > 1 ? `(File có ${src.sheets.length} sheet: [${src.sheets.map(s => `"${s}"`).join(", ")}])` : ''}
- **Tổng số dòng**: ${src.totalRows.toLocaleString('vi-VN')} dòng
- **Danh sách cột chính xác (${src.columns.length} cột)**:
${JSON.stringify(src.columns, null, 2)}
${mockSection}
`;
    });
  }

  let extraRules = [];

  // Quy tắc trỏ file
  let fileTargetRule = "";
  if (activeSources.length === 1 && activeSources[0].type === 'mock') {
    fileTargetRule = `1. **QUY TẮC TRỎ ĐÚNG FILE BẢNG MÔ PHỎNG**:
   - Bạn đang ở chế độ Mô Phỏng Excel. BẮT BUỘC trỏ vào biến \`rawRows\` để lấy dữ liệu: \`const data = rawRows;\`.`;
  } else if (activeSources.length > 1) {
    fileTargetRule = `1. **QUY TẮC TRỎ CHÍNH XÁC VÀO CÁC FILE ĐƯỢC CHỌN**:
   - Có ${activeSources.length} file được chọn. BẮT BUỘC truy cập các file thông qua object \`files\` với đúng tên file đã nạp: ${activeSources.map(f => `\`files["${f.name}"]\``).join(', ')}.
   - Khi kết nối (JOIN) giữa các file, hãy lập Map theo khóa chính để tra cứu O(1).`;
  } else if (activeSources.length === 1) {
    fileTargetRule = `1. **QUY TẮC TRỎ ĐÚNG FILE DUY NHẤT ĐÃ CHỌN**:
   - BẮT BUỘC trỏ vào \`rawRows\` hoặc \`files["${activeSources[0].name}"]\`: \`const data = rawRows;\`.`;
  } else {
    fileTargetRule = `1. **QUY TẮC TRỎ ĐÚNG DỮ LIỆU ĐẦU VÀO**:
   - Trỏ trực tiếp vào mảng \`rawRows\`: \`const data = rawRows;\`.`;
  }
  extraRules.push(fileTargetRule);

  // Chỉ thêm quy tắc về Mock Data nếu thực sự có truyền mock data
  if (isIncludeSample && activeSources.length > 0) {
    extraRules.push(`2. **HIỂU MOCK DATA ĐÃ CHE MỜ 75%**:
   - Dữ liệu mẫu 5 dòng được cung cấp ở mục 2 đã được che mờ 75% giá trị từng ô (thay bằng ký tự '*') cho mục đích bảo mật thông tin.
   - Khi chạy thực tế trong công cụ, code JavaScript của bạn sẽ chạy trên DỮ LIỆU THẬT KHÔNG BỊ CHE MỜ.
   - Do đó, KHÔNG viết code tìm kiếm hay xử lý ký tự '*' trừ khi yêu cầu nghiệp vụ yêu cầu, mà hãy dùng mock data để biết cấu trúc key của object \`row["Tên Cột"]\` và quy luật kiểu dữ liệu.`);
  }

  if (chkCleanNum) {
    extraRules.push(`3. **ÉP KIỂU SỐ AN TOÀN (NUMBER PARSING)**:
   - Số tiền/số lượng từ Excel có thể chứa dấu phẩy/chấm phân cách nghìn hoặc ký tự đơn vị (vd: "1,500,000", "1.500.000 đ", null).
   - LUÔN tạo và sử dụng hàm parse số an toàn:
     \`const parseNum = v => typeof v === 'number' ? (isNaN(v) ? 0 : v) : parseFloat(String(v || '').replace(/[^0-9.-]/g, '')) || 0;\``);
  }

  if (chkSummaryTable) {
    extraRules.push(`4. **IN BÁO CÁO TỔNG QUAN RA TERMINAL**:
   - Sử dụng \`console.log()\` và \`console.table()\` để in kết quả tóm tắt rõ ràng (tổng số dòng kết quả, tổng tiền, mẫu vài dòng đầu...) để người dùng theo dõi được tiến trình.`);
  }

  if (chkExport) {
    extraRules.push(`5. **XUẤT FILE KẾT QUẢ**:
   - Sau khi tính toán/lọc xong, gọi hàm \`exportExcel(ketQuaArray, "KetQua_${Date.now().toString().slice(-4)}.xlsx")\` để người dùng tải file kết quả về máy.`);
  }

  if (chkJoinOpt && activeSources.length > 1) {
    extraRules.push(`6. **TỐI ƯU JOIN ĐA FILE**:
   - Khi nối dữ liệu giữa nhiều file, hãy lập Map hoặc Object index theo khóa chính (Key) để tra cứu O(1), tránh dùng 2 vòng lặp lồng nhau O(N^2) gây đơ trình duyệt khi dữ liệu lớn.`);
  }

  const inputSection = inputDesc
    ? `${targetFilesDirective}\n\n- **Ghi chú đặc thù về Input từ người dùng**: ${inputDesc}\n\n- **Cấu trúc & Dữ liệu mô phỏng từng file**:\n${filesDetailedInfo}`
    : `${targetFilesDirective}\n\n- **Cấu trúc & Dữ liệu mô phỏng từng file**:\n${filesDetailedInfo}`;

  const prompt = `Bạn là Kỹ Sư Lập Trình Dữ Liệu JavaScript (Senior Data Processing Specialist).
Hãy viết một đoạn mã Vanilla JavaScript thuần chạy trực tiếp trong trình duyệt để xử lý dữ liệu theo đúng yêu cầu nghiệp vụ bên dưới.

---
## 1. MÔI TRƯỜNG THỰC THI (BROWSER RUNTIME)
Code sẽ chạy trực tiếp qua \`new Function('rawRows', 'files', 'fileSheets', 'getFile', 'getSheet', 'mergeAllRows', 'XLSX', 'exportExcel', 'context', <code>)\`.
Các biến và hàm toàn cục có sẵn:
- \`files\`: Object chứa toàn bộ dữ liệu các file đã nạp dạng \`{ "tên_file.xlsx": [các dòng...], ... }\`. Có thể truy cập trực tiếp từng sheet cụ thể qua \`files["tên_file.xlsx"]["Tên_Sheet"]\` hoặc \`files["tên_file.xlsx"].sheets["Tên_Sheet"]\`.
- \`fileSheets\`: Object phân cấp rõ ràng theo từng file và sheet: \`{ "tên_file.xlsx": { "Sheet1": [...], "Sheet2": [...] } }\`.
- \`getFile(query)\`: Hàm tìm file linh hoạt theo chỉ số \`getFile(0)\` hoặc theo tên \`getFile("tên_file")\`.
- \`getSheet(fileName, sheetName)\` hoặc \`getSheet(sheetName)\`: Hàm lấy trực tiếp mảng dòng của sheet mong muốn.
- \`rawRows\`: Mảng JSON các dòng của file đang chọn.
- \`mergeAllRows(fileNames)\`: Hàm tiện ích gộp toàn bộ dòng của các file đã nạp (mỗi dòng tự động có thêm trường \`_sheet\` và \`_file\`).
- \`exportExcel(dataArray, fileName)\`: Hàm xuất mảng JSON ra file Excel và tự động kích hoạt tải xuống trình duyệt.
- \`context\`: Object lưu trữ/chia sẻ dữ liệu nếu chạy qua Pipeline đa script.
- \`XLSX\`: Thư viện SheetJS toàn cục.
- \`console.log()\`, \`console.warn()\`, \`console.error()\`, \`console.table()\`: Hiển thị trực tiếp lên Terminal Monitor của công cụ.

---
## 2. MÔ TẢ DỮ LIỆU ĐẦU VÀO (INPUT)
${inputSection}

---
## 3. YÊU CẦU NGHIỆP VỤ CỤ THỂ CẦN XỬ LÝ (BUSINESS LOGIC)
${logicDesc}

---
## 4. KẾT QUẢ ĐẦU RA MONG MUỐN (EXPECTED OUTPUT)
${outDesc}

---
## 5. QUY TẮC BẮT BUỘC ĐỂ CODE CHẠY CHÍNH XÁC 100% (CRITICAL RULES)
1. **TRUY CẬP CỘT AN TOÀN**:
   - Tên cột trong file Excel/CSV có thể có khoảng trắng hoặc dấu tiếng Việt (ví dụ: \`row["Họ và Tên"]\`, \`row["Đơn giá"]\`).
   - LUÔN dùng cú pháp ngoặc vuông \`row["Tên Cột"]\`, KHÔNG dùng dấu chấm \`row.Tên Cột\`.
   - Luôn kiểm tra null/undefined trước khi gọi hàm chuỗi: \`String(row["Cột"] || "").trim()\`.

${extraRules.join('\n\n')}

*. **ĐỊNH DẠNG ĐẦU RA BẮT BUỘC**:
   - CHỈ TRẢ VỀ DUY NHẤT một khối mã JavaScript trong \`\`\`javascript ... \`\`\`.
   - KHÔNG viết lời chào hỏi, giải thích dài dòng trước hoặc sau khối mã.
   - KHÔNG dùng cú pháp \`import\`, \`export\`, \`require\` (vì code chạy trực tiếp trong browser).
   - KHÔNG bọc toàn bộ code trong \`function main() {}\` mà viết thẳng các câu lệnh ở cấp độ cao nhất (top-level code).`;

  fullPromptOutput.value = prompt.trim();
}

// ========================================================
// SỰ KIỆN MỞ/ĐÓNG MODAL VÀ CẬP NHẬT DỮ LIỆU
// ========================================================

if (btnOpenPrompt && promptModal) {
  btnOpenPrompt.onclick = () => {
    initPromptSources();
    promptModal.style.display = 'flex';
    buildPrompt();
  };
}

if (btnCloseModal && promptModal) {
  btnCloseModal.onclick = () => { promptModal.style.display = 'none'; };
}

if (btnCloseModalX && promptModal) {
  btnCloseModalX.onclick = () => { promptModal.style.display = 'none'; };
}

if (promptModal) {
  promptModal.onclick = (e) => {
    if (e.target === promptModal) promptModal.style.display = 'none';
  };
}

// Tự động cập nhật prompt khi thay đổi các checkbox tùy chọn
['chkMask75', 'chkCleanNum', 'chkExport', 'chkSummaryTable', 'chkJoinOpt'].forEach(id => {
  const el = document.getElementById(id);
  if (el) el.onchange = buildPrompt;
});

if (inputDataDesc) inputDataDesc.oninput = buildPrompt;
if (businessLogicDesc) businessLogicDesc.oninput = buildPrompt;
if (outputDesc) outputDesc.oninput = buildPrompt;

if (btnCopyPrompt && fullPromptOutput) {
  btnCopyPrompt.onclick = () => {
    fullPromptOutput.select();
    navigator.clipboard.writeText(fullPromptOutput.value).then(() => {
      alert("Đã sao chép prompt hoàn chỉnh vào Clipboard!");
    }).catch(() => {
      document.execCommand('copy');
      alert("Đã sao chép prompt!");
    });
  };
}

