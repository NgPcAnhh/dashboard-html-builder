/**
 * script-runner.js - Thực thi Mã JS và Quản lý Lưu / Nạp Template
 */

// ========================================================
// 1. CHẠY SCRIPT
// ========================================================
const btnRun = document.getElementById('btnRun');
if (btnRun) {
  btnRun.onclick = () => {
    // Đảm bảo dữ liệu mới nhất được đồng bộ nếu đang ở bảng mô phỏng
    if (typeof isUsingMockData !== 'undefined' && isUsingMockData && typeof syncInlineExcelToDataState === 'function') {
      syncInlineExcelToDataState(false);
    }
    const activeScript = scripts.find(s => s.id === activeScriptId);
    const scriptName = activeScript ? activeScript.name : "script";

    printToTerminal(`\n--- CHẠY [${scriptName}] ---`, "log-info");
    try {
      const sharedContext = window.pipelineContext || {};
      const userFunc = new Function('rawRows', 'files', 'fileSheets', 'getFile', 'getSheet', 'mergeAllRows', 'XLSX', 'exportExcel', 'context', codeEditor.value);
      userFunc(rawRows, files, fileSheets, window.getFile, window.getSheet, mergeAllRows, window.XLSX, window.exportExcel, sharedContext);
    } catch (err) {
      console.error(`Lỗi runtime trong [${scriptName}]: ` + err.stack);
    }
  };
}

// ========================================================
// 2. TIỆN ÍCH FILE SYSTEM: THƯ MỤC "data_processing_template"
// ========================================================
async function getTemplateDirectoryHandle(promptUserIfEmpty = true) {
  if (!savedDirectoryHandle) {
    if (!promptUserIfEmpty) return null;
    alert("Vui lòng chọn thư mục chứa công cụ (hoặc chọn thẳng folder 'data_processing_template'). Hệ thống sẽ tự động liên kết tới folder 'data_processing_template'.");
    savedDirectoryHandle = await window.showDirectoryPicker();
  }
  if (savedDirectoryHandle.name === 'data_processing_template') {
    return savedDirectoryHandle;
  }
  try {
    const subDir = await savedDirectoryHandle.getDirectoryHandle('data_processing_template', { create: true });
    return subDir;
  } catch (e) {
    return savedDirectoryHandle;
  }
}

// ========================================================
// 3. LƯU SCRIPT VÀO FOLDER "data_processing_template" HOẶC TẢI VỀ
// ========================================================
const btnSaveToFolder = document.getElementById('btnSaveToFolder');
if (btnSaveToFolder) {
  btnSaveToFolder.onclick = async () => {
    const activeScript = scripts.find(s => s.id === activeScriptId);
    const defaultName = activeScript ? activeScript.name : "main_process.js";
    const codeContent = codeEditor.value;

    const action = confirm("Nhấn 'OK' để lưu file script hiện tại vào folder 'data_processing_template'.\nNhấn 'Cancel' nếu muốn lưu TẤT CẢ các file script JS vào folder 'data_processing_template'.");

    if (action) {
      // Lưu script hiện tại
      let fileName = prompt("Nhập tên file JS cần lưu (vào folder data_processing_template):", defaultName);
      if (!fileName) return;
      if (!fileName.toLowerCase().endsWith('.js')) fileName += '.js';
      await saveSingleJsFile(codeContent, fileName);
    } else {
      // Lưu toàn bộ scripts
      await saveAllJsFiles();
    }
  };
}

async function saveSingleJsFile(content, fileName) {
  if ('showDirectoryPicker' in window) {
    try {
      const dirHandle = await getTemplateDirectoryHandle();
      const fileHandle = await dirHandle.getFileHandle(fileName, { create: true });
      const writable = await fileHandle.createWritable();
      await writable.write(content);
      await writable.close();
      const msg = `[Hệ thống] ĐÃ LƯU THÀNH CÔNG: data_processing_template/${fileName}`;
      console.log(msg);
      printToTerminal(msg, "log-info");
      alert(`Đã lưu thành công vào folder data_processing_template/${fileName}`);
    } catch (err) {
      if (err.name !== 'AbortError') {
        console.warn("[Ghi file] Gặp lỗi: " + err.message);
        fallbackDownload(content, fileName);
      }
    }
  } else {
    fallbackDownload(content, fileName);
  }
}

async function saveAllJsFiles() {
  if ('showDirectoryPicker' in window) {
    try {
      const dirHandle = await getTemplateDirectoryHandle();
      for (const s of scripts) {
        let fName = s.name;
        if (!fName.toLowerCase().endsWith('.js')) fName += '.js';
        const fileHandle = await dirHandle.getFileHandle(fName, { create: true });
        const writable = await fileHandle.createWritable();
        await writable.write(s.code);
        await writable.close();
      }
      const msg = `[Hệ thống] ĐÃ LƯU THÀNH CÔNG ${scripts.length} file JS vào folder 'data_processing_template'!`;
      console.log(msg);
      printToTerminal(msg, "log-info");
      alert(`Đã lưu thành công ${scripts.length} file JS vào folder 'data_processing_template'!`);
    } catch (err) {
      if (err.name !== 'AbortError') {
        console.warn("[Ghi file] " + err.message);
        scripts.forEach(s => fallbackDownload(s.code, s.name));
      }
    }
  } else {
    scripts.forEach(s => fallbackDownload(s.code, s.name));
  }
}

function fallbackDownload(content, fileName) {
  const blob = new Blob([content], { type: 'text/javascript' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = fileName;
  a.click();
  URL.revokeObjectURL(a.href);
  const msg = `[Hệ thống] Đã tải về file: ${fileName} (Hãy di chuyển vào folder data_processing_template)`;
  console.log(msg);
  printToTerminal(msg, "log-info");
}

// ========================================================
// 4. NẠP FILE JS TỪ THƯ MỤC "data_processing_template"
// ========================================================
const btnLoadJs = document.getElementById('btnLoadJs');
const jsFileInput = document.getElementById('jsFileInput');

async function loadJsFromTemplateFolder() {
  if ('showOpenFilePicker' in window) {
    try {
      const dirHandle = await getTemplateDirectoryHandle(true);
      if (dirHandle) {
        const fileHandles = await window.showOpenFilePicker({
          multiple: true,
          startIn: dirHandle,
          types: [
            {
              description: 'JavaScript Template Files (*.js)',
              accept: {
                'text/javascript': ['.js'],
                'application/javascript': ['.js']
              }
            }
          ]
        });

        if (fileHandles && fileHandles.length > 0) {
          let firstAddedId = null;
          for (const fh of fileHandles) {
            const file = await fh.getFile();
            const code = await file.text();
            const newId = Date.now() + Math.floor(Math.random() * 1000);
            scripts.push({
              id: newId,
              name: file.name,
              code: code
            });
            if (!firstAddedId) firstAddedId = newId;
          }
          const msg = `[Hệ thống] Đã nạp thành công ${fileHandles.length} file script JS từ folder data_processing_template!`;
          console.log(msg);
          printToTerminal(msg, "log-info");
          if (typeof renderTabs === 'function') renderTabs();
          if (firstAddedId && typeof switchTab === 'function') switchTab(firstAddedId);
          return;
        }
      }
    } catch (err) {
      if (err.name === 'AbortError') return;
      console.warn("[Nạp file JS] Chuyển sang fallback file input: ", err);
    }
  }

  // Fallback nếu trình duyệt không hỗ trợ File System Access API
  if (jsFileInput) jsFileInput.click();
}

if (btnLoadJs) {
  btnLoadJs.onclick = loadJsFromTemplateFolder;
}

if (jsFileInput) {
  jsFileInput.onchange = (e) => {
    const fileList = Array.from(e.target.files);
    if (!fileList.length) return;

    let firstAddedId = null;
    let loadedCount = 0;

    fileList.forEach(file => {
      const reader = new FileReader();
      reader.onload = (evt) => {
        const newId = Date.now() + Math.floor(Math.random() * 1000);
        scripts.push({
          id: newId,
          name: file.name,
          code: evt.target.result
        });
        if (!firstAddedId) firstAddedId = newId;
        loadedCount++;

        if (loadedCount === fileList.length) {
          console.log(`[Hệ thống] Đã nạp thành công ${loadedCount} file script JS!`);
          printToTerminal(`[Hệ thống] Đã nạp thành công ${loadedCount} file script JS!`, "log-info");
          if (typeof renderTabs === 'function') renderTabs();
          if (firstAddedId && typeof switchTab === 'function') switchTab(firstAddedId);
        }
      };
      reader.readAsText(file);
    });
  };
}
