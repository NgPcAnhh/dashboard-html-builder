/**
 * app.js - Khởi tạo Ứng dụng & Quản lý Layout Resizers
 */

// ========================================================
// 1. TÙY CHỈNH KÍCH THƯỚC GIỮA 2 BÊN (TRÁI / PHẢI)
// ========================================================
const resizer = document.getElementById('dragResizer');
const panelLeft = document.getElementById('panelLeft');
let isResizing = false;

if (resizer && panelLeft) {
  resizer.addEventListener('mousedown', (e) => {
    isResizing = true;
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';
    resizer.classList.add('is-dragging');
  });

  window.addEventListener('mousemove', (e) => {
    if (!isResizing) return;
    const totalWidth = document.body.clientWidth;
    const newWidth = e.clientX;
    if (newWidth >= 300 && (totalWidth - newWidth) >= 300) {
      const pct = (newWidth / totalWidth) * 100;
      panelLeft.style.width = pct + '%';
    }
  });

  window.addEventListener('mouseup', () => {
    if (isResizing) {
      isResizing = false;
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
      resizer.classList.remove('is-dragging');
    }
  });

  resizer.addEventListener('dblclick', () => {
    panelLeft.style.width = '50%';
  });
}

// ========================================================
// 2. THANH KÉO RESIZER ĐIỀU CHỈNH CHIỀU CAO BẢNG EXCEL & TERMINAL
// ========================================================
const dragResizerRight = document.getElementById('dragResizerRight');
const sourceContainer = document.getElementById('sourceContainer');
const panelRight = document.getElementById('panelRight');
let isResizingRight = false;
let startYRight = 0;
let startHeightSource = 0;

if (dragResizerRight && sourceContainer && panelRight) {
  dragResizerRight.addEventListener('mousedown', (e) => {
    isResizingRight = true;
    startYRight = e.clientY;
    startHeightSource = sourceContainer.offsetHeight;
    dragResizerRight.classList.add('is-dragging');
    document.body.style.cursor = 'row-resize';
    document.body.style.userSelect = 'none';
  });

  window.addEventListener('mousemove', (e) => {
    if (!isResizingRight) return;
    const dy = e.clientY - startYRight;
    const newHeight = startHeightSource + dy;
    const totalH = panelRight.offsetHeight;

    if (newHeight > 160 && newHeight < totalH - 120) {
      sourceContainer.style.flex = 'none';
      sourceContainer.style.height = `${newHeight}px`;
    }
  });

  window.addEventListener('mouseup', () => {
    if (isResizingRight) {
      isResizingRight = false;
      dragResizerRight.classList.remove('is-dragging');
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
    }
  });
}

// ========================================================
// 3. CHUYỂN ĐỔI CHẾ ĐỘ: FILE TỪ MÁY <-> MÔ PHỎNG EXCEL
// ========================================================
const btnTabFileMode = document.getElementById('btnTabFileMode');
const btnTabMockMode = document.getElementById('btnTabMockMode');
const sourceFileView = document.getElementById('sourceFileView');
const sourceMockView = document.getElementById('sourceMockView');

if (btnTabFileMode && btnTabMockMode && sourceFileView && sourceMockView) {
  btnTabFileMode.onclick = () => {
    btnTabFileMode.classList.add('active');
    btnTabMockMode.classList.remove('active');
    sourceFileView.style.display = 'flex';
    sourceMockView.style.display = 'none';
    const btnUse = document.getElementById('btnUseMockData');
    if (btnUse) btnUse.style.display = 'none';
    if (typeof buildPrompt === 'function') buildPrompt();
  };

  btnTabMockMode.onclick = () => {
    btnTabMockMode.classList.add('active');
    btnTabFileMode.classList.remove('active');
    sourceFileView.style.display = 'none';
    sourceMockView.style.display = 'flex';
    const btnUse = document.getElementById('btnUseMockData');
    if (btnUse) btnUse.style.display = 'inline-flex';
    if (typeof buildPrompt === 'function') buildPrompt();
  };
}

// ========================================================
// 4. KHỞI CHẠY ỨNG DỤNG (BOOTSTRAP)
// ========================================================
window.addEventListener('DOMContentLoaded', () => {
  // Render tabs script ban đầu và chuyển tab 1
  if (typeof switchTab === 'function') {
    switchTab(activeScriptId);
  }

  // Khởi tạo bảng mô phỏng Excel ngay khi mở trang
  if (typeof renderInlineExcel === 'function') {
    renderInlineExcel();
  }
});
