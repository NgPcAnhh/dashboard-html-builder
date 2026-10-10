/**
 * state.js - Quản lý Trạng thái Toàn cục & Các phần tử DOM chính
 */

// ========================================================
// CÁC BIẾN DỮ LIỆU TOÀN CỤC
// ========================================================
let loadedFiles = []; // [ { id, name, size, sizeStr, workbook, sheets: [], sheetsData: {}, selectedSheets: [], rows: [], columns: [] } ]
let activeFileId = null;

let rawRows = [];
let files = {}; // Map: { "file1.xlsx": [...rows], "file2.csv": [...rows] }
let fileSheets = {}; // Map: { "file1.xlsx": { "Sheet1": [...], "Sheet2": [...] } }
let dataColumns = [];
let sampleRecords = [];
let savedDirectoryHandle = null;

// Quản lý nhiều script JS
let scripts = [
  {
    id: 1,
    name: "main_process.js",
    code: `// ========================================================
// HƯỚNG DẪN TRUY CẬP ĐA FILE & ĐA SHEET (TÙY Ý CHỌN & GỌI)
// ========================================================
// 1. files: Object chứa TẤT CẢ các file Excel/CSV đã nạp
//    const file1 = files['ten_file_1.xlsx'];
//    const file2 = files['ten_file_2.xlsx'];
//
// 2. MẶC ĐỊNH CHỌN & TRUY CẬP TẤT CẢ CÁC SHEET TỪ FILE EXCEL:
//    const sheetA = files['ten_file_1.xlsx']['Sheet1'];
//    const sheetB = files['ten_file_1.xlsx']['Sheet2'];
//    hoặc: const sheetC = getSheet('ten_file_1.xlsx', 'Sheet1');
//    hoặc: const sheetD = getSheet('Sheet2');
//
// 3. getFile(0), getFile(1) hoặc getFile('ten_file'): Lấy theo chỉ số hoặc tên
// 4. rawRows: Mảng dòng của file đầu tiên / file đang chọn
// 5. mergeAllRows(): Tự động gộp toàn bộ dòng từ tất cả các file & sheet
// 6. exportExcel(dataArray, fileName): Xuất dữ liệu ra file Excel
// ========================================================

console.log("=== DANH SÁCH FILE & SHEET ĐÃ NẠP ===");
const fileNames = Object.keys(files);

if (fileNames.length === 0) {
  console.log("Chưa có file nào! Hãy kéo thả 1 hoặc nhiều file Excel/CSV vào ô bên phải.");
} else {
  console.log(\`Đã nạp \${fileNames.length} file:\`);
  fileNames.forEach((name, i) => {
    const fileData = files[name];
    const sheetList = fileData.sheets ? Object.keys(fileData.sheets) : [];
    console.log(\`  \${i + 1}. [\${name}] - Tổng \${fileData.length} dòng (Sheets: \${sheetList.join(', ') || 'Mặc định'})\`);
  });

  if (fileNames.length >= 2) {
    console.log("\\n-> Gọi nhiều file đồng thời trong code:");
    console.log(\`   const data1 = files['\${fileNames[0]}'];\`);
    console.log(\`   const data2 = files['\${fileNames[1]}'];\`);
  }
}
`
  }
];
let activeScriptId = 1;

// ========================================================
// TRẠNG THÁI BẢNG TÍNH MÔ PHỎNG EXCEL (INLINE)
// ========================================================
let inlineExcelState = {
  fileName: "mock_data.xlsx",
  sheetName: "Sheet1",
  headers: ["A", "B", "C", "D"],
  rows: [
    ["DH1001", "Bàn phím cơ AKKO", "1", "1250000"],
    ["DH1002", "Chuột Logitech G102", "2", "450000"],
    ["DH1003", "Tai nghe Sony WH", "1", "2300000"],
    ["DH1004", "Lót chuột Gaming Size L", "3", "85000"]
  ]
};

let isUsingMockData = false;

function onInlineDataChanged() {
  if (isUsingMockData && typeof syncInlineExcelToDataState === 'function') {
    syncInlineExcelToDataState(false);
  }
}
