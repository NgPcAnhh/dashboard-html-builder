// ========================================================
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
  console.log(`Đã nạp ${fileNames.length} file:`);
  fileNames.forEach((name, i) => {
    const fileData = files[name];
    const sheetList = fileData.sheets ? Object.keys(fileData.sheets) : [];
    console.log(`  ${i + 1}. [${name}] - Tổng ${fileData.length} dòng (Sheets: ${sheetList.join(', ') || 'Mặc định'})`);
  });

  if (fileNames.length >= 2) {
    console.log("\n-> Gọi nhiều file đồng thời trong code:");
    console.log(`   const data1 = files['${fileNames[0]}'];`);
    console.log(`   const data2 = files['${fileNames[1]}'];`);
  }
}
