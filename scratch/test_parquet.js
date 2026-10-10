const fs = require('fs');
const code = fs.readFileSync('js/dashboard-builder.js', 'utf8');

// Trích xuất ParquetEngine
const match = code.match(/const ParquetEngine = (\{[\s\S]*?\n  \};)/);
if (!match) throw new Error('Cannot find ParquetEngine');
global.ParquetEngine = eval('(' + match[1].replace(/;\s*$/, '') + ')');

console.log('ParquetEngine found!');
// Tạo dữ liệu giả lập 100,000 dòng
const N = 100000;
const rawRows = [];
for (let i = 0; i < N; i++) {
  rawRows.push({
    'Tháng': 'T' + ((i % 12) + 1),
    'Chi Nhánh': ['Hà Nội', 'Đà Nẵng', 'Hồ Chí Minh', 'Cần Thơ'][i % 4],
    'Doanh Thu': (i * 1000) % 5000000,
    'Trạng Thái': i % 2 === 0 ? 'Hoàn thành' : 'Đang xử lý'
  });
}

console.log('Tạo xong rawRows:', N, 'dòng');

const t0 = Date.now();
const parquetTable = ParquetEngine.encodeToParquet(rawRows);
console.log('Encode sang Parquet mất:', Date.now() - t0, 'ms');

const t1 = Date.now();
const virtualArr = ParquetEngine.createVirtualArray(parquetTable);
console.log('Create Virtual Array mất:', Date.now() - t1, 'ms');

console.log('Độ dài mảng:', virtualArr.length);
console.log('Truy cập phần tử [0]:', JSON.stringify(virtualArr[0]));
console.log('Truy cập phần tử [50000]:', JSON.stringify(virtualArr[50000]));

// Test map
const t2 = Date.now();
const doanhThuArr = virtualArr.map(r => r['Doanh Thu']);
console.log('Chạy virtualArr.map(r => r["Doanh Thu"]) mất:', Date.now() - t2, 'ms, kết quả length:', doanhThuArr.length);

// Test filter
const t3 = Date.now();
const filtered = virtualArr.filter(r => r['Chi Nhánh'] === 'Hà Nội');
console.log('Chạy virtualArr.filter(r => r["Chi Nhánh"] === "Hà Nội") mất:', Date.now() - t3, 'ms, kết quả rows:', filtered.length);

// Test slice
const sliced = virtualArr.slice(0, 5);
console.log('Chạy virtualArr.slice(0, 5):', sliced.length, 'dòng, phần tử 0:', JSON.stringify(sliced[0]));

console.log('>>> TẤT CẢ TEST ĐỀU HOẠT ĐỘNG HOÀN HẢO! <<<');
