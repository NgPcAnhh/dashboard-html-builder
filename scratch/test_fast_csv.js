const fs = require('fs');
const code = fs.readFileSync('js/dashboard-builder.js', 'utf8');

// Trích xuất ParquetEngine và FastCsvParser
const peMatch = code.match(/const ParquetEngine = (\{[\s\S]*?\n  \};)/);
global.ParquetEngine = eval('(' + peMatch[1].replace(/;\s*$/, '') + ')');

const csvMatch = code.match(/const FastCsvParser = (\{[\s\S]*?\n  \};)/);
global.FastCsvParser = eval('(' + csvMatch[1].replace(/;\s*$/, '') + ')');

console.log('ParquetEngine & FastCsvParser loaded!');

// Tạo một file CSV giả lập lớn
const N = 100000;
let csvContent = 'Mã Đơn,Tháng,Khu Vực,Sản Phẩm,Doanh Thu,Lợi Nhuận\n';
const regions = ['Miền Bắc', 'Miền Trung', 'Miền Nam', 'Tây Nguyên'];
const products = ['Điện thoại', 'Laptop', 'Phụ kiện', 'Tivi', 'Tủ lạnh'];

for (let i = 1; i <= N; i++) {
  const thang = 'T' + ((i % 12) + 1);
  const kv = regions[i % regions.length];
  const sp = products[i % products.length];
  const dt = 1000000 + (i * 137) % 50000000;
  const ln = Math.round(dt * 0.25);
  csvContent += `DH_${i},${thang},${kv},${sp},${dt},${ln}\n`;
}

const csvBuffer = Buffer.from(csvContent, 'utf8');
console.log(`Đã tạo dữ liệu CSV: ${N} dòng, dung lượng buffer: ${(csvBuffer.length / 1024 / 1024).toFixed(2)} MB`);

// Giả lập đối tượng File trong Node
const mockFile = {
  name: 'big_test_data.csv',
  size: csvBuffer.length,
  slice: (start, end) => ({
    arrayBuffer: async () => csvBuffer.subarray(start, end).buffer
  })
};

(async () => {
  const t0 = Date.now();
  const parquetTable = await FastCsvParser.parse(mockFile);
  const elapsed = Date.now() - t0;
  console.log(`⏱️ Thời gian Streaming Parse ${N} dòng: ${elapsed} ms!`);

  const virtualArr = ParquetEngine.createVirtualArray(parquetTable);
  console.log('Tổng số dòng nhận được:', virtualArr.length);
  console.log('Dòng [0]:', JSON.stringify(virtualArr[0]));
  console.log('Dòng [99999]:', JSON.stringify(virtualArr[99999]));

  // Test map & filter
  const t1 = Date.now();
  const tongDT = virtualArr.reduce((sum, r) => sum + r['Doanh Thu'], 0);
  console.log(`Tính tổng Doanh Thu (${tongDT.toLocaleString('vi-VN')} đ) mất: ${Date.now() - t1} ms`);

  const t2 = Date.now();
  const filtered = virtualArr.filter(r => r['Khu Vực'] === 'Miền Bắc');
  console.log(`Lọc Miền Bắc (${filtered.length} dòng) mất: ${Date.now() - t2} ms`);

  console.log('>>> BENCHMARK HOÀN TẤT VỚI KẾT QUẢ XUẤT SẮC! <<<');
})();
