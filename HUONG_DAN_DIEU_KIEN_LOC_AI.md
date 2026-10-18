# TÀI LIỆU CẤU HÌNH ĐIỀU KIỆN & BỘ LỌC DỮ LIỆU (DÀNH CHO AI & NGƯỜI DÙNG)
> **Phần mềm:** Superset Dashboard & BI Studio (`dashboard_builder.html` / `js/dashboard-builder.js`)  
> **Nguyên lý cốt lõi:** **State-Driven (JSON Schema)** – Mọi điều kiện lọc đều lưu trong State JSON trước khi trigger `renderCanvas()`.

---

## 1. Cấu Trúc Đối Tượng Điều Kiện (Filter State Schema)

Khi người dùng hoặc AI thiết lập một điều kiện lọc trên một cột (`column`), dữ liệu điều kiện được lưu dưới dạng object `filterState`:

```json
{
  "op": "=",
  "val": "Hà Nội",
  "min": "",
  "max": "",
  "start": "",
  "end": "",
  "values": [],
  "isDate": false
}
```

### Bảng Toán Tử Hỗ Trợ (`op`)
| Nhóm | Toán tử (`op`) | Các trường dữ liệu cần truyền | Ví dụ minh họa |
| :--- | :--- | :--- | :--- |
| **Bằng / Khác** | `=`, `!=` | `val` (số hoặc chuỗi) | `val: "Miền Bắc"` hoặc `val: 100` |
| **So sánh số** | `>`, `>=`, `<`, `<=` | `val` (số hoặc chuỗi) | `op: ">=", val: 50000000` |
| **Khoảng số (Range)** | `between`, `range` | `min`, `max` (số hoặc chuỗi) | `op: "between", min: 10, max: 100` |
| **Khoảng ngày** | `date_range` | `start`, `end` (`YYYY-MM-DD` hoặc `DD/MM/YYYY`), `isDate: true` | `start: "2026-01-01", end: "2026-03-31"` |
| **Ngày đơn** | `date_eq`, `date_gte`, `date_lte` | `val`, `isDate: true` | `op: "date_gte", val: "2026-01-01"` |
| **Chọn nhiều mục** | `in` | `values` (mảng danh sách các giá trị) | `values: ["Hà Nội", "Đà Nẵng", "TP.HCM"]` |
| **Chuỗi / Văn bản** | `contains`, `starts_with`, `ends_with`, `not_contains` | `val` (chuỗi tìm kiếm, không phân biệt hoa thường) | `op: "contains", val: "VIP"` |

---

## 2. Nơi Lưu Trữ Trạng Thái Điều Kiện Trong Hệ Thống

1. **Bộ lọc theo Tab (`tabFilters`):**
   ```javascript
   DashboardState.present.tabFilters[tabId][column] = filterState;
   // Ví dụ: DashboardState.present.tabFilters['tab_1']['KhuVuc'] = { op: '=', val: 'Hà Nội' };
   ```
2. **Bộ lọc tổng hợp toàn cục (`activeFilters`):**
   ```javascript
   DashboardState.activeFilters[column] = filterState;
   ```
3. **Cấu hình trên Widget Bộ Lọc (`item`):**
   ```json
   {
     "id": "filter_1718000000",
     "type": "filter",
     "name": "Lọc Theo Khu Vực",
     "dataSourceKey": "ds_0",
     "filterColumn": "KhuVuc",
     "filterType": "select",
     "operator": "=",
     "tabScope": "current",
     "parentFilterId": null,
     "currentFilterValue": "all"
   }
   ```
   * `filterType`: `'select'` (chọn 1 mục), `'multiselect'` (chọn nhiều mục), `'range'` (khoảng số/ngày), `'comparison'` (so sánh 1 mốc), `'text'` (tìm kiếm từ khóa).
   * `tabScope`: `'current'` (chỉ lọc tab hiện tại) hoặc `'all'` (áp dụng cho toàn bộ các tab).
   * `parentFilterId`: ID của bộ lọc cha (nếu muốn lọc phân cấp: chọn Tỉnh -> mới hiện Huyện).

---

## 3. Cách Hệ Thống Tính Toán Dữ Liệu Sau Khi Áp Điều Kiện

Mọi biểu đồ ECharts, Thẻ KPI, Bảng dữ liệu Superset đều lấy dữ liệu đã lọc qua hàm:

```javascript
const filteredRows = getFilteredData(sourceKey, excludeCols, tabId, chartId);
```

* **Cơ chế lọc tự động:**
  1. Lấy toàn bộ hàng gốc từ `DataStore.getData(sourceKey)`.
  2. Gom tất cả bộ lọc hợp lệ từ Tab và Global.
  3. Lọc qua từng dòng: Nếu dòng thỏa mãn đồng thời tất cả các cột điều kiện (logic `AND`) thì giữ lại.
  4. Trả về mảng `filteredRows` để ECharts vẽ lại hoặc KPI tính tổng `SUM/AVG/COUNT`.

---

## 4. Hướng Dẫn Dành Cho AI Khi Thao Tác Bằng Code

Khi AI cần tự động tạo hoặc thay đổi điều kiện lọc, **BẮT BUỘC** tuân thủ các bước:

```javascript
// BƯỚC 1: Lưu snapshot để hỗ trợ Undo/Redo
DashboardState.pushSnapshot();

// BƯỚC 2: Cập nhật điều kiện vào State (Tuyệt đối không sửa DOM trực tiếp)
const curTab = DashboardState.present.activeTabId || 'default';
if (!DashboardState.present.tabFilters[curTab]) {
  DashboardState.present.tabFilters[curTab] = {};
}

// Thiết lập điều kiện: Lọc DoanhThu >= 50,000,000
DashboardState.present.tabFilters[curTab]['DoanhThu'] = {
  op: '>=',
  val: 50000000
};

// Đồng bộ sang activeFilters của phiên hiện tại
DashboardState.activeFilters['DoanhThu'] = DashboardState.present.tabFilters[curTab]['DoanhThu'];

// BƯỚC 3: Đánh dấu dữ liệu đã thay đổi và kích hoạt render
DashboardState.isDirty = true;
renderCanvas();
triggerDebouncedAutoSave();
```

---

## 5. Hướng Dẫn Nhanh Dành Cho Người Dùng (Giao Diện)

1. **Thêm bộ lọc mới:**
   * Ở thanh công cụ nổi phía dưới, nhấn **`➕ Thêm Thành Phần`** ➔ Chọn **`🔍 Bộ Lọc (Filter)`**.
2. **Cấu hình điều kiện lọc:**
   * **Bước 1:** Chọn **File / Sheet nguồn** chứa dữ liệu cần lọc.
   * **Bước 2:** Chọn **Cột dữ liệu** (hệ thống sẽ tự nhận diện loại Dữ liệu: Chữ, Số hay Ngày).
   * **Bước 3:** Chọn **Kiểu bộ lọc**:
     * *Khoảng Ngày (Date Picker):* Lọc từ ngày ... đến ngày ...
     * *Chọn Nhiều Mục (Checkbox Dropdown):* Tích chọn nhiều nhóm cùng lúc.
     * *Chọn 1 Mục (Single Select):* Thả xuống chọn 1 giá trị duy nhất.
     * *Khoảng Số (Min - Max):* Nhập khoảng doanh thu, số lượng...
     * *So Sánh 1 Mốc:* Chọn toán tử `>`, `>=`, `<`, `<=`, `≠`.
     * *Tìm Kiếm Chữ:* Gõ từ khóa tìm kiếm (chứa từ, bắt đầu bằng...).
   * **Bước 4:** Chọn phạm vi áp dụng (Tab hiện tại hay Tất cả các Tab).
   * **Bước 5:** Nhấn **`💾 Lưu Bộ Lọc`**.
3. **Sử dụng:**
   * Thao tác trực tiếp trên widget Bộ Lọc vừa tạo trên màn hình thiết kế. Ngay khi bạn chọn hoặc nhập giá trị, toàn bộ Biểu đồ ECharts, Thẻ KPI và Bảng trong trang sẽ lập tức tự động tính toán lại theo đúng điều kiện đã chọn.
