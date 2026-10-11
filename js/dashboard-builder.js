/**
 * =============================================================================
 * SUPERSET DASHBOARD & BI STUDIO - JAVASCRIPT ENGINE
 * Modeled after BI Platform (d:\Project\data-analytic-platform)
 * - Pure JavaScript ECharts code engine for every chart
 * - Multi-Excel / Multi-Sheet data integration via SheetJS
 * - AI System Prompt Studio for Low-Code / No-Code chart generation
 * - No-Database JSON configuration persistence (Load, Edit, Re-save)
 * - 100% Identical Standalone HTML Viewer Export pointing to local Excel files
 * - Canvas Studio: 12-col Grid Drag & Resize, Tabs, Frames, Filter Widgets, KPI
 * =============================================================================
 */

(function () {
  'use strict';

  // ===========================================================================
  // 1. DATA SOURCE STORE (MULTI-EXCEL / MULTI-SHEET)
  // ===========================================================================
  // ===========================================================================
  // 0.8. PARQUET COLUMNAR ENGINE (COMPACT COLUMNAR STORAGE FOR LARGE DATASETS >= 50MB)
  // ===========================================================================
  const PARQUET_THRESHOLD_BYTES = 50 * 1024 * 1024; // Ngưỡng kích hoạt nén Parquet mặc định: 50MB

  const ParquetEngine = {
    isParquet(obj) {
      return Boolean(obj && (obj.__isParquet || (obj.__table && obj.__table.__isParquet)));
    },

    encodeToParquet(rows, metadata = {}) {
      if (!Array.isArray(rows) || rows.length === 0) return rows;
      if (this.isParquet(rows)) return rows;

      const colNames = Object.keys(rows[0] || {});
      const rowCount = rows.length;
      const columns = {};
      const columnTypes = {};

      for (let c = 0; c < colNames.length; c++) {
        const col = colNames[c];
        let isAllNumeric = true;
        let hasValue = false;
        const uniqueMap = new Map();
        let uniqueCounter = 0;
        const indices = new Array(rowCount);
        const rawNumbers = new Array(rowCount);

        for (let i = 0; i < rowCount; i++) {
          const val = rows[i][col];
          if (val !== null && val !== undefined && val !== '') {
            hasValue = true;
            const num = Number(val);
            if (isNaN(num)) {
              isAllNumeric = false;
            } else {
              rawNumbers[i] = num;
            }
          } else {
            rawNumbers[i] = 0;
          }

          const strKey = String(val !== undefined && val !== null ? val : '');
          if (!uniqueMap.has(strKey)) {
            uniqueMap.set(strKey, uniqueCounter++);
          }
          indices[i] = uniqueMap.get(strKey);
        }

        // Nếu là cột số hoàn toàn
        if (isAllNumeric && hasValue) {
          columnTypes[col] = 'numeric';
          columns[col] = {
            type: 'numeric',
            data: rawNumbers
          };
        } else {
          // Chuẩn Parquet Dictionary Encoding
          const dictArray = new Array(uniqueMap.size);
          uniqueMap.forEach((idx, key) => {
            dictArray[idx] = key;
          });
          columnTypes[col] = 'dictionary';
          columns[col] = {
            type: 'dictionary',
            dictionary: dictArray,
            indices: indices
          };
        }
      }

      const parquetTable = {
        __isParquet: true,
        rowCount,
        colNames,
        columnTypes,
        columns,
        metadata: {
          ...metadata,
          encodedAt: new Date().toISOString()
        },
        getValue(col, rowIdx) {
          const colData = this.columns[col];
          if (!colData) return undefined;
          if (colData.type === 'numeric') {
            return colData.data[rowIdx];
          }
          const dictIdx = colData.indices[rowIdx];
          return colData.dictionary[dictIdx];
        },
        hasColumn(col) {
          return Boolean(this.columns[col]);
        },
        extractRowObject(rowIdx) {
          const rowObj = {};
          for (let i = 0; i < this.colNames.length; i++) {
            const col = this.colNames[i];
            rowObj[col] = this.getValue(col, rowIdx);
          }
          return rowObj;
        }
      };

      return parquetTable;
    },

    createVirtualArray(table) {
      if (!table) return [];
      if (!table.__isParquet && Array.isArray(table)) return table;

      const tbl = table;
      const len = tbl.rowCount;

      // Xây dựng Virtual Array class
      class VirtualArray extends Array {
        constructor() {
          super(len);
          Object.defineProperty(this, '__table', { value: tbl, enumerable: false, writable: true });
          Object.defineProperty(this, '__isParquet', { value: true, enumerable: false });
          Object.defineProperty(this, '__columns', { value: tbl.colNames, enumerable: false });
          Object.defineProperty(this, 'length', { value: len, enumerable: false, writable: false });
        }

        map(callback, thisArg) {
          const res = new Array(len);
          let curIdx = 0;
          const flyweight = new Proxy({}, {
            get(_, prop) {
              if (prop === '__rawIndex') return curIdx;
              return tbl.getValue(prop, curIdx);
            },
            has(_, prop) {
              return tbl.hasColumn(prop);
            },
            ownKeys() {
              return tbl.colNames;
            },
            getOwnPropertyDescriptor() {
              return { enumerable: true, configurable: true };
            }
          });

          for (let i = 0; i < len; i++) {
            curIdx = i;
            res[i] = callback.call(thisArg, flyweight, i, this);
          }
          return res;
        }

        forEach(callback, thisArg) {
          let curIdx = 0;
          const flyweight = new Proxy({}, {
            get(_, prop) {
              if (prop === '__rawIndex') return curIdx;
              return tbl.getValue(prop, curIdx);
            },
            has(_, prop) {
              return tbl.hasColumn(prop);
            },
            ownKeys() {
              return tbl.colNames;
            },
            getOwnPropertyDescriptor() {
              return { enumerable: true, configurable: true };
            }
          });

          for (let i = 0; i < len; i++) {
            curIdx = i;
            callback.call(thisArg, flyweight, i, this);
          }
        }

        filter(callback, thisArg) {
          const res = [];
          let curIdx = 0;
          const flyweight = new Proxy({}, {
            get(_, prop) {
              if (prop === '__rawIndex') return curIdx;
              return tbl.getValue(prop, curIdx);
            },
            has(_, prop) {
              return tbl.hasColumn(prop);
            },
            ownKeys() {
              return tbl.colNames;
            },
            getOwnPropertyDescriptor() {
              return { enumerable: true, configurable: true };
            }
          });

          for (let i = 0; i < len; i++) {
            curIdx = i;
            if (callback.call(thisArg, flyweight, i, this)) {
              res.push(tbl.extractRowObject(i));
            }
          }
          return res;
        }

        slice(start = 0, end = len) {
          const actualStart = Math.max(0, start < 0 ? len + start : start);
          const actualEnd = Math.min(len, end < 0 ? len + end : end);
          const res = [];
          for (let i = actualStart; i < actualEnd; i++) {
            res.push(tbl.extractRowObject(i));
          }
          return res;
        }

        reduce(callback, initialValue) {
          let accumulator = initialValue;
          let startIndex = 0;
          if (accumulator === undefined) {
            if (len === 0) throw new TypeError('Reduce of empty array with no initial value');
            accumulator = tbl.extractRowObject(0);
            startIndex = 1;
          }
          let curIdx = 0;
          const flyweight = new Proxy({}, {
            get(_, prop) {
              return tbl.getValue(prop, curIdx);
            }
          });
          for (let i = startIndex; i < len; i++) {
            curIdx = i;
            accumulator = callback(accumulator, flyweight, i, this);
          }
          return accumulator;
        }

        toRawArray() {
          const res = new Array(len);
          for (let i = 0; i < len; i++) {
            res[i] = tbl.extractRowObject(i);
          }
          return res;
        }
      }

      const instance = new VirtualArray();

      return new Proxy(instance, {
        get(target, prop) {
          if (typeof prop === 'string') {
            const idx = Number(prop);
            if (!isNaN(idx) && Number.isInteger(idx) && idx >= 0 && idx < len) {
              return tbl.extractRowObject(idx);
            }
          }
          return target[prop];
        }
      });
    }
  };

  // ===========================================================================
  // 0.9. FAST STREAMING CSV PARSER (CHUNK-TO-COLUMNAR, ZERO MAIN-THREAD CRASH)
  // ===========================================================================
  const FastCsvParser = {
    async parse(file, options = {}) {
      const chunkSize = 2 * 1024 * 1024; // 2MB chunk
      const fileSize = file.size;
      let offset = 0;
      let leftover = '';
      let isFirstChunk = true;
      let colNames = [];
      let colCount = 0;
      let delimiter = ',';
      let columnsData = null;

      const decoder = new TextDecoder(options.encoding || 'utf-8');

      while (offset < fileSize) {
        const slice = file.slice(offset, Math.min(fileSize, offset + chunkSize));
        const buffer = await slice.arrayBuffer();
        const textChunk = leftover + decoder.decode(buffer, { stream: offset + chunkSize < fileSize });

        const lineEnd = textChunk.lastIndexOf('\n');
        if (lineEnd === -1) {
          leftover = textChunk;
          offset += chunkSize;
          continue;
        }

        const processText = textChunk.substring(0, lineEnd);
        leftover = textChunk.substring(lineEnd + 1);
        const lines = processText.split(/\r?\n/);

        for (let i = 0; i < lines.length; i++) {
          const line = lines[i].trim();
          if (!line) continue;

          if (isFirstChunk && colNames.length === 0) {
            const commaCount = (line.match(/,/g) || []).length;
            const semiCount = (line.match(/;/g) || []).length;
            const tabCount = (line.match(/\t/g) || []).length;
            if (semiCount > commaCount && semiCount > tabCount) delimiter = ';';
            else if (tabCount > commaCount && tabCount > semiCount) delimiter = '\t';
            else delimiter = ',';

            colNames = FastCsvParser.parseCsvLine(line, delimiter).map(c => c.trim().replace(/^["']|["']$/g, ''));
            colCount = colNames.length;
            columnsData = Array.from({ length: colCount }, () => []);
            continue;
          }

          const cells = FastCsvParser.parseCsvLine(line, delimiter);
          for (let c = 0; c < colCount; c++) {
            const raw = cells[c] !== undefined ? cells[c].trim().replace(/^["']|["']$/g, '') : '';
            columnsData[c].push(raw);
          }
        }

        offset += chunkSize;
        if (options.onProgress) {
          options.onProgress(Math.min(99, Math.round((offset / fileSize) * 100)));
        }
      }

      if (leftover && leftover.trim() && colNames.length > 0) {
        const cells = FastCsvParser.parseCsvLine(leftover.trim(), delimiter);
        for (let c = 0; c < colCount; c++) {
          const raw = cells[c] !== undefined ? cells[c].trim().replace(/^["']|["']$/g, '') : '';
          columnsData[c].push(raw);
        }
      }

      const rowCount = columnsData && columnsData[0] ? columnsData[0].length : 0;
      const columns = {};
      const columnTypes = {};

      for (let c = 0; c < colCount; c++) {
        const col = colNames[c];
        const rawCol = columnsData[c];
        let isAllNumeric = true;
        let hasValue = false;
        const uniqueMap = new Map();
        let uniqueCounter = 0;
        const indices = new Array(rowCount);
        const rawNumbers = new Array(rowCount);

        for (let r = 0; r < rowCount; r++) {
          const val = rawCol[r];
          if (val !== '') {
            hasValue = true;
            const num = Number(val);
            if (isNaN(num)) {
              isAllNumeric = false;
            } else {
              rawNumbers[r] = num;
            }
          } else {
            rawNumbers[r] = 0;
          }

          if (!uniqueMap.has(val)) {
            uniqueMap.set(val, uniqueCounter++);
          }
          indices[r] = uniqueMap.get(val);
        }

        if (isAllNumeric && hasValue) {
          columnTypes[col] = 'numeric';
          columns[col] = { type: 'numeric', data: rawNumbers };
        } else {
          const dictArray = new Array(uniqueMap.size);
          uniqueMap.forEach((idx, key) => { dictArray[idx] = key; });
          columnTypes[col] = 'dictionary';
          columns[col] = { type: 'dictionary', dictionary: dictArray, indices };
        }
      }

      return {
        __isParquet: true,
        rowCount,
        colNames,
        columnTypes,
        columns,
        metadata: {
          fileName: file.name,
          encodedAt: new Date().toISOString()
        },
        getValue(col, rowIdx) {
          const colData = this.columns[col];
          if (!colData) return undefined;
          if (colData.type === 'numeric') return colData.data[rowIdx];
          const dictIdx = colData.indices[rowIdx];
          return colData.dictionary[dictIdx];
        },
        hasColumn(col) {
          return Boolean(this.columns[col]);
        },
        extractRowObject(rowIdx) {
          const rowObj = {};
          for (let i = 0; i < this.colNames.length; i++) {
            const col = this.colNames[i];
            rowObj[col] = this.getValue(col, rowIdx);
          }
          return rowObj;
        }
      };
    },

    parseCsvLine(text, delimiter = ',') {
      const res = [];
      let cur = '';
      let inQuotes = false;
      for (let i = 0; i < text.length; i++) {
        const ch = text[i];
        if (ch === '"') {
          inQuotes = !inQuotes;
        } else if (ch === delimiter && !inQuotes) {
          res.push(cur);
          cur = '';
        } else {
          cur += ch;
        }
      }
      res.push(cur);
      return res;
    }
  };

  const DEFAULT_SAMPLE_DATA = [
    { 'Ngày': '2026-01-15', 'Tháng': 'T1/2026', 'Khu Vực': 'Miền Bắc', 'Nhóm Sản Phẩm': 'Điện tử', 'Doanh Thu': 125000000, 'Lợi Nhuận': 28000000, 'Chi Phí': 97000000, 'Số Đơn Hàng': 450 },
    { 'Ngày': '2026-02-18', 'Tháng': 'T2/2026', 'Khu Vực': 'Miền Bắc', 'Nhóm Sản Phẩm': 'Gia dụng', 'Doanh Thu': 98000000, 'Lợi Nhuận': 19500000, 'Chi Phí': 78500000, 'Số Đơn Hàng': 380 },
    { 'Ngày': '2026-03-22', 'Tháng': 'T3/2026', 'Khu Vực': 'Miền Nam', 'Nhóm Sản Phẩm': 'Điện tử', 'Doanh Thu': 165000000, 'Lợi Nhuận': 39000000, 'Chi Phí': 126000000, 'Số Đơn Hàng': 610 },
    { 'Ngày': '2026-04-10', 'Tháng': 'T4/2026', 'Khu Vực': 'Miền Nam', 'Nhóm Sản Phẩm': 'Thời trang', 'Doanh Thu': 112000000, 'Lợi Nhuận': 34000000, 'Chi Phí': 78000000, 'Số Đơn Hàng': 520 },
    { 'Ngày': '2026-05-25', 'Tháng': 'T5/2026', 'Khu Vực': 'Miền Trung', 'Nhóm Sản Phẩm': 'Gia dụng', 'Doanh Thu': 87000000, 'Lợi Nhuận': 16000000, 'Chi Phí': 71000000, 'Số Đơn Hàng': 290 },
    { 'Ngày': '2026-06-30', 'Tháng': 'T6/2026', 'Khu Vực': 'Miền Bắc', 'Nhóm Sản Phẩm': 'Thời trang', 'Doanh Thu': 142000000, 'Lợi Nhuận': 41000000, 'Chi Phí': 101000000, 'Số Đơn Hàng': 580 },
    { 'Ngày': '2026-07-14', 'Tháng': 'T7/2026', 'Khu Vực': 'Miền Nam', 'Nhóm Sản Phẩm': 'Điện tử', 'Doanh Thu': 195000000, 'Lợi Nhuận': 48000000, 'Chi Phí': 147000000, 'Số Đơn Hàng': 720 },
    { 'Ngày': '2026-08-08', 'Tháng': 'T8/2026', 'Khu Vực': 'Miền Trung', 'Nhóm Sản Phẩm': 'Điện tử', 'Doanh Thu': 105000000, 'Lợi Nhuận': 22000000, 'Chi Phí': 83000000, 'Số Đơn Hàng': 360 }
  ];

  const DataStore = {
    files: {}, // filename -> { name, rawBytes, sheets: { sheetName: rows[] } }

    initDefault() {
      const fileName = 'Du_Lieu_Mau_Kinh_Doanh.xlsx';
      const sheetName = 'BaoCaoDoanhSo';
      this.files[fileName] = {
        name: fileName,
        sheets: {
          [sheetName]: JSON.parse(JSON.stringify(DEFAULT_SAMPLE_DATA))
        },
        rawBytes: null
      };
    },

    getAvailableSources() {
      const list = [];
      for (const fileName in this.files) {
        const fileObj = this.files[fileName];
        for (const sheetName in fileObj.sheets) {
          const sheet = fileObj.sheets[sheetName];
          const count = sheet ? (sheet.length !== undefined ? sheet.length : (sheet.rowCount || 0)) : 0;
          list.push({
            fileName,
            sheetName,
            key: `${fileName}::${sheetName}`,
            label: `${fileName} » ${sheetName} (${count.toLocaleString('vi-VN')} dòng)`
          });
        }
      }
      return list;
    },

    getData(fileKey, fallbackCol = null) {
      if (!fileKey && fallbackCol) {
        const sources = this.getAvailableSources();
        const found = sources.find(s => {
          const raw = (this.files[s.fileName]?.sheets?.[s.sheetName]);
          if (!raw) return false;
          if (raw.__columns && raw.__columns.includes(fallbackCol)) return true;
          if (Array.isArray(raw) && raw[0] && raw[0][fallbackCol] !== undefined) return true;
          return false;
        });
        if (found) fileKey = found.key;
      }
      if (!fileKey) {
        const sources = this.getAvailableSources();
        if (sources.length > 0) fileKey = sources[0].key;
      }
      if (!fileKey) return [];

      const parts = String(fileKey).split('::');
      const fileName = parts[0];
      const sheetName = parts[1];

      // Trường hợp 1: fileKey dạng fileName::sheetName
      if (fileName && sheetName && this.files[fileName]?.sheets?.[sheetName]) {
        const rawSheet = this.files[fileName].sheets[sheetName];
        if (rawSheet.__isParquet && !rawSheet.__table) {
          this.files[fileName].sheets[sheetName] = ParquetEngine.createVirtualArray(rawSheet);
          return this.files[fileName].sheets[sheetName];
        }
        return rawSheet;
      }

      // Trường hợp 2: fileKey chỉ là fileName (lấy sheet đầu tiên)
      if (this.files[fileName]?.sheets) {
        const sKeys = Object.keys(this.files[fileName].sheets);
        if (sKeys.length > 0) {
          const rawSheet = this.files[fileName].sheets[sKeys[0]];
          if (rawSheet.__isParquet && !rawSheet.__table) {
            this.files[fileName].sheets[sKeys[0]] = ParquetEngine.createVirtualArray(rawSheet);
            return this.files[fileName].sheets[sKeys[0]];
          }
          return rawSheet;
        }
      }

      // Trường hợp 3: fileKey là sheetName nằm trong file bất kỳ
      for (const fn in this.files) {
        if (this.files[fn]?.sheets?.[fileKey]) {
          const rawSheet = this.files[fn].sheets[fileKey];
          if (rawSheet.__isParquet && !rawSheet.__table) {
            this.files[fn].sheets[fileKey] = ParquetEngine.createVirtualArray(rawSheet);
            return this.files[fn].sheets[fileKey];
          }
          return rawSheet;
        }
      }

      return [];
    },

    getAllFilesObject() {
      const obj = {};
      for (const fileName in this.files) {
        obj[fileName] = this.files[fileName].sheets;
      }
      return obj;
    },

    compactRows(rows) {
      if (!rows) return rows;
      if (ParquetEngine.isParquet(rows)) {
        const tbl = rows.__table || rows;
        return {
          __isParquet: true,
          rowCount: tbl.rowCount,
          colNames: tbl.colNames,
          columnTypes: tbl.columnTypes,
          columns: tbl.columns,
          metadata: tbl.metadata
        };
      }
      if (!Array.isArray(rows) || rows.length === 0) return rows;
      if (typeof rows[0] !== 'object' || rows[0] === null || Array.isArray(rows[0])) return rows;
      const cols = Object.keys(rows[0]);
      if (cols.length === 0) return rows;
      const data = rows.map(r => cols.map(c => (r[c] !== undefined && r[c] !== null) ? r[c] : ''));
      return { __c: true, cols, data };
    },

    uncompactRows(sheetData) {
      if (!sheetData) return [];
      if (sheetData.__isParquet) {
        return ParquetEngine.createVirtualArray(sheetData);
      }
      if (Array.isArray(sheetData)) return sheetData;
      if (sheetData.__c && Array.isArray(sheetData.cols) && Array.isArray(sheetData.data)) {
        const cols = sheetData.cols;
        const len = cols.length;
        return sheetData.data.map(rowArr => {
          const rowObj = {};
          for (let i = 0; i < len; i++) {
            rowObj[cols[i]] = rowArr[i];
          }
          return rowObj;
        });
      }
      return sheetData;
    },

    getAllFilesObjectCompact() {
      const obj = {};
      for (const fileName in this.files) {
        obj[fileName] = {};
        const sheets = this.files[fileName].sheets || {};
        for (const sheetName in sheets) {
          obj[fileName][sheetName] = this.compactRows(sheets[sheetName]);
        }
      }
      return obj;
    },

    getColumns(fileKey) {
      const rows = this.getData(fileKey);
      if (!rows || rows.length === 0) return [];
      if (rows.__columns) return rows.__columns;
      if (rows[0] && typeof rows[0] === 'object') return Object.keys(rows[0]);
      return [];
    }
  };

  // ===========================================================================
  // 1.5. UNIVERSAL SYSTEM DIALOG POPUP (REPLACES NATIVE ALERT / CONFIRM / PROMPT)
  // ===========================================================================
  const AppDialog = {
    modalEl: null,
    resolveFn: null,

    init() {
      if (this._initialized) return;
      this._initialized = true;
      this.modalEl = document.getElementById('appDialogModal');
      if (!this.modalEl) return;

      document.getElementById('appDialogCloseBtn')?.addEventListener('click', () => this.close(null));
      document.getElementById('appDialogCancelBtn')?.addEventListener('click', () => this.close(false));
      document.getElementById('appDialogConfirmBtn')?.addEventListener('click', () => {
        const promptWrap = document.getElementById('appDialogPromptWrap');
        if (promptWrap && promptWrap.style.display !== 'none') {
          const input = document.getElementById('appDialogInput');
          this.close(input ? input.value : '');
        } else {
          this.close(true);
        }
      });

      document.getElementById('appDialogInput')?.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
          e.preventDefault();
          const input = document.getElementById('appDialogInput');
          this.close(input ? input.value : '');
        } else if (e.key === 'Escape') {
          e.preventDefault();
          this.close(null);
        }
      });

      window.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && this.modalEl && this.modalEl.style.display !== 'none') {
          this.close(null);
        }
      });

      // Override window alert safely in page context
      window.alert = (msg) => { this.alert(String(msg)); };
    },

    close(result) {
      if (!this.modalEl) return;
      this.modalEl.style.display = 'none';
      if (typeof this.resolveFn === 'function') {
        const fn = this.resolveFn;
        this.resolveFn = null;
        fn(result);
      }
    },

    alert(message, title = 'Thông Báo') {
      if (!this.modalEl) this.init();
      return new Promise((resolve) => {
        this.resolveFn = resolve;
        const icon = document.getElementById('appDialogIcon');
        const titleEl = document.getElementById('appDialogTitle');
        const msgEl = document.getElementById('appDialogMessage');
        const promptWrap = document.getElementById('appDialogPromptWrap');
        const cancelBtn = document.getElementById('appDialogCancelBtn');
        const confirmBtn = document.getElementById('appDialogConfirmBtn');

        if (icon) icon.textContent = 'ℹ️';
        if (titleEl) titleEl.textContent = title;
        if (msgEl) msgEl.textContent = message;
        if (promptWrap) promptWrap.style.display = 'none';
        if (cancelBtn) cancelBtn.style.display = 'none';
        if (confirmBtn) {
          confirmBtn.textContent = 'Đồng Ý';
          confirmBtn.classList.remove('is-danger');
        }

        if (this.modalEl) this.modalEl.style.display = 'flex';
        confirmBtn?.focus();
      });
    },

    confirm(message, title = 'Xác Nhận', options = {}) {
      if (!this.modalEl) this.init();
      const { confirmText = 'Xác Nhận', cancelText = 'Hủy Bỏ', isDanger = false } = options;
      return new Promise((resolve) => {
        this.resolveFn = resolve;
        const icon = document.getElementById('appDialogIcon');
        const titleEl = document.getElementById('appDialogTitle');
        const msgEl = document.getElementById('appDialogMessage');
        const promptWrap = document.getElementById('appDialogPromptWrap');
        const cancelBtn = document.getElementById('appDialogCancelBtn');
        const confirmBtn = document.getElementById('appDialogConfirmBtn');

        if (icon) icon.textContent = isDanger ? '⚠️' : '❓';
        if (titleEl) titleEl.textContent = title;
        if (msgEl) msgEl.textContent = message;
        if (promptWrap) promptWrap.style.display = 'none';
        if (cancelBtn) {
          cancelBtn.style.display = 'inline-flex';
          cancelBtn.textContent = cancelText;
        }
        if (confirmBtn) {
          confirmBtn.textContent = confirmText;
          confirmBtn.classList.toggle('is-danger', Boolean(isDanger));
        }

        if (this.modalEl) this.modalEl.style.display = 'flex';
        confirmBtn?.focus();
      });
    },

    prompt(message, defaultValue = '', title = 'Nhập Thông Tin') {
      if (!this.modalEl) this.init();
      return new Promise((resolve) => {
        this.resolveFn = resolve;
        const icon = document.getElementById('appDialogIcon');
        const titleEl = document.getElementById('appDialogTitle');
        const msgEl = document.getElementById('appDialogMessage');
        const promptWrap = document.getElementById('appDialogPromptWrap');
        const input = document.getElementById('appDialogInput');
        const cancelBtn = document.getElementById('appDialogCancelBtn');
        const confirmBtn = document.getElementById('appDialogConfirmBtn');

        if (icon) icon.textContent = '✏️';
        if (titleEl) titleEl.textContent = title;
        if (msgEl) msgEl.textContent = message;
        if (promptWrap) promptWrap.style.display = 'block';
        if (input) {
          input.value = defaultValue || '';
          setTimeout(() => {
            input.focus();
            input.select();
          }, 60);
        }
        if (cancelBtn) {
          cancelBtn.style.display = 'inline-flex';
          cancelBtn.textContent = 'Hủy Bỏ';
        }
        if (confirmBtn) {
          confirmBtn.textContent = 'Xác Nhận';
          confirmBtn.classList.remove('is-danger');
        }

        if (this.modalEl) this.modalEl.style.display = 'flex';
      });
    }
  };

  // ===========================================================================
  // 1.9 SYNC DOM POSITIONS TO STATE ITEMS (CHỐNG MẤT TOẠ ĐỘ THỰC TẾ KHI THAO TÁC)
  // ===========================================================================
  function syncItemPositionsFromDOM() {
    if (!DashboardState.present || !Array.isArray(DashboardState.present.items)) return;
    const canvas = document.getElementById('dashboardCanvas');
    if (!canvas) return;

    DashboardState.present.items.forEach(item => {
      const card = document.getElementById(`item_${item.id}`);
      if (!card) return;

      const leftVal = parseInt(card.style.left, 10);
      const topVal = parseInt(card.style.top, 10);
      const widthVal = parseInt(card.style.width, 10);
      const heightVal = parseInt(card.style.height, 10);

      if (!item.pos) item.pos = {};

      if (!isNaN(leftVal)) {
        item.pos.x = leftVal;
      } else if (card.offsetLeft !== undefined && card.offsetLeft > 0) {
        item.pos.x = card.offsetLeft;
      }

      if (!isNaN(topVal)) {
        item.pos.y = topVal;
      } else if (card.offsetTop !== undefined && card.offsetTop > 0) {
        item.pos.y = card.offsetTop;
      }

      if (!isNaN(widthVal) && widthVal > 0) {
        item.pos.w = widthVal;
      }
      if (!isNaN(heightVal) && heightVal > 0) {
        item.pos.h = heightVal;
      }
    });
  }

  // ===========================================================================
  // 2. DASHBOARD STATE (TABS, ITEMS, FRAMES, WIDGETS, FILTERS)
  // ===========================================================================
  const DashboardState = {
    present: {
      id: 'dashboard_' + Date.now(),
      title: 'Báo Cáo Hoạt Động & Phân Tích Kinh Doanh',
      description: 'Dashboard BI tùy biến hoàn toàn bằng JavaScript ECharts và nguồn dữ liệu Excel',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      theme: { themeMode: 'light', canvasBg: (typeof localStorage !== 'undefined' && localStorage.getItem('mb_dashboard_canvas_bg')) || 'dots' },
      tabs: [
        { id: 'tab_overview', name: '📊 Tổng Quan Kinh Doanh' }
      ],
      activeTabId: 'tab_overview',
      tabFilters: {}, // tabId -> { column -> filterState }
      items: [] // charts, frames, widgets
    },
    past: [],
    future: [],
    isDirty: false,
    mode: 'edit', // 'edit' | 'preview'
    zoom: 1.0,
    pan: { x: 0, y: 0 },
    selectedItemId: null,
    activeFilters: {}, // column -> value or { op, val } of current active tab
    tabFilters: {},    // tabId -> { column -> filterState }

    pushSnapshot() {
      syncItemPositionsFromDOM();
      if (!this.tabFilters) this.tabFilters = {};
      const currentTab = this.present.activeTabId || 'default';
      this.tabFilters[currentTab] = JSON.parse(JSON.stringify(this.activeFilters || {}));
      this.present.tabFilters = JSON.parse(JSON.stringify(this.tabFilters));
      this.present.activeFilters = JSON.parse(JSON.stringify(this.activeFilters || {}));
      const snapshot = JSON.parse(JSON.stringify(this.present));
      delete snapshot.embeddedData; // Loại bỏ dữ liệu dataset lớn khỏi mảng lịch sử Undo/Redo để tránh nhân bản RAM
      this.past.push(snapshot);
      if (this.past.length > 30) this.past.shift();
      this.future = [];
      this.isDirty = true;
      this.present.updatedAt = new Date().toISOString();
      this.updateUI();
      triggerDebouncedAutoSave();
    },

    undo() {
      if (this.past.length === 0) return;
      if (!this.tabFilters) this.tabFilters = {};
      const currentTab = this.present.activeTabId || 'default';
      this.tabFilters[currentTab] = JSON.parse(JSON.stringify(this.activeFilters || {}));
      this.present.tabFilters = JSON.parse(JSON.stringify(this.tabFilters));
      this.present.activeFilters = JSON.parse(JSON.stringify(this.activeFilters || {}));
      this.future.push(JSON.parse(JSON.stringify(this.present)));
      this.present = this.past.pop();
      this.tabFilters = JSON.parse(JSON.stringify(this.present.tabFilters || {}));
      const curTab = this.present.activeTabId || 'default';
      this.activeFilters = JSON.parse(JSON.stringify((this.tabFilters && this.tabFilters[curTab]) || this.present.activeFilters || {}));
      this.isDirty = true;
      renderAllTabs();
      renderCanvas();
      this.updateUI();
      showToast('Đã Hoàn Tác (Undo)', 'info');
    },

    redo() {
      if (this.future.length === 0) return;
      if (!this.tabFilters) this.tabFilters = {};
      const currentTab = this.present.activeTabId || 'default';
      this.tabFilters[currentTab] = JSON.parse(JSON.stringify(this.activeFilters || {}));
      this.present.tabFilters = JSON.parse(JSON.stringify(this.tabFilters));
      this.present.activeFilters = JSON.parse(JSON.stringify(this.activeFilters || {}));
      this.past.push(JSON.parse(JSON.stringify(this.present)));
      this.present = this.future.pop();
      this.tabFilters = JSON.parse(JSON.stringify(this.present.tabFilters || {}));
      const curTab = this.present.activeTabId || 'default';
      this.activeFilters = JSON.parse(JSON.stringify((this.tabFilters && this.tabFilters[curTab]) || this.present.activeFilters || {}));
      this.isDirty = true;
      renderAllTabs();
      renderCanvas();
      this.updateUI();
      showToast('Đã Làm Lại (Redo)', 'info');
    },

    updateUI() {
      const undoBtn = document.getElementById('btnUndo');
      const redoBtn = document.getElementById('btnRedo');
      if (undoBtn) undoBtn.disabled = this.past.length === 0;
      if (redoBtn) redoBtn.disabled = this.future.length === 0;

      const dirtyBadge = document.getElementById('dirtyBadge');
      if (dirtyBadge) {
        if (this.isDirty) {
          dirtyBadge.textContent = '● Chưa lưu thay đổi';
          dirtyBadge.className = 'db-dirty-badge unsaved';
        } else {
          dirtyBadge.textContent = '✓ Đã lưu';
          dirtyBadge.className = 'db-dirty-badge saved';
        }
      }

      const titleInput = document.getElementById('dashboardTitleInput');
      if (titleInput && document.activeElement !== titleInput && titleInput.value !== (this.present.title || '')) {
        titleInput.value = this.present.title || '';
      }

      const themeBtn = document.getElementById('btnToggleTheme');
      if (themeBtn) {
        const isDark = (document.documentElement.getAttribute('data-theme') || 'dark') === 'dark';
        themeBtn.innerHTML = isDark ? '🌙 Dark Mode' : '☀️ Light Mode';
      }
    }
  };

  let autoSaveSessionTimer = null;
  function triggerDebouncedAutoSave() {
    clearTimeout(autoSaveSessionTimer);
    autoSaveSessionTimer = setTimeout(async () => {
      saveLastEditedSessionToLocalStorage();
      if (typeof autoSaveToActiveDirectory === 'function') {
        await autoSaveToActiveDirectory();
      }
    }, 400);
  }

  function saveLastEditedSessionToLocalStorage() {
    try {
      if (!DashboardState.present) return;
      if (!Array.isArray(DashboardState.present.items) && !Array.isArray(DashboardState.present.tabs)) return;

      const jsonStr = prepareDashboardConfigJson();
      const sessionObj = JSON.parse(jsonStr);

      try {
        localStorage.setItem('mb_dashboard_last_edited_session', JSON.stringify(sessionObj));
      } catch (quotaErr) {
        // Nếu chạm trần quota 5MB của localStorage, loại bỏ embeddedData lớn để bảo toàn 100% cấu hình items, tabs, charts...
        delete sessionObj.embeddedData;
        try {
          localStorage.setItem('mb_dashboard_last_edited_session', JSON.stringify(sessionObj));
        } catch (_) {}
      }

      // Lưu siêu dữ liệu các file trong DataStore nếu có
      if (DataStore.files && Object.keys(DataStore.files).length > 0) {
        const filesMeta = {};
        for (const fn in DataStore.files) {
          const fObj = DataStore.files[fn];
          filesMeta[fn] = {
            name: fObj.name,
            sizeStr: fObj.sizeStr,
            uploadedAt: fObj.uploadedAt,
            sheets: fObj.sheets
          };
        }
        try {
          localStorage.setItem('mb_dashboard_last_files_data', JSON.stringify(filesMeta));
        } catch (_) {}
      }
    } catch (e) {
      console.warn('Auto-save session to localStorage failed:', e);
    }
  }

  // ===========================================================================
  // 3. SYNTAX HIGHLIGHTING, INDENT RAINBOW & BLACK TERMINAL ENGINE
  // ===========================================================================
  function renderIndentGuides(indentStr) {
    if (!indentStr) return '';
    let html = '';
    let level = 1;
    let i = 0;

    while (i < indentStr.length) {
      const ch = indentStr[i];
      const lvlClass = `indent-lvl-${((level - 1) % 6) + 1}`;

      if (ch === '\t') {
        html += `<span class="indent-guide ${lvlClass}">\t</span>`;
        level++;
        i++;
      } else if (ch === ' ' && indentStr[i + 1] === ' ') {
        html += `<span class="indent-guide ${lvlClass}">  </span>`;
        level++;
        i += 2;
      } else {
        html += `<span class="indent-guide ${lvlClass}">${ch}</span>`;
        i++;
      }
    }
    return html;
  }

  const JS_TOKEN_REGEX = /(\/\*[\s\S]*?\*\/|\/\/[^\n]*)|("(?:\\.|[^"\n\\])*"|'(?:\\.|[^'\n\\])*'|`(?:\\.|[^`\\])*`|`(?:\\.|[^`\\])*$)|(\b(?:const|let|var|function|return|if|else|for|while|do|switch|case|default|break|continue|try|catch|finally|throw|new|typeof|instanceof|class|extends|import|export|from|async|await|yield|in|of|void|delete|debugger)\b)|(\b(?:true|false|null|undefined|NaN|Infinity)\b)|(\b(?:XLSX|files|getSheet|getFile|data|echarts|console|window|document|Array|Object|String|Number|Boolean|Math|JSON|Date|RegExp|Map|Set|Promise)\b)|(\b\d+(?:\.\d+)?(?:e[+-]?\d+)?\b)|(\b[a-zA-Z_$][a-zA-Z0-9_$]*(?=\s*\())|(\b[a-zA-Z_$][a-zA-Z0-9_$]*(?=\s*:))|([{}()\[\]])|(===|!==|==|!=|<=|>=|=>|&&|\|\||\?\?|\+=|-=|\*=|\/=|[\+\-\*\/%=\?:\!&\|\^~<>])/g;

  function highlightCodeLine(codePart, state) {
    if (state.inBlockComment) {
      const closeIdx = codePart.indexOf('*/');
      if (closeIdx !== -1) {
        const commentText = codePart.substring(0, closeIdx + 2);
        state.inBlockComment = false;
        return `<span class="token-comment">${escapeHtml(commentText)}</span>` + highlightCodeLine(codePart.substring(closeIdx + 2), state);
      } else {
        return `<span class="token-comment">${escapeHtml(codePart)}</span>`;
      }
    }

    if (state.inTemplateStr) {
      let closeIdx = -1;
      for (let i = 0; i < codePart.length; i++) {
        if (codePart[i] === '`' && (i === 0 || codePart[i - 1] !== '\\')) {
          closeIdx = i;
          break;
        }
      }
      if (closeIdx !== -1) {
        const strText = codePart.substring(0, closeIdx + 1);
        state.inTemplateStr = false;
        return `<span class="token-string">${escapeHtml(strText)}</span>` + highlightCodeLine(codePart.substring(closeIdx + 1), state);
      } else {
        return `<span class="token-string">${escapeHtml(codePart)}</span>`;
      }
    }

    let lastIndex = 0;
    let html = '';
    JS_TOKEN_REGEX.lastIndex = 0;
    let match;

    while ((match = JS_TOKEN_REGEX.exec(codePart)) !== null) {
      if (match.index > lastIndex) {
        html += escapeHtml(codePart.substring(lastIndex, match.index));
      }

      const [full, comment, str, kw, bool, builtin, num, func, prop, bracket, op] = match;

      if (comment) {
        if (comment.startsWith('/*') && !comment.endsWith('*/')) state.inBlockComment = true;
        html += `<span class="token-comment">${escapeHtml(comment)}</span>`;
      } else if (str) {
        if (str.startsWith('`') && (str.length === 1 || !str.endsWith('`') || str.endsWith('\\`'))) state.inTemplateStr = true;
        html += `<span class="token-string">${escapeHtml(str)}</span>`;
      } else if (kw) {
        html += `<span class="token-keyword">${escapeHtml(kw)}</span>`;
      } else if (bool) {
        html += `<span class="token-boolean">${escapeHtml(bool)}</span>`;
      } else if (builtin) {
        html += `<span class="token-builtin">${escapeHtml(builtin)}</span>`;
      } else if (num) {
        html += `<span class="token-number">${escapeHtml(num)}</span>`;
      } else if (func) {
        html += `<span class="token-func">${escapeHtml(func)}</span>`;
      } else if (prop) {
        html += `<span class="token-prop">${escapeHtml(prop)}</span>`;
      } else if (bracket) {
        let bClass = 'token-bracket-1';
        if ('{([<'.includes(bracket)) {
          bClass = `token-bracket-${(state.bracketDepth % 3) + 1}`;
          state.bracketDepth++;
        } else {
          state.bracketDepth = Math.max(0, state.bracketDepth - 1);
          bClass = `token-bracket-${(state.bracketDepth % 3) + 1}`;
        }
        html += `<span class="${bClass}">${escapeHtml(bracket)}</span>`;
      } else if (op) {
        html += `<span class="token-operator">${escapeHtml(op)}</span>`;
      } else {
        html += escapeHtml(full);
      }

      lastIndex = JS_TOKEN_REGEX.lastIndex;
    }

    if (lastIndex < codePart.length) {
      html += escapeHtml(codePart.substring(lastIndex));
    }

    return html;
  }

  function highlightJS(code) {
    if (!code) return '';
    const lines = code.split('\n');
    const state = { inBlockComment: false, inTemplateStr: false, bracketDepth: 0 };

    const highlightedLines = lines.map(line => {
      const indentMatch = line.match(/^([ \t]+)/);
      const indentStr = indentMatch ? indentMatch[1] : '';
      const codePart = line.substring(indentStr.length);
      const indentHtml = renderIndentGuides(indentStr);
      const codeHtml = highlightCodeLine(codePart, state);
      return indentHtml + codeHtml;
    });

    return highlightedLines.join('\n');
  }

  function formatJSCode(code) {
    if (!code || typeof code !== 'string') return code;
    const lines = code.split('\n');
    let indentLevel = 0;
    const indentStr = '  ';
    const resultLines = [];
    let inMultiComment = false;

    for (let i = 0; i < lines.length; i++) {
      let line = lines[i].trim();
      if (!line) {
        if (resultLines.length > 0 && resultLines[resultLines.length - 1] !== '') {
          resultLines.push('');
        }
        continue;
      }

      if (line.startsWith('/*')) inMultiComment = true;
      if (inMultiComment) {
        resultLines.push(indentStr.repeat(indentLevel) + line);
        if (line.includes('*/')) inMultiComment = false;
        continue;
      }

      let leadingClosers = 0;
      for (let c of line) {
        if (c === '}' || c === ']' || c === ')') leadingClosers++;
        else if (c === ' ' || c === '\t') continue;
        else break;
      }

      let currentIndent = Math.max(0, indentLevel - leadingClosers);
      resultLines.push(indentStr.repeat(currentIndent) + line);

      let delta = 0;
      let inStr = null;
      let isEscaped = false;

      for (let j = 0; j < line.length; j++) {
        const ch = line[j];
        if (inStr) {
          if (ch === inStr && !isEscaped) inStr = null;
          isEscaped = (ch === '\\' && !isEscaped);
          continue;
        }
        if (ch === '"' || ch === "'" || ch === '`') {
          inStr = ch;
          isEscaped = false;
          continue;
        }
        if (ch === '/' && line[j + 1] === '/') break;
        if (ch === '{' || ch === '[' || ch === '(') delta++;
        else if (ch === '}' || ch === ']' || ch === ')') delta--;
      }

      indentLevel = Math.max(0, indentLevel + delta);
    }

    return resultLines.join('\n');
  }

  function updateCeCodeHighlight() {
    const textarea = document.getElementById('ceCodeTextarea');
    const highlight = document.getElementById('ceCodeHighlight');
    const lineNumbers = document.getElementById('ceLineNumbers');
    if (!textarea || !highlight) return;

    const code = textarea.value;
    highlight.innerHTML = highlightJS(code) + '\n';

    if (lineNumbers) {
      const lineCount = Math.max(1, code.split('\n').length);
      let nums = '';
      for (let i = 1; i <= lineCount; i++) {
        nums += i + '\n';
      }
      lineNumbers.textContent = nums;
    }
  }

  function printCeTerminal(text, type = 'log') {
    const body = document.getElementById('ceTerminalBody');
    const statusDot = document.getElementById('ceTermStatusDot');
    if (!body) return;

    const line = document.createElement('div');
    line.className = `ce-term-line ce-term-${type}`;
    const timestamp = new Date().toLocaleTimeString('vi-VN', { hour12: false });

    let prefix = '> ';
    if (type === 'error') {
      prefix = '❌ [Error] ';
      if (statusDot) statusDot.style.color = '#ff7b72';
    } else if (type === 'success') {
      prefix = '✓ [Render] ';
      if (statusDot) statusDot.style.color = '#3fb950';
    } else if (type === 'warn') {
      prefix = '⚠️ [Warn] ';
    } else if (type === 'info') {
      prefix = 'ℹ [Info] ';
    }

    line.textContent = `[${timestamp}] ${prefix}${text}`;
    body.appendChild(line);
    body.scrollTop = body.scrollHeight;
  }

  function clearCeTerminal() {
    const body = document.getElementById('ceTerminalBody');
    const statusDot = document.getElementById('ceTermStatusDot');
    if (body) {
      body.innerHTML = '<div class="ce-term-line ce-term-system">> Terminal đã được làm sạch.</div>';
    }
    if (statusDot) statusDot.style.color = '#3fb950';
  }

  function formatConsoleArg(arg) {
    if (arg === null) return 'null';
    if (arg === undefined) return 'undefined';
    if (typeof arg === 'object') {
      try {
        return JSON.stringify(arg);
      } catch (e) {
        return String(arg);
      }
    }
    return String(arg);
  }

  // ===========================================================================
  // 3. DATA FILTERING ENGINE & ECHARTS EVALUATION
  // ===========================================================================
  function parseDateToTimestamp(val) {
    if (val === null || val === undefined || val === '') return null;
    if (val instanceof Date) {
      if (isNaN(val.getTime())) return null;
      return new Date(val.getFullYear(), val.getMonth(), val.getDate()).getTime();
    }
    if (typeof val === 'number') {
      if (val > 1000 && val < 100000) {
        const utcDays = Math.floor(val - 25569);
        const d = new Date(utcDays * 86400 * 1000);
        return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
      }
      if (val > 1000000000000) {
        const d = new Date(val);
        return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
      }
    }
    const s = String(val).trim();
    if (!s) return null;

    // YYYY-MM-DD or YYYY/MM/DD or YYYY.MM.DD (ưu tiên chuẩn ISO trước)
    const ymd = s.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})/);
    if (ymd) {
      const year = parseInt(ymd[1], 10);
      const month = parseInt(ymd[2], 10) - 1;
      const day = parseInt(ymd[3], 10);
      const d = new Date(year, month, day);
      if (!isNaN(d.getTime())) return d.getTime();
    }

    // DD/MM/YYYY or DD-MM-YYYY or DD.MM.YYYY
    const dmy = s.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})/);
    if (dmy) {
      const day = parseInt(dmy[1], 10);
      const month = parseInt(dmy[2], 10) - 1;
      const year = parseInt(dmy[3], 10);
      const d = new Date(year, month, day);
      if (!isNaN(d.getTime())) return d.getTime();
    }

    // Do not parse pure numeric strings (like 100, 50, 400) as dates
    if (/^\d{1,7}$/.test(s) || /^\d{9,}$/.test(s)) {
      return null;
    }

    const p = Date.parse(s);
    if (!isNaN(p)) {
      const d = new Date(p);
      return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
    }
    return null;
  }

  function formatDateToYMD(val) {
    if (!val) return '';
    const s = String(val).trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    const ts = parseDateToTimestamp(val);
    if (!ts) return '';
    const d = new Date(ts);
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  }

  function isProbableDateColumn(colName, rows = []) {
    if (!colName) return false;
    const lower = String(colName).toLowerCase();
    if (lower.includes('ngày') || lower.includes('ngay') || lower.includes('date') || lower.includes('time') || lower.includes('thời gian') || lower.includes('thoi_gian')) {
      return true;
    }
    if (!rows || rows.length === 0) return false;
    let validDateCount = 0;
    let checkedCount = 0;
    for (let i = 0; i < Math.min(rows.length, 20); i++) {
      const v = rows[i][colName];
      if (v !== undefined && v !== null && v !== '') {
        checkedCount++;
        if (parseDateToTimestamp(v) !== null) validDateCount++;
      }
    }
    return checkedCount > 0 && (validDateCount / checkedCount) >= 0.6;
  }

  function getFilteredData(sourceKey = null, excludeCols = [], targetTabId = null) {
    let rows = DataStore.getData(sourceKey);
    if (!rows || rows.length === 0) return [];

    const tabId = targetTabId || DashboardState.present.activeTabId || 'default';
    let filters = {};
    if (DashboardState.tabFilters && DashboardState.tabFilters[tabId]) {
      filters = DashboardState.tabFilters[tabId];
    } else if (tabId === DashboardState.present.activeTabId) {
      filters = DashboardState.activeFilters || {};
    }

    const filterKeys = Object.keys(filters).filter(k => !excludeCols.includes(k));
    if (filterKeys.length === 0) return rows;

    return rows.filter(row => {
      for (const col of filterKeys) {
        const filterVal = filters[col];
        if (!filterVal) continue;
        if (row[col] === undefined) continue;

        let op = '=';
        let target = '';
        let startDate = '';
        let endDate = '';
        let minVal = '';
        let maxVal = '';
        let valuesList = null;
        let isDateFilter = false;

        if (typeof filterVal === 'object' && filterVal !== null) {
          op = filterVal.op || '=';
          target = filterVal.val !== undefined ? filterVal.val : '';
          startDate = filterVal.start || '';
          endDate = filterVal.end || '';
          minVal = filterVal.min !== undefined ? filterVal.min : (filterVal.start !== undefined ? filterVal.start : '');
          maxVal = filterVal.max !== undefined ? filterVal.max : (filterVal.end !== undefined ? filterVal.end : '');
          valuesList = Array.isArray(filterVal.values) ? filterVal.values : (Array.isArray(filterVal.val) ? filterVal.val : null);
          isDateFilter = Boolean(filterVal.isDate || op === 'date_range' || op.startsWith('date_'));
        } else {
          target = filterVal;
        }

        const cell = row[col];
        if (cell === undefined || cell === null) return false;

        // 1. Bộ lọc Trong Khoảng (Between / Range) - Hỗ trợ cả Số và Ngày
        if (op === 'between' || op === 'range' || op === 'date_range') {
          // A. Nếu là Date Range
          if (isDateFilter || op === 'date_range') {
            const cellTs = parseDateToTimestamp(cell);
            if (cellTs === null) return false;
            let startTs = startDate ? parseDateToTimestamp(startDate) : (minVal ? parseDateToTimestamp(minVal) : null);
            let endTs = endDate ? parseDateToTimestamp(endDate) : (maxVal ? parseDateToTimestamp(maxVal) : null);
            if (startTs !== null && endTs !== null && startTs > endTs) {
              const tmp = startTs;
              startTs = endTs;
              endTs = tmp;
            }
            if (startTs !== null && cellTs < startTs) return false;
            if (endTs !== null && cellTs > endTs) return false;
            continue;
          }

          // B. Nếu là Số hoặc Giá trị thường
          const sMin = minVal !== undefined && minVal !== null ? String(minVal).trim() : '';
          const sMax = maxVal !== undefined && maxVal !== null ? String(maxVal).trim() : '';
          if (sMin === '' && sMax === '') continue; // Không nhập mốc nào => Không lọc

          const numCell = Number(cell);
          const hasNumCell = !isNaN(numCell);
          let numMin = sMin !== '' ? Number(sMin) : null;
          let numMax = sMax !== '' ? Number(sMax) : null;

          if (numMin !== null && numMax !== null && !isNaN(numMin) && !isNaN(numMax) && numMin > numMax) {
            const tmp = numMin;
            numMin = numMax;
            numMax = tmp;
          }

          if (hasNumCell && ((numMin !== null && !isNaN(numMin)) || (numMax !== null && !isNaN(numMax)))) {
            if (numMin !== null && !isNaN(numMin) && numCell < numMin) return false;
            if (numMax !== null && !isNaN(numMax) && numCell > numMax) return false;
            continue;
          } else {
            // Chuỗi
            let strMin = sMin !== '' ? sMin.toLowerCase() : null;
            let strMax = sMax !== '' ? sMax.toLowerCase() : null;
            if (strMin && strMax && strMin > strMax) {
              const tmp = strMin;
              strMin = strMax;
              strMax = tmp;
            }
            if (strMin !== null && String(cell).toLowerCase() < strMin) return false;
            if (strMax !== null && String(cell).toLowerCase() > strMax) return false;
            continue;
          }
        }

        // 2. Bộ lọc Chọn Nhiều Mục (Multi-Select / In list)
        if (op === 'in' || valuesList !== null) {
          const list = valuesList || [];
          if (!list || list.length === 0 || list.includes('all')) continue;
          const strCell = String(cell).trim().toLowerCase();
          const match = list.some(item => String(item).trim().toLowerCase() === strCell);
          if (!match) return false;
          continue;
        }

        // 3. Bộ lọc 1 giá trị: Nếu giá trị rỗng hoặc 'all' thì bỏ qua
        if (target === undefined || target === null || target === '' || target === 'all') {
          continue;
        }

        // 4. Lọc theo Date đơn (date_eq, date_gte, date_lte, >, <, >=, <=)
        if (isDateFilter) {
          const cellTs = parseDateToTimestamp(cell);
          if (cellTs === null) return false;
          const targetTs = parseDateToTimestamp(target);
          if (targetTs === null) continue;

          if (op === 'date_eq' || op === '=') {
            if (cellTs !== targetTs) return false;
          } else if (op === 'date_gte' || op === '>=') {
            if (cellTs < targetTs) return false;
          } else if (op === 'date_lte' || op === '<=') {
            if (cellTs > targetTs) return false;
          } else if (op === '>') {
            if (cellTs <= targetTs) return false;
          } else if (op === '<') {
            if (cellTs >= targetTs) return false;
          }
          continue;
        }

        // 5. Lọc Văn Bản (Text Search: contains, starts_with, ends_with, not_contains)
        const strCell = String(cell).toLowerCase();
        const strTarget = String(target).toLowerCase();

        if (op === 'contains') {
          if (!strCell.includes(strTarget)) return false;
          continue;
        } else if (op === 'starts_with') {
          if (!strCell.startsWith(strTarget)) return false;
          continue;
        } else if (op === 'ends_with') {
          if (!strCell.endsWith(strTarget)) return false;
          continue;
        } else if (op === 'not_contains') {
          if (strCell.includes(strTarget)) return false;
          continue;
        }

        // 6. Lọc So Sánh 1 Giá Trị Số / Chuỗi (=, >, >=, <, <=, !=)
        const numCell = Number(cell);
        const numTarget = Number(target);
        const isBothNumber = !isNaN(numCell) && !isNaN(numTarget) && String(target).trim() !== '';

        if (op === '>') {
          if (isBothNumber) {
            if (!(numCell > numTarget)) return false;
          } else {
            if (!(String(cell) > String(target))) return false;
          }
        } else if (op === '>=') {
          if (isBothNumber) {
            if (!(numCell >= numTarget)) return false;
          } else {
            if (!(String(cell) >= String(target))) return false;
          }
        } else if (op === '<') {
          if (isBothNumber) {
            if (!(numCell < numTarget)) return false;
          } else {
            if (!(String(cell) < String(target))) return false;
          }
        } else if (op === '<=') {
          if (isBothNumber) {
            if (!(numCell <= numTarget)) return false;
          } else {
            if (!(String(cell) <= String(target))) return false;
          }
        } else if (op === '!=') {
          if (isBothNumber) {
            if (numCell === numTarget) return false;
          } else {
            if (String(cell).trim().toLowerCase() === String(target).trim().toLowerCase()) return false;
          }
        } else {
          // op === '='
          if (isBothNumber) {
            if (numCell !== numTarget) return false;
          } else {
            if (String(cell).trim().toLowerCase() !== String(target).trim().toLowerCase()) return false;
          }
        }
      }
      return true;
    });
  }

  function getFilteredFilesObject(targetTabId = null) {
    const allFiles = DataStore.getAllFilesObject();
    const filtered = {};
    const tabId = targetTabId || DashboardState.present.activeTabId || 'default';
    for (const fileName in allFiles) {
      filtered[fileName] = {};
      for (const sheetName in allFiles[fileName]) {
        filtered[fileName][sheetName] = getFilteredData(`${fileName}::${sheetName}`, [], tabId);
      }
    }
    return filtered;
  }

  // ===========================================================================
  // 3.5. SMART ECHARTS ACCELERATION FOR LARGE DATASETS
  // ===========================================================================
  function optimizeChartOptionForLargeData(option) {
    if (!option || typeof option !== 'object') return option;

    let hasHeavyData = false;
    const seriesList = Array.isArray(option.series) 
      ? option.series 
      : (option.series && typeof option.series === 'object' ? [option.series] : []);

    seriesList.forEach(s => {
      if (!s || typeof s !== 'object') return;
      const dataLen = Array.isArray(s.data) ? s.data.length : 0;
      if (dataLen >= 300) {
        hasHeavyData = true;
        // 1. Kích hoạt Canvas Large Mode của ECharts cho line, bar, scatter
        if (s.large === undefined && (s.type === 'line' || s.type === 'bar' || s.type === 'scatter')) {
          s.large = true;
          s.largeThreshold = 300;
        }
        // 2. Kích hoạt thuật toán lấy mẫu LTTB giữ biên độ đồ thị nhưng giảm 90% điểm vẽ
        if (s.type === 'line' && !s.sampling) {
          s.sampling = 'lttb';
        }
        // 3. Phân luồng render tuần tự (Progressive Rendering)
        if (dataLen >= 1500 && s.progressive === undefined) {
          s.progressive = 800;
          s.progressiveThreshold = 1500;
        }
      }
    });

    // Tắt animation để tránh đơ giao diện khi số điểm dữ liệu quá lớn
    if (hasHeavyData && option.animation === undefined) {
      option.animation = false;
    }

    return option;
  }

  function enhanceInteractiveGraphic(option) {
    if (!option || typeof option !== 'object' || !option.graphic) return option;
    const processItem = (g) => {
      if (!g || typeof g !== 'object') return;
      const hasInteraction = typeof g.onclick === 'function' || typeof g.onmouseover === 'function' || typeof g.onmousedown === 'function';
      if (hasInteraction) {
        if (g.cursor === undefined) g.cursor = 'pointer';
        if (!g.style) g.style = {};
        if (g.style.cursor === undefined) g.style.cursor = 'pointer';
        if (g.z === undefined) g.z = 100;
        if (g.silent === true) g.silent = false;
      }
      if (Array.isArray(g.children)) {
        g.children.forEach(processItem);
      }
    };

    if (Array.isArray(option.graphic)) {
      option.graphic.forEach(processItem);
    } else if (typeof option.graphic === 'object') {
      processItem(option.graphic);
    }
    return option;
  }

  function evaluateEChartsCode(code, optionalData = null, context = {}, chartInstance = null) {
    if (!code || !code.trim()) {
      return { option: {}, error: 'Chưa có code JavaScript cấu hình.', logs: [] };
    }

    const logs = [];
    const customConsole = {
      log: (...args) => logs.push({ type: 'log', text: args.map(formatConsoleArg).join(' ') }),
      info: (...args) => logs.push({ type: 'info', text: args.map(formatConsoleArg).join(' ') }),
      warn: (...args) => logs.push({ type: 'warn', text: args.map(formatConsoleArg).join(' ') }),
      error: (...args) => logs.push({ type: 'error', text: args.map(formatConsoleArg).join(' ') })
    };

    try {
      const allFiles = DataStore.getAllFilesObject();
      const tabId = context.tabId || DashboardState.present.activeTabId || 'default';
      const filteredFiles = getFilteredFilesObject(tabId);

      // Dữ liệu mặc định: dùng optionalData nếu được truyền (kể cả mảng rỗng []), nếu không thì gọi getFilteredData()
      let defaultData = optionalData;
      if (defaultData === null || defaultData === undefined) {
        defaultData = getFilteredData(null, [], tabId);
      }

      // Hàm tiện ích getSheet(sheetName, fileName?) cho người dùng gọi trong code -> TRẢ VỀ DỮ LIỆU ĐÃ LỌC
      const getSheet = (sheetName, fileName = null) => {
        if (fileName && filteredFiles[fileName] && filteredFiles[fileName][sheetName]) {
          return filteredFiles[fileName][sheetName];
        }
        for (const f in filteredFiles) {
          if (filteredFiles[f][sheetName]) return filteredFiles[f][sheetName];
        }
        return [];
      };

      // Dữ liệu gốc chưa lọc (để so sánh hoặc tính % tỷ trọng trên toàn bộ dữ liệu gốc)
      const getRawSheet = (sheetName, fileName = null) => {
        if (fileName && allFiles[fileName] && allFiles[fileName][sheetName]) {
          return allFiles[fileName][sheetName];
        }
        for (const f in allFiles) {
          if (allFiles[f][sheetName]) return allFiles[f][sheetName];
        }
        return [];
      };

      // Hàm tiện ích getFile(fileName)
      const getFile = (fileName) => {
        return filteredFiles[fileName] || null;
      };

      const setFilter = (col, val) => {
        if (!col) return;
        if (!DashboardState.tabFilters) DashboardState.tabFilters = {};
        if (!DashboardState.tabFilters[tabId]) DashboardState.tabFilters[tabId] = {};
        if (val === undefined || val === null || val === '' || val === 'all') {
          delete DashboardState.tabFilters[tabId][col];
          if (tabId === (DashboardState.present?.activeTabId || 'default')) {
            delete DashboardState.activeFilters[col];
          }
        } else {
          DashboardState.tabFilters[tabId][col] = val;
          if (tabId === (DashboardState.present?.activeTabId || 'default')) {
            DashboardState.activeFilters[col] = val;
          }
        }
        if (DashboardState.present) {
          DashboardState.present.tabFilters = JSON.parse(JSON.stringify(DashboardState.tabFilters));
          DashboardState.present.activeFilters = JSON.parse(JSON.stringify(DashboardState.activeFilters));
        }
        renderCanvas();
      };

      const resetFilter = (col = null) => {
        if (col) {
          setFilter(col, null);
        } else {
          if (!DashboardState.tabFilters) DashboardState.tabFilters = {};
          DashboardState.tabFilters[tabId] = {};
          if (tabId === (DashboardState.present?.activeTabId || 'default')) {
            DashboardState.activeFilters = {};
          }
          if (DashboardState.present) {
            DashboardState.present.tabFilters = JSON.parse(JSON.stringify(DashboardState.tabFilters));
            DashboardState.present.activeFilters = JSON.parse(JSON.stringify(DashboardState.activeFilters));
          }
          renderCanvas();
        }
      };

      const curFilters = (DashboardState.tabFilters && DashboardState.tabFilters[tabId]) || DashboardState.activeFilters || {};
      const safeContext = {
        setFilter,
        resetFilter,
        activeFilters: curFilters,
        rawFiles: allFiles,
        getRawSheet,
        tabId,
        myChart: chartInstance,
        chart: chartInstance,
        ...context
      };

      // Khởi tạo hàm thực thi nhận data, files, getSheet, getFile, context, echarts, XLSX, console, myChart, chart, setFilter, resetFilter
      const evalFunc = new Function(
        'data', 'files', 'getSheet', 'getFile', 'context', 'echarts', 'XLSX', 'console', 'myChart', 'chart', 'setFilter', 'resetFilter',
        `
        let option = null;
        try {
          ${code}
        } catch (e) {
          return { __error: e.stack || e.message };
        }
        if (typeof option === 'object' && option !== null) return option;
        `
      );

      const startTime = performance.now();
      let result = evalFunc(
        defaultData,
        filteredFiles,
        getSheet,
        getFile,
        safeContext,
        window.echarts,
        window.XLSX,
        customConsole,
        chartInstance,
        chartInstance,
        setFilter,
        resetFilter
      );
      const durationMs = Math.round(performance.now() - startTime);

      if (result && result.__error) {
        return { option: {}, error: result.__error, logs, durationMs };
      }

      if ((!result || typeof result !== 'object') && chartInstance && chartInstance.getOption) {
        const curOpt = chartInstance.getOption();
        if (curOpt && Object.keys(curOpt).length > 0) {
          result = curOpt;
        }
      }

      if (!result || typeof result !== 'object') {
        return { option: {}, error: 'Code phải return về một object option ECharts (Ví dụ: return { xAxis: {...}, series: [...] };) hoặc gán myChart.setOption(option);', logs, durationMs };
      }

      enhanceInteractiveGraphic(result);

      // Tự động tối ưu hoá ECharts option nếu tập dữ liệu lớn
      const optimizedOption = optimizeChartOptionForLargeData(result);
      return { option: optimizedOption, error: null, logs, durationMs };
    } catch (err) {
      return { option: {}, error: err.message || String(err), logs };
    }
  }

  // ===========================================================================
  // 4. CHART TEMPLATES GENERATOR (PRESET CODE SNIPPETS)
  // ===========================================================================
  const ChartTemplates = {
    _getCols(sourceKey = null) {
      if (sourceKey) return DataStore.getColumns(sourceKey);
      const sources = DataStore.getAvailableSources();
      return DataStore.getColumns(sources[0] ? sources[0].key : null);
    },

    bar(sourceKey = null) {
      const cols = this._getCols(sourceKey);
      const catCol = cols.includes('Tháng') ? 'Tháng' : (cols[1] || cols[0] || 'Danh Mục');
      const valCol = cols.includes('Doanh Thu') ? 'Doanh Thu' : (cols[4] || cols[3] || cols[1] || 'Giá Trị');

      return `// 💡 Hướng dẫn lấy dữ liệu Excel trong code:
// - Biến 'data': Mảng dòng mặc định từ Sheet đầu tiên đã nạp (đã áp dụng bộ lọc).
// - Lấy sheet bất kỳ: const rows = getSheet('TênSheet');
// - Lấy file cụ thể: const rows = files['TenFile.xlsx']['TênSheet'];

const categories = data.map(row => row['${catCol}']);
const values = data.map(row => Number(row['${valCol}']) || 0);

return {
  title: { text: '${valCol} theo ${catCol}', textStyle: { color: '#0f172a', fontSize: 13, fontWeight: 700 } },
  tooltip: { trigger: 'axis', valueFormatter: v => new Intl.NumberFormat('vi-VN').format(v) },
  grid: { left: '4%', right: '4%', bottom: '10%', top: '16%', containLabel: true },
  xAxis: {
    type: 'category',
    data: categories,
    axisLabel: { color: '#64748b' }
  },
  yAxis: {
    type: 'value',
    splitLine: { lineStyle: { color: 'rgba(0,0,0,0.06)' } },
    axisLabel: { color: '#64748b' }
  },
  series: [
    {
      name: '${valCol}',
      type: 'bar',
      data: values,
      itemStyle: {
        color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
          { offset: 0, color: '#0284c7' },
          { offset: 1, color: '#0369a1' }
        ]),
        borderRadius: [4, 4, 0, 0]
      }
    }
  ]
};`;
    },

    horizontal_bar(sourceKey = null) {
      const cols = this._getCols(sourceKey);
      const catCol = cols[2] || cols[0] || 'Nhóm Sản Phẩm';
      const valCol = cols[3] || cols[1] || 'Doanh Thu';

      return `// Biểu đồ cột nằm ngang sắp xếp theo giá trị
const grouped = {};
data.forEach(r => {
  const k = r['${catCol}'] || 'Khác';
  grouped[k] = (grouped[k] || 0) + (Number(r['${valCol}']) || 0);
});

const sorted = Object.keys(grouped).map(k => ({ name: k, value: grouped[k] }))
  .sort((a, b) => a.value - b.value);

return {
  title: { text: '${valCol} theo ${catCol}', textStyle: { color: '#0f172a', fontSize: 13, fontWeight: 700 } },
  tooltip: { trigger: 'axis', valueFormatter: v => new Intl.NumberFormat('vi-VN').format(v) },
  grid: { left: '4%', right: '6%', bottom: '8%', top: '16%', containLabel: true },
  xAxis: { type: 'value', splitLine: { lineStyle: { color: 'rgba(0,0,0,0.06)' } }, axisLabel: { color: '#64748b' } },
  yAxis: { type: 'category', data: sorted.map(d => d.name), axisLabel: { color: '#64748b' } },
  series: [{
    name: '${valCol}',
    type: 'bar',
    data: sorted.map(d => d.value),
    itemStyle: {
      color: new echarts.graphic.LinearGradient(1, 0, 0, 0, [
        { offset: 0, color: '#38bdf8' },
        { offset: 1, color: '#0284c7' }
      ]),
      borderRadius: [0, 4, 4, 0]
    }
  }]
};`;
    },

    line(sourceKey = null) {
      const cols = this._getCols(sourceKey);
      const catCol = cols[0] || 'Tháng';
      const valCol = cols[3] || cols[1] || 'Doanh Thu';

      return `const categories = data.map(row => row['${catCol}']);
const values = data.map(row => Number(row['${valCol}']) || 0);

return {
  title: { text: 'Xu Hướng ${valCol}', textStyle: { color: '#0f172a', fontSize: 13, fontWeight: 700 } },
  tooltip: { trigger: 'axis', valueFormatter: v => new Intl.NumberFormat('vi-VN').format(v) },
  grid: { left: '4%', right: '4%', bottom: '10%', top: '16%', containLabel: true },
  xAxis: { type: 'category', data: categories, axisLabel: { color: '#64748b' } },
  yAxis: { type: 'value', splitLine: { lineStyle: { color: 'rgba(0,0,0,0.06)' } }, axisLabel: { color: '#64748b' } },
  series: [
    {
      name: '${valCol}',
      type: 'line',
      smooth: true,
      data: values,
      itemStyle: { color: '#10b981' },
      lineStyle: { width: 3 }
    }
  ]
};`;
    },

    area(sourceKey = null) {
      const cols = this._getCols(sourceKey);
      const catCol = cols[0] || 'Tháng';
      const valCol = cols[3] || cols[1] || 'Doanh Thu';

      return `const categories = data.map(row => row['${catCol}']);
const values = data.map(row => Number(row['${valCol}']) || 0);

return {
  title: { text: '${valCol} Vùng Phủ', textStyle: { color: '#0f172a', fontSize: 13, fontWeight: 700 } },
  tooltip: { trigger: 'axis', valueFormatter: v => new Intl.NumberFormat('vi-VN').format(v) },
  grid: { left: '4%', right: '4%', bottom: '10%', top: '16%', containLabel: true },
  xAxis: { type: 'category', data: categories, axisLabel: { color: '#64748b' } },
  yAxis: { type: 'value', splitLine: { lineStyle: { color: 'rgba(0,0,0,0.06)' } }, axisLabel: { color: '#64748b' } },
  series: [
    {
      name: '${valCol}',
      type: 'line',
      smooth: true,
      data: values,
      itemStyle: { color: '#f97316' },
      areaStyle: {
        color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
          { offset: 0, color: 'rgba(249, 115, 22, 0.45)' },
          { offset: 1, color: 'rgba(249, 115, 22, 0.02)' }
        ])
      }
    }
  ]
};`;
    },

    pie(sourceKey = null) {
      const cols = this._getCols(sourceKey);
      const catCol = cols[2] || cols[1] || cols[0] || 'Nhóm Sản Phẩm';
      const valCol = cols[3] || 'Doanh Thu';

      return `// Nhóm và tính tổng giá trị theo danh mục
const grouped = {};
data.forEach(row => {
  const cat = String(row['${catCol}'] || 'Khác');
  const val = Number(row['${valCol}']) || 0;
  grouped[cat] = (grouped[cat] || 0) + val;
});

const pieData = Object.keys(grouped).map(cat => ({
  name: cat,
  value: grouped[cat]
}));

return {
  title: { text: 'Tỷ Trọng ${valCol} Theo ${catCol}', textStyle: { color: '#0f172a', fontSize: 13, fontWeight: 700 } },
  tooltip: { trigger: 'item', formatter: '{b}: {c} ({d}%)' },
  legend: { bottom: '5%', textStyle: { color: '#64748b' } },
  series: [
    {
      name: '${catCol}',
      type: 'pie',
      radius: ['45%', '70%'],
      avoidLabelOverlap: true,
      itemStyle: { borderRadius: 6, borderColor: '#ffffff', borderWidth: 2 },
      label: { show: false },
      data: pieData
    }
  ]
};`;
    },

    stacked(sourceKey = null) {
      const cols = this._getCols(sourceKey);
      const catCol = cols[0] || 'Tháng';
      const m1 = cols[3] || 'Doanh Thu';
      const m2 = cols[4] || 'Lợi Nhuận';

      return `const categories = data.map(r => r['${catCol}']);
const series1 = data.map(r => Number(r['${m1}']) || 0);
const series2 = data.map(r => Number(r['${m2}']) || 0);

return {
  title: { text: 'Cơ Cấu ${m1} & ${m2}', textStyle: { color: '#0f172a', fontSize: 13, fontWeight: 700 } },
  tooltip: { trigger: 'axis', valueFormatter: v => new Intl.NumberFormat('vi-VN').format(v) },
  legend: { textStyle: { color: '#64748b' } },
  grid: { left: '4%', right: '4%', bottom: '10%', top: '16%', containLabel: true },
  xAxis: { type: 'category', data: categories, axisLabel: { color: '#64748b' } },
  yAxis: { type: 'value', splitLine: { lineStyle: { color: 'rgba(0,0,0,0.06)' } }, axisLabel: { color: '#64748b' } },
  series: [
    { name: '${m1}', type: 'bar', stack: 'total', data: series1, itemStyle: { color: '#0284c7' } },
    { name: '${m2}', type: 'bar', stack: 'total', data: series2, itemStyle: { color: '#10b981' } }
  ]
};`;
    },

    dual_axes(sourceKey = null) {
      const cols = this._getCols(sourceKey);
      const catCol = cols[0] || 'Tháng';
      const barCol = cols[3] || 'Doanh Thu';
      const lineCol = cols[4] || 'Lợi Nhuận';

      return `const categories = data.map(r => r['${catCol}']);
const barVals = data.map(r => Number(r['${barCol}']) || 0);
const lineVals = data.map(r => Number(r['${lineCol}']) || 0);

return {
  title: { text: '${barCol} (Cột) & ${lineCol} (Đường)', textStyle: { color: '#0f172a', fontSize: 13, fontWeight: 700 } },
  tooltip: { trigger: 'axis' },
  legend: { textStyle: { color: '#64748b' } },
  grid: { left: '4%', right: '5%', bottom: '10%', top: '18%', containLabel: true },
  xAxis: { type: 'category', data: categories, axisLabel: { color: '#64748b' } },
  yAxis: [
    { type: 'value', name: '${barCol}', splitLine: { lineStyle: { color: 'rgba(0,0,0,0.06)' } }, axisLabel: { color: '#64748b' } },
    { type: 'value', name: '${lineCol}', splitLine: { show: false }, axisLabel: { color: '#10b981' } }
  ],
  series: [
    { name: '${barCol}', type: 'bar', data: barVals, itemStyle: { color: '#0284c7', borderRadius: [4, 4, 0, 0] } },
    { name: '${lineCol}', type: 'line', yAxisIndex: 1, data: lineVals, itemStyle: { color: '#10b981' }, lineStyle: { width: 3 } }
  ]
};`;
    },

    gauge() {
      return `// Đồng hồ đo chỉ số KPI
const totalRevenue = data.reduce((sum, r) => sum + (Number(r['Doanh Thu']) || 0), 0);
const target = 1000000000;
const percent = Math.min(100, Math.round((totalRevenue / target) * 100));

return {
  series: [
    {
      type: 'gauge',
      progress: { show: true, width: 14, itemStyle: { color: '#0284c7' } },
      axisLine: { lineStyle: { width: 14, color: [[1, '#e2e8f0']] } },
      axisTick: { show: false },
      splitLine: { length: 8, lineStyle: { width: 2, color: '#cbd5e1' } },
      axisLabel: { distance: 18, color: '#64748b', fontSize: 10 },
      detail: { valueAnimation: true, fontSize: 24, offsetCenter: [0, '40%'], color: '#0f172a', formatter: '{value}%' },
      title: { fontSize: 12, offsetCenter: [0, '70%'], color: '#64748b' },
      data: [{ value: percent, name: 'Chỉ Tiêu KPI' }]
    }
  ]
};`;
    },

    radar(sourceKey = null) {
      const cols = this._getCols(sourceKey);
      const catCol = cols[0] || 'Tháng';
      const m1 = cols[3] || 'Doanh Thu';

      return `const indicators = data.slice(0, 6).map(r => ({ name: String(r['${catCol}']), max: 200000000 }));
const vals = data.slice(0, 6).map(r => Number(r['${m1}']) || 0);

return {
  tooltip: {},
  radar: {
    indicator: indicators,
    axisName: { color: '#64748b' },
    splitLine: { lineStyle: { color: 'rgba(0,0,0,0.08)' } },
    splitArea: { show: false }
  },
  series: [
    {
      name: '${m1}',
      type: 'radar',
      data: [{ value: vals, name: '${m1}' }],
      itemStyle: { color: '#7c3aed' },
      areaStyle: { opacity: 0.3 }
    }
  ]
};`;
    },

    funnel(sourceKey = null) {
      const cols = this._getCols(sourceKey);
      const catCol = cols[0] || 'Tháng';
      const valCol = cols[3] || 'Doanh Thu';

      return `const funnelData = data.slice(0, 5).map(r => ({
  name: String(r['${catCol}']),
  value: Number(r['${valCol}']) || 0
})).sort((a, b) => b.value - a.value);

return {
  title: { text: 'Biểu Đồ Phễu (Funnel)', textStyle: { color: '#0f172a', fontSize: 13, fontWeight: 700 } },
  tooltip: { trigger: 'item', formatter: '{b}: {c}' },
  series: [
    {
      name: '${catCol}',
      type: 'funnel',
      left: '10%',
      top: 50,
      bottom: 20,
      width: '80%',
      sort: 'descending',
      gap: 2,
      label: { show: true, position: 'inside' },
      data: funnelData
    }
  ]
};`;
    },

    scatter(sourceKey = null) {
      const cols = this._getCols(sourceKey);
      const xCol = cols[5] || cols[3] || 'Chi Phí';
      const yCol = cols[3] || cols[4] || 'Doanh Thu';

      return `const scatterData = data.map(r => [
  Number(r['${xCol}']) || 0,
  Number(r['${yCol}']) || 0
]);

return {
  title: { text: 'Tương Quan ${xCol} vs ${yCol}', textStyle: { color: '#0f172a', fontSize: 13, fontWeight: 700 } },
  tooltip: { trigger: 'item', formatter: p => '${xCol}: ' + p.value[0] + '<br/>${yCol}: ' + p.value[1] },
  grid: { left: '5%', right: '5%', bottom: '10%', top: '16%', containLabel: true },
  xAxis: { type: 'value', name: '${xCol}', splitLine: { lineStyle: { color: 'rgba(0,0,0,0.06)' } }, axisLabel: { color: '#64748b' } },
  yAxis: { type: 'value', name: '${yCol}', splitLine: { lineStyle: { color: 'rgba(0,0,0,0.06)' } }, axisLabel: { color: '#64748b' } },
  series: [
    {
      symbolSize: 16,
      data: scatterData,
      type: 'scatter',
      itemStyle: { color: '#f59e0b' }
    }
  ]
};`;
    },

    interactive_buttons(sourceKey = null) {
      const cols = this._getCols(sourceKey);
      const catCol = cols.includes('Tháng') ? 'Tháng' : (cols[1] || cols[0] || 'Danh Mục');
      const m1 = cols.includes('Doanh Thu') ? 'Doanh Thu' : (cols[3] || 'Doanh Thu');
      const m2 = cols.includes('Lợi Nhuận') ? 'Lợi Nhuận' : (cols[4] || 'Lợi Nhuận');

      return `// 💡 BIỂU ĐỒ TƯƠNG TÁC CÓ PHÍM CHỨC NĂNG (ECharts Graphic Buttons)
// Bạn có thể viết các nút bấm vào 'graphic', gán onclick, gọi myChart.setOption(...) hoặc setFilter(...) để tương tác!

const categories = data.map(r => r['${catCol}']);
const val1 = data.map(r => Number(r['${m1}']) || 0);
const val2 = data.map(r => Number(r['${m2}']) || 0);

let activeMetric = '${m1}';
let chartType = 'bar';

const option = {
  title: {
    text: 'Biểu Đồ Tương Tác: ' + activeMetric,
    subtext: 'Bấm các nút phím chức năng trên biểu đồ để chuyển đổi góc nhìn',
    textStyle: { color: '#0f172a', fontSize: 14, fontWeight: 700 },
    subtextStyle: { color: '#64748b', fontSize: 11 }
  },
  tooltip: { trigger: 'axis', valueFormatter: v => new Intl.NumberFormat('vi-VN').format(v) },
  grid: { left: '4%', right: '4%', bottom: '10%', top: '22%', containLabel: true },
  xAxis: {
    type: 'category',
    data: categories,
    axisLabel: { color: '#64748b' }
  },
  yAxis: {
    type: 'value',
    splitLine: { lineStyle: { color: 'rgba(0,0,0,0.06)' } },
    axisLabel: { color: '#64748b' }
  },
  series: [
    {
      name: activeMetric,
      type: chartType,
      data: val1,
      itemStyle: {
        color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
          { offset: 0, color: '#0284c7' },
          { offset: 1, color: '#0369a1' }
        ]),
        borderRadius: [4, 4, 0, 0]
      }
    }
  ],
  // 🔘 CÁC PHÍM CHỨC NĂNG GRAPHIC TƯƠNG TÁC TRỰC TIẾP TRÊN BIỂU ĐỒ:
  graphic: [
    {
      type: 'group',
      right: 16,
      top: 10,
      children: [
        // Phím 1: Xem ${m1}
        {
          type: 'rect',
          z: 100,
          left: 0,
          top: 0,
          shape: { width: 100, height: 28, r: 6 },
          style: { fill: '#0284c7', cursor: 'pointer' },
          textContent: {
            type: 'text',
            style: { text: '📊 ${m1}', fill: '#ffffff', font: 'bold 11px sans-serif' }
          },
          textConfig: { position: 'inside' },
          onclick: function () {
            activeMetric = '${m1}';
            myChart.setOption({
              title: { text: 'Biểu Đồ: ' + activeMetric },
              series: [{ name: activeMetric, type: chartType, data: val1, itemStyle: { color: '#0284c7' } }]
            });
          }
        },
        // Phím 2: Xem ${m2}
        {
          type: 'rect',
          z: 100,
          left: 110,
          top: 0,
          shape: { width: 100, height: 28, r: 6 },
          style: { fill: '#10b981', cursor: 'pointer' },
          textContent: {
            type: 'text',
            style: { text: '💰 ${m2}', fill: '#ffffff', font: 'bold 11px sans-serif' }
          },
          textConfig: { position: 'inside' },
          onclick: function () {
            activeMetric = '${m2}';
            myChart.setOption({
              title: { text: 'Biểu Đồ: ' + activeMetric },
              series: [{ name: activeMetric, type: chartType, data: val2, itemStyle: { color: '#10b981' } }]
            });
          }
        },
        // Phím 3: Đổi Cột / Đường (Toggle Bar / Line)
        {
          type: 'rect',
          z: 100,
          left: 220,
          top: 0,
          shape: { width: 95, height: 28, r: 6 },
          style: { fill: '#f1f5f9', stroke: '#cbd5e1', lineWidth: 1, cursor: 'pointer' },
          textContent: {
            type: 'text',
            style: { text: '📈 Đổi Cột/Đường', fill: '#0f172a', font: '11px sans-serif' }
          },
          textConfig: { position: 'inside' },
          onclick: function () {
            chartType = chartType === 'bar' ? 'line' : 'bar';
            const curData = activeMetric === '${m1}' ? val1 : val2;
            const curColor = activeMetric === '${m1}' ? '#0284c7' : '#10b981';
            myChart.setOption({
              series: [{ type: chartType, smooth: true, data: curData, itemStyle: { color: curColor }, lineStyle: { width: 3 } }]
            });
          }
        }
      ]
    }
  ]
};

return option;`;
    }
  };

  // ===========================================================================
  // 5. CHART CODE STUDIO MODAL (CHUẨN 100% GIAO DIỆN ECHARTS VISUALIZER)
  // ===========================================================================
  const CORPORATE_CHART_PRESETS = [
    {
      id: 'interactive-buttons',
      name: 'Biểu Đồ Tương Tác Phím Chức Năng (Graphic Buttons)',
      icon: '🔘',
      category: 'Tương Tác Nâng Cao',
      desc: 'Tích hợp các nút bấm chuyển đổi chỉ số (Doanh thu/Lợi nhuận) và kiểu hiển thị (Cột/Đường) trực tiếp ngay trên mặt biểu đồ.',
      generator: () => ChartTemplates.interactive_buttons()
    },
    {
      id: 'bar-gradient',
      name: 'Cột Gradient Doanh Thu Doanh Nghiệp',
      icon: '📊',
      category: 'Doanh Số & Tài Chính',
      desc: 'Biểu đồ cột dải màu gradient từ xanh dương đến cyan, bo góc mềm mại, hiển thị giá trị tiền tệ VNĐ rõ nét.',
      generator: () => ChartTemplates.bar()
    },
    {
      id: 'horizontal-bar',
      name: 'Thanh Ngang Xếp Hạng Top / So Sánh',
      icon: '📊',
      category: 'Xếp Hạng & So Sánh',
      desc: 'Biểu đồ thanh ngang sắp xếp trực quan, dễ dàng so sánh hiệu suất giữa các chi nhánh hoặc danh mục.',
      generator: () => ChartTemplates.horizontal_bar()
    },
    {
      id: 'line-trend',
      name: 'Đường Xu Hướng Thời Gian (Line Chart)',
      icon: '📈',
      category: 'Xu Hướng Thời Gian',
      desc: 'Đường xu hướng mượt mà với điểm neo rõ ràng, thích hợp theo dõi tiến độ doanh thu hoặc biến động tháng/quý.',
      generator: () => ChartTemplates.line()
    },
    {
      id: 'area-smooth',
      name: 'Vùng Gradient Mềm Mại (Smooth Area)',
      icon: '📉',
      category: 'Xu Hướng & Vùng',
      desc: 'Đường cong mượt mà có diện tích phủ gradient đổ bóng trong suốt, tạo hiệu ứng thị giác hiện đại, chuyên nghiệp.',
      generator: () => ChartTemplates.area()
    },
    {
      id: 'pie-donut',
      name: 'Bánh Donut % Cơ Cấu Tỷ Trọng',
      icon: '🥧',
      category: 'Cơ Cấu & Tỷ Trọng',
      desc: 'Biểu đồ tròn khuyết tâm (Donut) hiện đại với bo góc cánh cung, có legend chi tiết và nhãn phần trăm %.',
      generator: () => ChartTemplates.pie()
    },
    {
      id: 'stacked-bar',
      name: 'Cột Chồng Phân Bổ (Stacked Bar)',
      icon: '📊',
      category: 'Phân Tích Cơ Cấu',
      desc: 'Cột chồng thể hiện tổng thể và chi tiết từng cấu phần cấu thành trong một chu kỳ phân tích.',
      generator: () => ChartTemplates.stacked()
    },
    {
      id: 'dual-axes',
      name: 'Trục Kép Cột & Đường (Dual Axes Combo)',
      icon: '📈',
      category: 'Tài Chính Nâng Cao',
      desc: 'Hỗn hợp cột biểu diễn Doanh thu (trục Y1) và đường biểu diễn Tỷ suất lợi nhuận / Tăng trưởng % (trục Y2).',
      generator: () => ChartTemplates.dual_axes()
    },
    {
      id: 'gauge-kpi',
      name: 'Đồng Hồ Đo Tốc Độ / KPI Tiến Độ',
      icon: '⏱️',
      category: 'Chỉ Số KPI',
      desc: 'Mặt đồng hồ đo tỷ lệ hoàn thành mục tiêu kế hoạch theo % với kim chỉ và dải phân vùng màu tiêu chuẩn.',
      generator: () => ChartTemplates.gauge()
    },
    {
      id: 'radar-kpi',
      name: 'Radar Đa Chiều Năng Lực & Hiệu Suất',
      icon: '🕸️',
      category: 'Đánh Giá Toàn Diện',
      desc: 'Biểu đồ mạng nhện đa chiều so sánh các tiêu chí đánh giá phòng ban, nhân sự hoặc chất lượng vận hành.',
      generator: () => ChartTemplates.radar()
    },
    {
      id: 'funnel-sales',
      name: 'Phễu Chuyển Đổi Bán Hàng (Sales Funnel)',
      icon: '🔻',
      category: 'Chuyển Đổi & Quy Trình',
      desc: 'Biểu đồ hình phễu trực quan hóa từng giai đoạn từ Tiếp cận -> Quan tâm -> Báo giá -> Chốt hợp đồng.',
      generator: () => ChartTemplates.funnel()
    },
    {
      id: 'scatter-plot',
      name: 'Phân Tán Tương Quan (Scatter Plot)',
      icon: '⚪',
      category: 'Thống Kê Chuyên Sâu',
      desc: 'Phân tích mối tương quan giữa hai biến số lượng đơn và doanh thu để nhận diện phân khúc khách hàng.',
      generator: () => ChartTemplates.scatter()
    }
  ];

  const ChartStudioModal = {
    editingItemId: null,
    editorInitialized: false,
    activeSheetKey: null,

    open(itemId = null) {
      this.editingItemId = itemId;
      const modal = document.getElementById('chartEditorModal');
      this.initEditorEvents();

      const hideHeaderCheckbox = document.getElementById('ceHideHeaderCheckbox');

      if (itemId) {
        const item = DashboardState.present.items.find(i => i.id === itemId);
        if (item) {
          const titleInput = document.getElementById('ceChartTitleInput');
          if (titleInput) titleInput.value = item.name || 'Biểu Đồ ECharts';
          const ta = document.getElementById('ceCodeTextarea');
          if (ta) ta.value = item.code || '';
          if (hideHeaderCheckbox) hideHeaderCheckbox.checked = Boolean(item.hideHeader);
        }
      } else {
        const titleInput = document.getElementById('ceChartTitleInput');
        if (titleInput) titleInput.value = `Biểu Đồ ${DashboardState.present.items.length + 1}`;
        const ta = document.getElementById('ceCodeTextarea');
        if (ta) ta.value = ChartTemplates.bar();
        if (hideHeaderCheckbox) hideHeaderCheckbox.checked = false;
      }

      // Đưa tab về Code JavaScript
      const tabBtnCode = document.getElementById('ceTabBtnCode');
      const tabBtnData = document.getElementById('ceTabBtnData');
      const tabContentCode = document.getElementById('ceTabCodeContent');
      const tabContentData = document.getElementById('ceTabDataContent');
      if (tabBtnCode && tabBtnData) {
        tabBtnCode.classList.add('active');
        tabBtnData.classList.remove('active');
      }
      if (tabContentCode && tabContentData) {
        tabContentCode.classList.add('active');
        tabContentData.classList.remove('active');
      }

      updateCeCodeHighlight();
      this.renderDataStudio();
      this.runAndPreview();

      // Đảm bảo chia đều 50% - 50% và dành toàn bộ nửa bên phải của popup chỉ để hiển thị biểu đồ ECharts
      const editorPanel = document.getElementById('ceEditorPanel');
      if (editorPanel) {
        editorPanel.style.width = '50%';
      }
      const dataExplorer = document.querySelector('.ce-data-explorer-wrap');
      if (dataExplorer) {
        dataExplorer.style.display = 'none';
      }

      if (modal) modal.classList.add('open');
      setTimeout(() => {
        const previewDom = document.getElementById('cePreviewChart');
        if (previewDom) {
          const chart = echarts.getInstanceByDom(previewDom);
          if (chart) chart.resize();
        }
      }, 100);
    },

    initEditorEvents() {
      if (this.editorInitialized) return;
      this.editorInitialized = true;

      const textarea = document.getElementById('ceCodeTextarea');
      const highlight = document.getElementById('ceCodeHighlight');
      const lineNumbers = document.getElementById('ceLineNumbers');

      // 1. Resizer kéo giãn độ rộng 2 cột (Editor vs Preview)
      const resizer = document.getElementById('ceResizer');
      const editorPanel = document.getElementById('ceEditorPanel');
      const card = document.getElementById('chartEditorCard');
      if (resizer && editorPanel && card) {
        let isDragging = false;
        resizer.addEventListener('mousedown', () => {
          isDragging = true;
          resizer.classList.add('resizing');
          document.body.style.cursor = 'col-resize';
          document.body.style.userSelect = 'none';
        });
        window.addEventListener('mousemove', (e) => {
          if (!isDragging) return;
          const cardRect = card.getBoundingClientRect();
          const newWidth = e.clientX - cardRect.left;
          const minW = 340;
          const maxW = cardRect.width * 0.75;
          if (newWidth >= minW && newWidth <= maxW) {
            editorPanel.style.width = `${newWidth}px`;
            const previewDom = document.getElementById('cePreviewChart');
            if (previewDom) {
              const chart = echarts.getInstanceByDom(previewDom);
              if (chart) chart.resize();
            }
          }
        });
        window.addEventListener('mouseup', () => {
          if (isDragging) {
            isDragging = false;
            resizer.classList.remove('resizing');
            document.body.style.cursor = '';
            document.body.style.userSelect = '';
            const previewDom = document.getElementById('cePreviewChart');
            if (previewDom) {
              const chart = echarts.getInstanceByDom(previewDom);
              if (chart) chart.resize();
            }
          }
        });
      }

      // 2. Tab switching bên cột trái (Code JavaScript vs Data)
      const tabBtnCode = document.getElementById('ceTabBtnCode');
      const tabBtnData = document.getElementById('ceTabBtnData');
      const tabContentCode = document.getElementById('ceTabCodeContent');
      const tabContentData = document.getElementById('ceTabDataContent');

      if (tabBtnCode && tabBtnData) {
        tabBtnCode.addEventListener('click', () => {
          tabBtnCode.classList.add('active');
          tabBtnData.classList.remove('active');
          tabContentCode?.classList.add('active');
          tabContentData?.classList.remove('active');
        });
        tabBtnData.addEventListener('click', () => {
          tabBtnData.classList.add('active');
          tabBtnCode.classList.remove('active');
          tabContentData?.classList.add('active');
          tabContentCode?.classList.remove('active');
          this.renderDataStudio();
        });
      }

      // Đóng modal bằng nút Close, click ra ngoài màn hình backdrop hoặc phím Esc
      document.getElementById('btnCeCancel')?.addEventListener('click', () => this.close());

      const modalEl = document.getElementById('chartEditorModal');
      modalEl?.addEventListener('click', (e) => {
        if (e.target === modalEl) {
          this.close();
        }
      });

      window.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
          const m = document.getElementById('chartEditorModal');
          if (m && m.classList.contains('open')) {
            const presetModal = document.getElementById('vizPresetModal');
            const aiModal = document.getElementById('vizAiPromptModal');
            if (presetModal && presetModal.style.display === 'flex') {
              presetModal.style.display = 'none';
              return;
            }
            if (aiModal && aiModal.style.display === 'flex') {
              aiModal.style.display = 'none';
              return;
            }
            this.close();
          }
        }
      });

      // 3. Tab Tools Buttons
      document.getElementById('ceTabRunCode')?.addEventListener('click', () => this.runAndPreview());
      document.getElementById('ceTabAiPrompt')?.addEventListener('click', () => AiPromptStudio.open());
      document.getElementById('ceTabFormatCode')?.addEventListener('click', () => this.formatCode());
      document.getElementById('ceTabCopyCode')?.addEventListener('click', () => {
        const ta = document.getElementById('ceCodeTextarea');
        if (ta && ta.value) {
          navigator.clipboard.writeText(ta.value).then(() => {
            showToast('Đã sao chép toàn bộ code vào Clipboard!', 'success');
          });
        }
      });
      document.getElementById('ceTabResetCode')?.addEventListener('click', async () => {
        const proceed = await AppDialog.confirm('Bạn có chắc chắn muốn khôi phục code về mẫu ban đầu không?', 'Khôi Phục Code Mẫu');
        if (proceed) {
          const ta = document.getElementById('ceCodeTextarea');
          if (ta) ta.value = ChartTemplates.bar();
          updateCeCodeHighlight();
          this.runAndPreview();
          showToast('Đã khôi phục code mẫu ban đầu!', 'info');
        }
      });

      // 4. Modal Presets (Mẫu biểu đồ doanh nghiệp)
      document.getElementById('btnCeOpenPresets')?.addEventListener('click', () => {
        this.renderPresets();
        const presetModal = document.getElementById('vizPresetModal');
        if (presetModal) presetModal.style.display = 'flex';
      });
      document.getElementById('btnClosePresetModal')?.addEventListener('click', () => {
        const presetModal = document.getElementById('vizPresetModal');
        if (presetModal) presetModal.style.display = 'none';
      });

      if (!textarea) return;

      // Đồng bộ highlight khi gõ code
      textarea.addEventListener('input', () => {
        updateCeCodeHighlight();
      });

      // Đồng bộ cuộn giữa textarea và highlight pre & line numbers
      textarea.addEventListener('scroll', () => {
        if (highlight) {
          highlight.scrollTop = textarea.scrollTop;
          highlight.scrollLeft = textarea.scrollLeft;
        }
        if (lineNumbers) {
          lineNumbers.scrollTop = textarea.scrollTop;
        }
      });

      // Xử lý phím: Tab (2 spaces), Enter (auto indent), Ctrl+Enter (chạy code), Shift+Alt+F (format)
      textarea.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
          e.preventDefault();
          this.runAndPreview();
          return;
        }

        if (e.key === 'F' && e.shiftKey && e.altKey) {
          e.preventDefault();
          this.formatCode();
          return;
        }

        if (e.key === 'Tab') {
          e.preventDefault();
          const start = textarea.selectionStart;
          const end = textarea.selectionEnd;
          const val = textarea.value;

          if (e.shiftKey) {
            // Un-indent 2 spaces
            const beforeStart = val.lastIndexOf('\n', start - 1) + 1;
            if (val.slice(beforeStart, beforeStart + 2) === '  ') {
              textarea.value = val.slice(0, beforeStart) + val.slice(beforeStart + 2);
              textarea.selectionStart = Math.max(beforeStart, start - 2);
              textarea.selectionEnd = Math.max(beforeStart, end - 2);
              updateCeCodeHighlight();
            }
          } else {
            // Indent 2 spaces
            textarea.value = val.substring(0, start) + '  ' + val.substring(end);
            textarea.selectionStart = textarea.selectionEnd = start + 2;
            updateCeCodeHighlight();
          }
          return;
        }

        if (e.key === 'Enter') {
          const start = textarea.selectionStart;
          const val = textarea.value;
          const lineStart = val.lastIndexOf('\n', start - 1) + 1;
          const currentLine = val.substring(lineStart, start);
          const indentMatch = currentLine.match(/^(\s+)/);
          let indent = indentMatch ? indentMatch[1] : '';

          if (currentLine.trim().endsWith('{') || currentLine.trim().endsWith('[') || currentLine.trim().endsWith('(')) {
            indent += '  ';
          }

          if (indent) {
            e.preventDefault();
            const insert = '\n' + indent;
            textarea.value = val.substring(0, start) + insert + val.substring(start);
            textarea.selectionStart = textarea.selectionEnd = start + insert.length;
            updateCeCodeHighlight();
          }
        }
      });
    },

    renderPresets() {
      const grid = document.getElementById('vizPresetGrid');
      if (!grid) return;
      grid.innerHTML = CORPORATE_CHART_PRESETS.map(p => `
        <div class="preset-card" data-preset-id="${p.id}">
          <div class="preset-card-header">
            <span class="preset-card-icon">${p.icon}</span>
            <span class="preset-card-title">${p.name}</span>
          </div>
          <div class="preset-tag">${p.category}</div>
          <div class="preset-card-desc">${p.desc}</div>
        </div>
      `).join('');

      grid.querySelectorAll('.preset-card').forEach(card => {
        card.onclick = () => {
          const id = card.getAttribute('data-preset-id');
          const preset = CORPORATE_CHART_PRESETS.find(p => p.id === id);
          if (preset) {
            const ta = document.getElementById('ceCodeTextarea');
            if (ta) {
              ta.value = preset.generator();
              updateCeCodeHighlight();
              ChartStudioModal.runAndPreview();
              const presetModal = document.getElementById('vizPresetModal');
              if (presetModal) presetModal.style.display = 'none';
              showToast(`Đã áp dụng mẫu biểu đồ: "${preset.name}"`, 'success');
            }
          }
        };
      });
    },

    formatCode() {
      const textarea = document.getElementById('ceCodeTextarea');
      if (!textarea) return;
      const formatted = formatJSCode(textarea.value);
      textarea.value = formatted;
      updateCeCodeHighlight();
      printCeTerminal('Đã căn chỉnh thụt lề code tự động (Format JS).', 'info');
    },

    close() {
      const modal = document.getElementById('chartEditorModal');
      if (modal) modal.classList.remove('open');
    },

    runAndPreview() {
      const textarea = document.getElementById('ceCodeTextarea');
      const code = textarea ? textarea.value : '';
      const previewDom = document.getElementById('cePreviewChart');
      const renderTimeEl = document.getElementById('ceRenderTime');
      const dataCountEl = document.getElementById('ceDataCount');

      printCeTerminal('Đang biên dịch và thực thi mã ECharts...', 'info');

      let chart = null;
      if (previewDom) {
        chart = echarts.getInstanceByDom(previewDom);
        if (chart) chart.dispose();
        chart = echarts.init(previewDom, null);
      }

      const { option, error, logs, durationMs } = evaluateEChartsCode(code, null, {}, chart);

      if (logs && logs.length > 0) {
        logs.forEach(l => printCeTerminal(l.text, l.type));
      }

      if (renderTimeEl) {
        renderTimeEl.textContent = `${durationMs || 1} ms`;
      }

      if (error) {
        printCeTerminal(`Lỗi thực thi: ${error}`, 'error');
        if (dataCountEl) dataCountEl.textContent = 'Lỗi code';
      } else {
        printCeTerminal(`Biên dịch thành công trong ${durationMs || 1}ms. Biểu đồ đã hiển thị trên Live Preview.`, 'success');
        if (dataCountEl) dataCountEl.textContent = 'Sẵn sàng';

        if (chart && option && Object.keys(option).length > 0) {
          chart.setOption(option, true);
        }
      }
    },

    renderDataStudio() {
      const tabsBar = document.getElementById('ceSheetTabsBar');
      const previewBody = document.getElementById('ceSheetPreviewBody');
      const container1 = document.getElementById('cePreviewDataTable');

      const sources = DataStore.getAvailableSources();

      if (sources.length === 0) {
        const emptyMsg = `
          <div class="ce-data-empty-state">
            <div style="font-size:32px;margin-bottom:8px;">📂</div>
            <div style="font-weight:700;font-size:13px;margin-bottom:4px;color:#cbd5e1;">Chưa có dữ liệu Excel / CSV nào</div>
            <div style="font-size:11.5px;color:#94a3b8;max-width:320px;margin:0 auto;">Vui lòng nạp file Excel hoặc CSV từ thanh công cụ chính của Dashboard.</div>
          </div>
        `;
        if (tabsBar) tabsBar.innerHTML = '';
        if (previewBody) previewBody.innerHTML = emptyMsg;
        if (container1) container1.innerHTML = '<div style="color:#94a3b8;text-align:center;padding:14px;font-size:11px;">Chưa có dữ liệu Excel.</div>';
        return;
      }

      // Đảm bảo activeSheetKey hợp lệ
      if (!this.activeSheetKey || !sources.some(s => s.key === this.activeSheetKey)) {
        this.activeSheetKey = sources[0].key;
      }

      // 1. Render Sub-Tabs: TênFile [TênSheet]
      if (tabsBar) {
        tabsBar.innerHTML = sources.map(src => {
          const isActive = src.key === this.activeSheetKey;
          const rows = DataStore.getData(src.key);
          const count = rows ? rows.length : 0;
          return `
            <button type="button" class="ce-sheet-subtab ${isActive ? 'active' : ''}" data-key="${src.key}" title="${src.fileName} - Sheet: ${src.sheetName}">
              <span style="font-size:13px;">📑</span>
              <span class="ce-sheet-subtab-title">${src.fileName} <strong style="color:${isActive ? '#38bdf8' : '#e2e8f0'}">[${src.sheetName}]</strong></span>
              <span class="ce-sheet-subtab-badge">${count.toLocaleString('vi-VN')} dòng</span>
            </button>
          `;
        }).join('');

        tabsBar.querySelectorAll('.ce-sheet-subtab').forEach(btn => {
          btn.onclick = () => {
            this.activeSheetKey = btn.getAttribute('data-key');
            this.renderDataStudio();
          };
        });
      }

      // 2. Render Demo Bảng Dữ Liệu & Info Của Sheet Đang Chọn
      const activeSource = sources.find(s => s.key === this.activeSheetKey) || sources[0];
      const rows = DataStore.getData(activeSource.key) || [];
      const cols = rows.length > 0 ? Object.keys(rows[0]) : [];
      const isDefault = activeSource.key === sources[0].key;

      if (previewBody) {
        let bodyHtml = `
          <div class="ce-sheet-info-toolbar">
            <div class="ce-sheet-meta-row">
              <div class="ce-sheet-meta-badges">
                <span style="font-weight:700;color:#f1f5f9;">📄 ${activeSource.fileName}</span>
                <span style="color:#64748b;">›</span>
                <span style="font-weight:700;color:#38bdf8;">📑 Sheet: "${activeSource.sheetName}"</span>
                <span style="background:rgba(56,189,248,0.15);color:#38bdf8;padding:2px 8px;border-radius:12px;font-size:10.5px;font-weight:600;">
                  ${rows.length.toLocaleString('vi-VN')} dòng • ${cols.length} cột
                </span>
              </div>
              <div class="ce-sheet-copy-snippets">
                <span style="font-size:11px;color:#94a3b8;">Lấy dữ liệu:</span>
                ${isDefault ? `<button type="button" class="ce-copy-pill" data-copy="data" title="Mảng dữ liệu mặc định trong code"><code>data</code></button>` : ''}
                <button type="button" class="ce-copy-pill" data-copy="getSheet('${activeSource.sheetName}')" title="Click để sao chép cú pháp getSheet">
                  <code>getSheet('${activeSource.sheetName}')</code>
                </button>
                <button type="button" class="ce-copy-pill" data-copy="files['${activeSource.fileName}']['${activeSource.sheetName}']" title="Click để sao chép cú pháp lấy từ file cụ thể">
                  <code>files['${activeSource.fileName}']...</code>
                </button>
              </div>
            </div>

            <!-- Bấm vào cột để sao chép -->
            <div class="ce-sheet-cols-row">
              <span style="color:#94a3b8;font-size:11px;white-space:nowrap;">Bấm cột để chép:</span>
              <div style="display:flex;flex-wrap:wrap;gap:4px;overflow-x:auto;">
                ${cols.map(c => {
                  const escaped = c.replace(/'/g, "\\'");
                  return `<span class="ce-hint-tag" title="Click để sao chép tên cột '${escaped}'" data-copy="${escaped}" style="cursor:pointer;font-size:10.5px;padding:2px 6px;border-radius:3px;background:rgba(56,189,248,0.12);color:#7dd3fc;border:1px solid rgba(56,189,248,0.25);"><code>${c}</code></span>`;
                }).join('')}
              </div>
            </div>
          </div>

          <!-- DEMO BẢNG DỮ LIỆU -->
          <div class="ce-data-table-container">
        `;

        if (rows.length === 0) {
          bodyHtml += `<div style="text-align:center;padding:30px;color:#64748b;font-size:12px;">Sheet này không có dòng dữ liệu nào.</div>`;
        } else {
          const previewRows = rows.slice(0, 100);
          bodyHtml += `
            <table class="ce-demo-table">
              <thead>
                <tr>
                  <th class="ce-th-idx">#</th>
                  ${cols.map(c => `<th title="Cột: ${c}">${c}</th>`).join('')}
                </tr>
              </thead>
              <tbody>
                ${previewRows.map((r, idx) => `
                  <tr>
                    <td class="ce-td-idx">${idx + 1}</td>
                    ${cols.map(c => {
                      const val = r[c];
                      const isNum = typeof val === 'number' || (!isNaN(val) && val !== '' && val !== null && typeof val !== 'boolean');
                      const displayVal = (val === null || val === undefined || val === '') ? '<span class="ce-cell-empty">—</span>' : (typeof val === 'number' ? val.toLocaleString('vi-VN') : String(val));
                      return `<td class="${isNum && val !== '' ? 'ce-cell-num' : ''}" title="${String(val ?? '')}">${displayVal}</td>`;
                    }).join('')}
                  </tr>
                `).join('')}
              </tbody>
            </table>
          `;
        }

        bodyHtml += `</div>`;

        bodyHtml += `
          <div class="ce-data-table-footer">
            <span>Hiển thị ${Math.min(100, rows.length).toLocaleString('vi-VN')} / ${rows.length.toLocaleString('vi-VN')} dòng thực tế</span>
            <span>💡 Click vào tên cột để sao chép nhanh vào code ECharts</span>
          </div>
        `;

        previewBody.innerHTML = bodyHtml;

        previewBody.querySelectorAll('.ce-copy-pill, .ce-hint-tag').forEach(el => {
          el.addEventListener('click', () => {
            const txt = el.getAttribute('data-copy');
            if (txt) {
              navigator.clipboard.writeText(txt).then(() => {
                showToast(`Đã sao chép: ${txt}`, 'info');
              });
            }
          });
        });
      }
      // Khung dữ liệu Excel bên phải đã được ẩn để nhường toàn bộ không gian cho biểu đồ ECharts
    },

    updateDataPreviewTable() {
      this.renderDataStudio();
    },

    async save() {
      const title = document.getElementById('ceChartTitleInput').value.trim() || 'Biểu Đồ ECharts';
      const code = document.getElementById('ceCodeTextarea').value;

      const { error } = evaluateEChartsCode(code);
      if (error) {
        const proceed = await AppDialog.confirm(`Code đang có lỗi:\n"${error}"\nBạn có chắc chắn muốn lưu không?`, 'Cảnh Báo Lỗi Code ECharts', { isDanger: true, confirmText: 'Vẫn Lưu' });
        if (!proceed) {
          return;
        }
      }

      syncItemPositionsFromDOM();
      DashboardState.pushSnapshot();
      const hideHeader = Boolean(document.getElementById('ceHideHeaderCheckbox')?.checked);

      if (this.editingItemId) {
        const item = DashboardState.present.items.find(i => i.id === this.editingItemId);
        if (item) {
          item.name = title;
          item.code = code;
          item.hideHeader = hideHeader;
        }
      } else {
        const newId = 'chart_' + Date.now();
        const pos = getNextAvailablePosition(560, 340);

        DashboardState.present.items.push({
          id: newId,
          type: 'chart',
          name: title,
          code: code,
          hideHeader: hideHeader,
          tabId: DashboardState.present.activeTabId,
          pos: pos
        });
        DashboardState.selectedItemId = newId;
        setTimeout(() => {
          document.getElementById(`item_${newId}`)?.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }, 60);
      }

      this.close();
      DashboardState.isDirty = true;
      renderCanvas();
      triggerDebouncedAutoSave();
      showToast('Đã lưu biểu đồ thành công vào Dashboard!', 'success');
    }
  };

  // ===========================================================================
  // 6. AI SYSTEM PROMPT STUDIO (CHUẨN 100% GIAO DIỆN ECHARTS VISUALIZER)
  // ===========================================================================
  const AiPromptStudio = {
    initialized: false,

    open(defaultReq = '') {
      this.initEvents();
      const modal = document.getElementById('vizAiPromptModal');
      const fileBadge = document.getElementById('aiPromptFileBadge');
      const colsBadge = document.getElementById('aiPromptColsBadge');

      const allFiles = DataStore.getAllFilesObject();
      const fNames = Object.keys(allFiles);
      if (fNames.length > 0) {
        const firstFile = fNames[0];
        const sheets = Object.keys(allFiles[firstFile] || {});
        const firstSheet = sheets[0] || '';
        const rows = allFiles[firstFile]?.[firstSheet] || [];
        const cols = rows.length > 0 ? Object.keys(rows[0]) : [];

        if (fileBadge) {
          fileBadge.textContent = `📁 ${firstFile} • Sheet: ${firstSheet} (${rows.length.toLocaleString('vi-VN')} dòng)`;
          fileBadge.style.color = '#34d399';
        }
        if (colsBadge) {
          colsBadge.innerHTML = cols.map(c => `<span style="background:rgba(56,189,248,0.12);border:1px solid rgba(56,189,248,0.25);color:#7dd3fc;padding:2px 6px;border-radius:3px;margin:2px;display:inline-block;font-size:11px;">${c}</span>`).join(' ') || 'Không có cột nào';
        }
      } else {
        if (fileBadge) {
          fileBadge.textContent = '⚡ Dùng dữ liệu mẫu mặc định trong data';
          fileBadge.style.color = '#f59e0b';
        }
        if (colsBadge) {
          colsBadge.textContent = 'Các cột mẫu mặc định: Khu Vực, Doanh Thu, Đơn Hàng, Lợi Nhuận';
        }
      }

      const bizLogic = document.getElementById('aiBusinessLogic');
      if (bizLogic && defaultReq) bizLogic.value = defaultReq;

      this.generatePrompt();
      if (modal) modal.style.display = 'flex';
    },

    close() {
      const modal = document.getElementById('vizAiPromptModal');
      if (modal) modal.style.display = 'none';
    },

    initEvents() {
      if (this.initialized) return;
      this.initialized = true;

      document.getElementById('btnCloseAiModal')?.addEventListener('click', () => this.close());
      document.getElementById('btnGenerateAiPrompt')?.addEventListener('click', () => this.generatePrompt());
      document.getElementById('btnCopyAiPrompt')?.addEventListener('click', () => this.copyPrompt());
      document.getElementById('btnApplyAiCode')?.addEventListener('click', () => this.applyAiCodeToEditor());

      // Type Chips
      document.querySelectorAll('.ai-type-chip').forEach(chip => {
        chip.addEventListener('click', () => {
          const typeDesc = chip.getAttribute('data-type');
          const outputArea = document.getElementById('aiExpectedOutput');
          if (outputArea && typeDesc) {
            outputArea.value = typeDesc;
            this.generatePrompt();
          }
        });
      });
    },

    generatePrompt() {
      const allFiles = DataStore.getAllFilesObject();
      const fNames = Object.keys(allFiles);
      let schemaDesc = '';
      if (fNames.length > 0) {
        const firstFile = fNames[0];
        const sheets = Object.keys(allFiles[firstFile] || {});
        const firstSheet = sheets[0] || '';
        const rows = allFiles[firstFile]?.[firstSheet] || [];
        const cols = rows.length > 0 ? Object.keys(rows[0]) : [];
        schemaDesc = `File nạp: "${firstFile}", Sheet: "${firstSheet}" (${rows.length} dòng).\nDanh sách cột: [${cols.join(', ')}].`;
      } else {
        schemaDesc = `Mảng mẫu mặc định trong data: [{ "Khu Vực": "Hà Nội", "Doanh Thu": 450000000, "Đơn Hàng": 120, "Lợi Nhuận": 135000000 }, ...]`;
      }

      const inputExtra = document.getElementById('aiInputDataDesc')?.value.trim() || '';
      const bizLogic = document.getElementById('aiBusinessLogic')?.value.trim() || 'Vẽ biểu đồ phân tích phù hợp làm nổi bật các chỉ số kinh doanh quan trọng.';
      const expectedOutput = document.getElementById('aiExpectedOutput')?.value.trim() || 'Biểu đồ trực quan, định dạng tiền tệ VNĐ, màu sắc gradient hiện đại, tooltip đầy đủ.';

      const promptText = `=== SYSTEM PROMPT: CHUYÊN GIA ECHARTS 5 & JAVASCRIPT ===

BẠN LÀ CHUYÊN GIA SENIOR DATA VISUALIZATION VÀ LẬP TRÌNH APACHE ECHARTS 5 JAVASCRIPT.
HÃY VIẾT DUY NHẤT MÃ JAVASCRIPT CHUẨN XÁC ĐỂ XỬ LÝ DỮ LIỆU EXCEL VÀ RETURN VỀ OBJECT OPTION ECHARTS.

1. THÔNG TIN DỮ LIỆU ĐẦU VÀO:
- ${schemaDesc}
${inputExtra ? `- Ghi chú bổ sung: ${inputExtra}` : ''}

2. YÊU CẦU NGHIỆP VỤ & TỔNG HỢP:
${bizLogic}

3. KIỂU BIỂU ĐỒ & KẾT QUẢ ĐẦU RA MONG MUỐN:
${expectedOutput}

4. NGUYÊN TẮC BẮT BUỘC (CRITICAL CONTRACT):
- Chỉ dùng JavaScript thuần chạy trên trình duyệt.
- Dữ liệu đã có sẵn trong biến toàn cục 'data' (mảng các object của sheet).
- Trích xuất dữ liệu: const categories = data.map(r => r['Tên Cột X']);
- Luôn ép kiểu số an toàn: Number(r['Tên Cột Y']) || 0.
- Nếu cần lấy sheet khác: const rows = getSheet('TênSheetKhác');
- Kết thúc mã code bằng lệnh: return { ...cấu hình option ECharts... };
- KHÔNG bọc mã trong hàm không trả về, KHÔNG giải thích, KHÔNG viết văn bản chào hỏi.`;

      const out = document.getElementById('aiPromptOutput');
      if (out) out.value = promptText;
    },

    copyPrompt() {
      const outputBox = document.getElementById('aiPromptOutput');
      if (!outputBox || !outputBox.value) return;
      navigator.clipboard.writeText(outputBox.value).then(() => {
        showToast('Đã sao chép System Prompt vào Clipboard!', 'success');
      }).catch(() => {
        outputBox.select();
        document.execCommand('copy');
        showToast('Đã sao chép System Prompt!', 'success');
      });
    },

    applyAiCodeToEditor() {
      let code = document.getElementById('aiCodePasteArea')?.value.trim() || '';
      if (!code) {
        showToast('Vui lòng dán mã code do AI sinh ra vào ô trước!', 'error');
        return;
      }
      code = code.replace(/^```[a-zA-Z]*\n?/m, '').replace(/\n?```$/m, '').trim();

      const textarea = document.getElementById('ceCodeTextarea');
      if (textarea) textarea.value = code;

      this.close();
      updateCeCodeHighlight();
      ChartStudioModal.runAndPreview();
      showToast('Đã áp dụng mã JavaScript từ AI vào Editor và chạy thử Live Preview!', 'success');
    }
  };

  // ===========================================================================
  // 7. WIDGET CONFIGURATION MODAL (FILTER, KPI, HEADING, FRAME, SHAPE, TEXT, BUTTON)
  // ===========================================================================
  function getFontSelectOptions(selectedFont) {
    const fonts = [
      { val: 'Inter, sans-serif', label: 'Inter (Chuẩn)' },
      { val: 'Roboto, sans-serif', label: 'Roboto' },
      { val: 'Montserrat, sans-serif', label: 'Montserrat' },
      { val: 'Arial, sans-serif', label: 'Arial' },
      { val: "'Segoe UI', sans-serif", label: 'Segoe UI' },
      { val: "'Times New Roman', serif", label: 'Times New Roman' },
      { val: 'Georgia, serif', label: 'Georgia' },
      { val: "'Playfair Display', serif", label: 'Playfair Display' },
      { val: "'Courier New', monospace", label: 'Courier New' },
      { val: 'Consolas, monospace', label: 'Consolas' }
    ];
    return fonts.map(f => `<option value="${f.val}" ${selectedFont === f.val ? 'selected' : ''}>${f.label}</option>`).join('');
  }

  function buildCanvaBarHtml(prefix, font, sizeNum, color, isBold, isItalic, isUnderline, align) {
    return `
      <div class="wm-canva-bar">
        <select class="wm-canva-select" id="${prefix}FontFamily" title="Font chữ">
          ${getFontSelectOptions(font)}
        </select>
        <div style="display:flex;align-items:center;gap:2px;">
          <button type="button" class="wm-canva-btn" id="${prefix}SizeMinus" title="Giảm cỡ chữ">−</button>
          <input type="number" class="wm-canva-select" id="${prefix}FontSize" value="${sizeNum}" min="8" max="140" style="width:50px;text-align:center;padding:4px;" title="Cỡ chữ (px)" />
          <button type="button" class="wm-canva-btn" id="${prefix}SizePlus" title="Tăng cỡ chữ">+</button>
        </div>
        <input type="color" class="wm-canva-select" id="${prefix}Color" value="${color && color.startsWith('#') ? color : '#0f172a'}" style="width:36px;height:28px;padding:1px;cursor:pointer;" title="Màu chữ" />
        <button type="button" class="wm-canva-btn ${isBold ? 'active' : ''}" id="${prefix}BoldBtn" title="In đậm (Ctrl + B)"><strong>B</strong></button>
        <button type="button" class="wm-canva-btn ${isItalic ? 'active' : ''}" id="${prefix}ItalicBtn" title="In nghiêng (Ctrl + I)"><em>I</em></button>
        <button type="button" class="wm-canva-btn ${isUnderline ? 'active' : ''}" id="${prefix}UnderlineBtn" title="Gạch chân (Ctrl + U)"><u>U</u></button>
        <div style="display:flex;gap:2px;">
          <button type="button" class="wm-canva-btn ${align === 'left' ? 'active' : ''}" data-wm-align="left" title="Căn trái">⯇</button>
          <button type="button" class="wm-canva-btn ${align === 'center' ? 'active' : ''}" data-wm-align="center" title="Căn giữa">☰</button>
          <button type="button" class="wm-canva-btn ${align === 'right' ? 'active' : ''}" data-wm-align="right" title="Căn phải">⯈</button>
          <button type="button" class="wm-canva-btn ${align === 'justify' ? 'active' : ''}" data-wm-align="justify" title="Căn đều">⯎</button>
        </div>
        <input type="hidden" id="${prefix}Align" value="${align}" />
      </div>
    `;
  }

  function bindCanvaBarEvents(prefix, textInputId, previewId) {
    const textInput = document.getElementById(textInputId);
    const previewEl = document.getElementById(previewId);
    const fontEl = document.getElementById(`${prefix}FontFamily`);
    const sizeEl = document.getElementById(`${prefix}FontSize`);
    const colorEl = document.getElementById(`${prefix}Color`);
    const boldBtn = document.getElementById(`${prefix}BoldBtn`);
    const italicBtn = document.getElementById(`${prefix}ItalicBtn`);
    const underlineBtn = document.getElementById(`${prefix}UnderlineBtn`);
    const alignInput = document.getElementById(`${prefix}Align`);

    function updatePreview() {
      if (!previewEl) return;
      const text = textInput ? (textInput.value || 'Xem trước chữ...') : 'Xem trước...';
      const font = fontEl ? fontEl.value : 'Inter, sans-serif';
      const size = sizeEl ? `${sizeEl.value}px` : '16px';
      const color = colorEl ? colorEl.value : '#0f172a';
      const isBold = boldBtn?.classList.contains('active');
      const isItalic = italicBtn?.classList.contains('active');
      const isUnderline = underlineBtn?.classList.contains('active');
      const align = alignInput ? alignInput.value : 'left';

      previewEl.style.fontFamily = font;
      previewEl.style.fontSize = size;
      previewEl.style.color = color;
      previewEl.style.fontWeight = isBold ? 'bold' : 'normal';
      previewEl.style.fontStyle = isItalic ? 'italic' : 'normal';
      previewEl.style.textDecoration = isUnderline ? 'underline' : 'none';
      previewEl.style.textAlign = align;
      previewEl.textContent = text;
    }

    textInput?.addEventListener('input', updatePreview);
    fontEl?.addEventListener('change', updatePreview);
    colorEl?.addEventListener('input', updatePreview);

    document.getElementById(`${prefix}SizeMinus`)?.addEventListener('click', () => {
      if (sizeEl) {
        sizeEl.value = Math.max(8, (parseInt(sizeEl.value, 10) || 16) - 2);
        updatePreview();
      }
    });
    document.getElementById(`${prefix}SizePlus`)?.addEventListener('click', () => {
      if (sizeEl) {
        sizeEl.value = Math.min(140, (parseInt(sizeEl.value, 10) || 16) + 2);
        updatePreview();
      }
    });
    sizeEl?.addEventListener('change', updatePreview);

    boldBtn?.addEventListener('click', () => {
      boldBtn.classList.toggle('active');
      updatePreview();
    });
    italicBtn?.addEventListener('click', () => {
      italicBtn.classList.toggle('active');
      updatePreview();
    });
    underlineBtn?.addEventListener('click', () => {
      underlineBtn.classList.toggle('active');
      updatePreview();
    });

    document.querySelectorAll(`#widgetModalBody [data-wm-align]`).forEach(btn => {
      btn.addEventListener('click', () => {
        const al = btn.dataset.wmAlign;
        if (alignInput) alignInput.value = al;
        document.querySelectorAll(`#widgetModalBody [data-wm-align]`).forEach(b => b.classList.toggle('active', b === btn));
        updatePreview();
      });
    });

    // Shortcuts Ctrl + B, I, U in text input
    textInput?.addEventListener('keydown', (e) => {
      if (e.ctrlKey || e.metaKey) {
        const key = e.key.toLowerCase();
        if (key === 'b') {
          e.preventDefault();
          boldBtn?.classList.toggle('active');
          updatePreview();
        } else if (key === 'i') {
          e.preventDefault();
          italicBtn?.classList.toggle('active');
          updatePreview();
        } else if (key === 'u') {
          e.preventDefault();
          underlineBtn?.classList.toggle('active');
          updatePreview();
        }
      }
    });

    updatePreview();
  }

  const WidgetConfigModal = {
    currentType: 'filter',
    editingItemId: null,

    open(type, itemId = null) {
      this.currentType = type;
      this.editingItemId = itemId;
      const modal = document.getElementById('widgetConfigModal');
      const titleEl = document.getElementById('widgetModalTitle');
      const bodyEl = document.getElementById('widgetModalBody');
      const availableSources = DataStore.getAvailableSources();
      let existing = null;
      if (itemId) {
        existing = DashboardState.present.items.find(i => i.id === itemId);
      }

      let curSourceKey = existing?.dataSourceKey || '';
      if (!curSourceKey && existing?.filterColumn) {
        const found = availableSources.find(s => DataStore.getColumns(s.key).includes(existing.filterColumn));
        if (found) curSourceKey = found.key;
      }
      if (!curSourceKey && availableSources.length > 0) {
        curSourceKey = availableSources[0].key;
      }

      let cols = DataStore.getColumns(curSourceKey);
      if (cols.length === 0 && availableSources.length > 0) {
        curSourceKey = availableSources[0].key;
        cols = DataStore.getColumns(curSourceKey);
      }

      if (bodyEl) bodyEl.scrollTop = 0;

      if (type === 'filter') {
        const modalBox = modal.querySelector('.db-simple-box');
        if (modalBox) {
          modalBox.style.maxWidth = '760px';
          modalBox.style.maxHeight = '90vh';
        }

        titleEl.innerHTML = `<span style="display:inline-flex;align-items:center;gap:8px;"><span>🔍</span><span>${existing ? 'Chỉnh Sửa Bộ Lọc Dữ Liệu' : 'Thêm Bộ Lọc Dữ Liệu (Filter Studio)'}</span></span>`;
        const curCol = existing?.filterColumn && cols.includes(existing.filterColumn) ? existing.filterColumn : (cols[0] || '');
        const allDataRows = DataStore.getData(curSourceKey, curCol);
        const isColDate = isProbableDateColumn(curCol, allDataRows);
        const curType = existing?.filterType || (isColDate ? 'date' : 'range');
        const curParent = existing?.parentFilterId || '';
        let curOp = existing?.operator;
        if (!curOp) {
          if (curType === 'range') curOp = 'between';
          else if (curType === 'multiselect') curOp = 'in';
          else if (curType === 'date') curOp = 'date_range';
          else if (curType === 'text') curOp = 'contains';
          else if (curType === 'comparison') curOp = '>';
          else curOp = '=';
        }

        const otherFilters = DashboardState.present.items.filter(i => i.type === 'filter' && (!existing || i.id !== existing.id));

        bodyEl.innerHTML = `
          <div class="filter-studio-wrapper">
            <!-- 1. CHỌN NGUỒN DỮ LIỆU & SHEET -->
            <div class="filter-studio-sec">
              <div class="filter-studio-sec-header">
                <div class="filter-studio-sec-title">
                  <span class="filter-studio-step-badge">1</span>
                  <span>Nguồn Dữ Liệu & Sheet Cần Lọc</span>
                </div>
                <span style="font-size:11px;color:#0284c7;font-weight:600;">Hỗ trợ lọc mọi Sheet của file Excel</span>
              </div>

              <div style="margin-top:6px;">
                <div style="position:relative;display:flex;align-items:center;">
                  <span style="position:absolute;left:12px;font-size:14px;pointer-events:none;">📑</span>
                  <select class="ce-select" id="wmFilterSourceSelect" style="width:100%;padding-left:36px;height:38px;font-weight:600;font-size:13px;border:1.5px solid #cbd5e1;background:#fff;border-radius:8px;">
                    ${availableSources.map(s => `<option value="${escapeHtml(s.key)}" ${s.key === curSourceKey ? 'selected' : ''}>📄 ${escapeHtml(s.label)}</option>`).join('')}
                  </select>
                </div>
              </div>
            </div>

            <!-- 2. CHỌN CỘT DỮ LIỆU NGUỒN -->
            <div class="filter-studio-sec">
              <div class="filter-studio-sec-header">
                <div class="filter-studio-sec-title">
                  <span class="filter-studio-step-badge">2</span>
                  <span>Cột Dữ Liệu Cần Lọc (Excel Column)</span>
                </div>
                <div id="wmFilterTypeMetaBadge" class="filter-col-meta-badge"></div>
              </div>

              <div class="filter-column-selector-row">
                <div class="filter-column-select-box">
                  <span class="filter-column-select-icon">📊</span>
                  <select class="ce-select" id="wmFilterColSelect">
                    ${cols.map(c => `<option value="${escapeHtml(c)}" ${c === curCol ? 'selected' : ''}>${escapeHtml(c)}</option>`).join('')}
                  </select>
                </div>
                <div style="display:flex;align-items:center;justify-content:space-between;background:#f8fafc;padding:8px 12px;border-radius:8px;border:1px solid #e2e8f0;">
                  <span style="font-size:11.5px;color:#64748b;">Mẫu dữ liệu: <strong id="wmFilterUniqueCount" style="color:#0f172a;">--</strong></span>
                  <span style="font-size:11px;color:#0284c7;font-weight:600;" id="wmFilterColAdvice"></span>
                </div>
              </div>

              <!-- Dải giá trị mẫu thực tế -->
              <div style="margin-top:10px;">
                <div style="font-size:11px;color:#64748b;margin-bottom:5px;font-weight:600;display:flex;align-items:center;justify-content:space-between;">
                  <span>Mẫu giá trị thực tế trong cột:</span>
                  <span style="font-size:10.5px;color:#94a3b8;">(10 giá trị đầu)</span>
                </div>
                <div class="filter-sample-chips" id="wmFilterSampleBadges"></div>
              </div>
            </div>

            <!-- 3. LỰA CHỌN KIỂU BỘ LỌC TRỰC QUAN (6 CARDS) -->
            <div class="filter-studio-sec">
              <div class="filter-studio-sec-header">
                <div class="filter-studio-sec-title">
                  <span class="filter-studio-step-badge">3</span>
                  <span>Chọn Kiểu Bộ Lọc Trực Quan (Interactive Filter Type)</span>
                </div>
                <span style="font-size:11px;color:#64748b;font-weight:600;">Chuẩn BI & Analytics</span>
              </div>

              <!-- Input ẩn lưu giá trị kiểu lọc để save() đồng bộ -->
              <input type="hidden" id="wmFilterTypeSelect" value="${curType}" />

              <div class="filter-type-grid" id="wmFilterTypeCardsGrid">
                <div class="filter-type-card ${curType === 'date' ? 'active' : ''}" data-type="date">
                  <div class="filter-type-card-top">
                    <div class="filter-type-icon">📅</div>
                    <div class="filter-type-check">✓</div>
                  </div>
                  <div class="filter-type-name">Khoảng Ngày</div>
                  <div class="filter-type-desc">Lọc từ ngày... đến ngày bằng lịch Date Picker</div>
                </div>

                <div class="filter-type-card ${curType === 'multiselect' ? 'active' : ''}" data-type="multiselect">
                  <div class="filter-type-card-top">
                    <div class="filter-type-icon">☑️</div>
                    <div class="filter-type-check">✓</div>
                  </div>
                  <div class="filter-type-name">Chọn Nhiều Mục</div>
                  <div class="filter-type-desc">Dropdown có checkbox chọn nhiều giá trị cùng lúc</div>
                </div>

                <div class="filter-type-card ${curType === 'select' ? 'active' : ''}" data-type="select">
                  <div class="filter-type-card-top">
                    <div class="filter-type-icon">📋</div>
                    <div class="filter-type-check">✓</div>
                  </div>
                  <div class="filter-type-name">Chọn 1 Mục</div>
                  <div class="filter-type-desc">Menu thả xuống chọn duy nhất 1 giá trị</div>
                </div>

                <div class="filter-type-card ${curType === 'range' ? 'active' : ''}" data-type="range">
                  <div class="filter-type-card-top">
                    <div class="filter-type-icon">↔️</div>
                    <div class="filter-type-check">✓</div>
                  </div>
                  <div class="filter-type-name">Khoảng Số (Min - Max)</div>
                  <div class="filter-type-desc">Nhập từ số... đến số (doanh thu, số lượng...)</div>
                </div>

                <div class="filter-type-card ${curType === 'comparison' ? 'active' : ''}" data-type="comparison">
                  <div class="filter-type-card-top">
                    <div class="filter-type-icon">🔢</div>
                    <div class="filter-type-check">✓</div>
                  </div>
                  <div class="filter-type-name">So Sánh 1 Mốc</div>
                  <div class="filter-type-desc">Toán tử >, >=, <, <=, =, ≠ so với 1 giá trị</div>
                </div>

                <div class="filter-type-card ${curType === 'text' ? 'active' : ''}" data-type="text">
                  <div class="filter-type-card-top">
                    <div class="filter-type-icon">🔤</div>
                    <div class="filter-type-check">✓</div>
                  </div>
                  <div class="filter-type-name">Tìm Kiếm Chữ</div>
                  <div class="filter-type-desc">Ô tìm kiếm từ khóa (chứa, bắt đầu, bằng...)</div>
                </div>
              </div>
            </div>

            <!-- 4. THIẾT LẬP NHÃN & TOÁN TỬ -->
            <div class="filter-studio-sec">
              <div class="filter-studio-sec-header">
                <div class="filter-studio-sec-title">
                  <span class="filter-studio-step-badge">4</span>
                  <span>Thiết Lập Nhãn & Liên Kết Bộ Lọc</span>
                </div>
              </div>

              <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;">
                <div class="ai-form-group">
                  <label class="ai-label">Tiêu Đề Nhãn Bộ Lọc:</label>
                  <div style="display:flex;gap:6px;">
                    <input type="text" class="ce-input" id="wmFilterTitleInput" value="${escapeHtml(existing ? existing.name : (curCol ? 'Lọc: ' + curCol : 'Bộ Lọc'))}" style="flex:1;" />
                    <button type="button" class="db-btn-tool" id="wmBtnAutoTitle" title="Tự động đặt theo tên cột" style="padding:0 10px;font-size:11.5px;background:#f8fafc;white-space:nowrap;">🔄 Đặt Lại</button>
                  </div>
                </div>

                <div class="ai-form-group" id="wmFilterOpGroup">
                  <label class="ai-label">Toán Tử Lọc Mặc Định:</label>
                  <select class="ce-select" id="wmFilterOpSelect" style="width:100%;"></select>
                </div>
              </div>

              <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:12px;align-items:flex-end;">
                <div class="ai-form-group">
                  <label class="ai-label">Bộ Lọc Cha (Cascading / Liên Kết Cha - Con):</label>
                  <select class="ce-select" id="wmFilterParentSelect" style="width:100%;">
                    <option value="">-- Không có (Bộ lọc độc lập) --</option>
                    ${otherFilters.map(f => `<option value="${f.id}" ${f.id === curParent ? 'selected' : ''}>🔗 ${escapeHtml(f.name || f.filterColumn)} (Cột: ${escapeHtml(f.filterColumn)})</option>`).join('')}
                  </select>
                  <div style="font-size:11px;color:#64748b;margin-top:4px;">💡 Danh sách giá trị sẽ tự động thu hẹp theo lựa chọn ở bộ lọc cha.</div>
                </div>

                <div class="ai-form-group">
                  <label style="display:inline-flex;align-items:center;gap:8px;font-size:12.5px;color:#0f172a;cursor:pointer;padding:9px 12px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;user-select:none;width:100%;">
                    <input type="checkbox" id="wmHideHeaderCheckbox" ${existing?.hideHeader ? 'checked' : ''} style="accent-color:var(--db-accent, #0284c7);width:16px;height:16px;cursor:pointer;" />
                    <span style="font-weight:600;">Ẩn thanh tiêu đề (Chỉ hiện ô lọc)</span>
                  </label>
                </div>
              </div>
            </div>

            <!-- 5. KHUNG XEM TRƯỚC TRỰC QUAN TRÊN CANVAS (LIVE PREVIEW) -->
            <div class="filter-live-preview-box">
              <div class="filter-live-preview-header">
                <div class="filter-live-preview-title">
                  <span>👁️ Xem Trước Thẻ Bộ Lọc Trên Dashboard</span>
                </div>
                <span style="font-size:10.5px;background:#e0f2fe;color:#0369a1;padding:2px 8px;border-radius:999px;font-weight:700;">Live Canvas Mockup</span>
              </div>

              <div class="filter-preview-card-mockup" id="wmFilterLiveMockup"></div>
            </div>
          </div>
        `;

        const updateModalOps = (type, chosenOp) => {
          const opSel = document.getElementById('wmFilterOpSelect');
          if (!opSel) return;
          if (type === 'range') {
            opSel.innerHTML = `
              <option value="between" ${chosenOp === 'between' ? 'selected' : ''}>↔️ Trong Khoảng [Từ số (Min) ➔ Đến số (Max)]</option>
            `;
          } else if (type === 'comparison') {
            opSel.innerHTML = `
              <option value="=" ${chosenOp === '=' ? 'selected' : ''}>= Bằng giá trị</option>
              <option value=">" ${chosenOp === '>' ? 'selected' : ''}>&gt; Lớn hơn giá trị</option>
              <option value=">=" ${chosenOp === '>=' ? 'selected' : ''}>&ge; Lớn hơn hoặc bằng giá trị</option>
              <option value="<" ${chosenOp === '<' ? 'selected' : ''}>&lt; Bé hơn giá trị</option>
              <option value="<=" ${chosenOp === '<=' ? 'selected' : ''}>&le; Bé hơn hoặc bằng giá trị</option>
              <option value="!=" ${chosenOp === '!=' ? 'selected' : ''}>≠ Khác giá trị</option>
            `;
          } else if (type === 'multiselect') {
            opSel.innerHTML = `
              <option value="in" ${chosenOp === 'in' ? 'selected' : ''}>☑️ Thuộc danh sách đã chọn (Nhiều mục)</option>
            `;
          } else if (type === 'date') {
            opSel.innerHTML = `
              <option value="date_range" ${chosenOp === 'date_range' ? 'selected' : ''}>📅 Khoảng ngày (Từ ngày ➔ Đến ngày với Date Picker)</option>
              <option value="date_eq" ${chosenOp === 'date_eq' ? 'selected' : ''}>🎯 Đúng ngày (= Date Picker)</option>
              <option value="date_gte" ${chosenOp === 'date_gte' ? 'selected' : ''}>⏩ Từ ngày trở đi (&ge; Date Picker)</option>
              <option value="date_lte" ${chosenOp === 'date_lte' ? 'selected' : ''}>⏪ Đến ngày trở về trước (&le; Date Picker)</option>
            `;
          } else if (type === 'text') {
            opSel.innerHTML = `
              <option value="contains" ${chosenOp === 'contains' ? 'selected' : ''}>Chứa từ khóa (Tìm kiếm)</option>
              <option value="starts_with" ${chosenOp === 'starts_with' ? 'selected' : ''}>Bắt đầu bằng...</option>
              <option value="ends_with" ${chosenOp === 'ends_with' ? 'selected' : ''}>Kết thúc bằng...</option>
              <option value="not_contains" ${chosenOp === 'not_contains' ? 'selected' : ''}>Không chứa từ khóa</option>
              <option value="=" ${chosenOp === '=' ? 'selected' : ''}>= Bằng chính xác toàn bộ</option>
            `;
          } else {
            opSel.innerHTML = `
              <option value="=" ${chosenOp === '=' ? 'selected' : ''}>= Bằng (Chọn 1 từ danh sách Excel)</option>
            `;
          }
        };

        const renderFilterLiveMockup = (col, type, name, vals) => {
          const mockupEl = document.getElementById('wmFilterLiveMockup');
          if (!mockupEl) return;
          const hideH = document.getElementById('wmHideHeaderCheckbox')?.checked;
          const displayTitle = name || `Lọc: ${col}`;
          
          let headerHtml = hideH ? '' : `
            <div class="filter-preview-mockup-label">
              <span>${escapeHtml(displayTitle)}</span>
              <span style="font-size:10.5px;color:#64748b;font-weight:normal;">[Cột: ${escapeHtml(col)}]</span>
            </div>
          `;

          let controlHtml = '';
          if (type === 'date') {
            controlHtml = `
              <div style="display:flex;gap:6px;align-items:center;">
                <input type="date" value="2026-01-01" class="fps-input" style="flex:1;height:34px;font-size:11.5px;background:#f8fafc;" disabled />
                <span style="color:#64748b;font-size:11px;">➔</span>
                <input type="date" value="2026-12-31" class="fps-input" style="flex:1;height:34px;font-size:11.5px;background:#f8fafc;" disabled />
              </div>
            `;
          } else if (type === 'multiselect') {
            controlHtml = `
              <div style="position:relative;border:1.5px solid #cbd5e1;background:#f8fafc;padding:7px 12px;border-radius:6px;display:flex;align-items:center;justify-content:space-between;cursor:default;">
                <span style="font-size:12px;color:#0f172a;font-weight:600;">☑️ Đã chọn tất cả (${vals.length || 5} mục)...</span>
                <span style="font-size:10px;color:#64748b;">▼</span>
              </div>
            `;
          } else if (type === 'range') {
            controlHtml = `
              <div style="display:flex;gap:6px;align-items:center;">
                <input type="text" placeholder="Từ số (Min)..." class="fps-input" style="flex:1;height:34px;font-size:11.5px;background:#f8fafc;" disabled />
                <span style="color:#64748b;font-size:11px;">-</span>
                <input type="text" placeholder="Đến số (Max)..." class="fps-input" style="flex:1;height:34px;font-size:11.5px;background:#f8fafc;" disabled />
              </div>
            `;
          } else if (type === 'comparison') {
            controlHtml = `
              <div style="display:flex;gap:6px;align-items:center;">
                <select class="fps-select" style="width:85px;height:34px;font-size:11.5px;background:#f8fafc;" disabled><option>&gt; Lớn hơn</option></select>
                <input type="text" placeholder="Nhập giá trị mốc..." class="fps-input" style="flex:1;height:34px;font-size:11.5px;background:#f8fafc;" disabled />
              </div>
            `;
          } else if (type === 'text') {
            controlHtml = `
              <div style="position:relative;">
                <input type="text" placeholder="🔍 Nhập từ khóa tìm kiếm trong cột ${escapeHtml(col)}..." class="fps-input" style="width:100%;height:34px;font-size:11.5px;background:#f8fafc;" disabled />
              </div>
            `;
          } else {
            // select
            controlHtml = `
              <select class="fps-select" style="width:100%;height:34px;font-size:12px;background:#f8fafc;" disabled>
                <option>-- Chọn giá trị [Tất cả] --</option>
                ${vals.slice(0, 3).map(v => `<option>${escapeHtml(v)}</option>`).join('')}
              </select>
            `;
          }

          mockupEl.innerHTML = headerHtml + controlHtml;
        };

        setTimeout(() => {
          const sourceSel = document.getElementById('wmFilterSourceSelect');
          const colSel = document.getElementById('wmFilterColSelect');
          const typeHiddenInp = document.getElementById('wmFilterTypeSelect');
          const titleInp = document.getElementById('wmFilterTitleInput');
          const badgesEl = document.getElementById('wmFilterSampleBadges');
          const metaBadge = document.getElementById('wmFilterTypeMetaBadge');
          const uniqueCountEl = document.getElementById('wmFilterUniqueCount');
          const adviceEl = document.getElementById('wmFilterColAdvice');
          const cardsGrid = document.getElementById('wmFilterTypeCardsGrid');
          const autoTitleBtn = document.getElementById('wmBtnAutoTitle');
          const hideHCheck = document.getElementById('wmHideHeaderCheckbox');

          const refreshColumnInsights = (selectedCol, customSourceKey = null) => {
            const curSrc = customSourceKey || sourceSel?.value || curSourceKey;
            const rows = DataStore.getData(curSrc, selectedCol);
            const isDate = isProbableDateColumn(selectedCol, rows);
            const vals = Array.from(new Set(rows.map(r => r[selectedCol]).filter(v => v !== null && v !== undefined && v !== ''))).slice(0, 10);
            const totalUnique = new Set(rows.map(r => r[selectedCol])).size;

            if (uniqueCountEl) uniqueCountEl.textContent = `${totalUnique.toLocaleString()} giá trị`;

            if (isDate) {
              if (metaBadge) {
                metaBadge.className = 'filter-col-meta-badge type-date';
                metaBadge.innerHTML = '📅 Cột Ngày Tháng';
              }
              if (adviceEl) adviceEl.innerHTML = '💡 Khuyên dùng: <strong>Khoảng Ngày (Date)</strong>';
            } else {
              const isNumeric = vals.some(v => !isNaN(Number(v)) && String(v).trim() !== '');
              if (isNumeric) {
                if (metaBadge) {
                  metaBadge.className = 'filter-col-meta-badge type-number';
                  metaBadge.innerHTML = '🔢 Cột Dữ Liệu Số';
                }
                if (adviceEl) adviceEl.innerHTML = '💡 Khuyên dùng: <strong>Khoảng Số hoặc So Sánh</strong>';
              } else {
                if (metaBadge) {
                  metaBadge.className = 'filter-col-meta-badge type-text';
                  metaBadge.innerHTML = '🔤 Cột Chuỗi Phân Loại';
                }
                if (adviceEl) adviceEl.innerHTML = '💡 Khuyên dùng: <strong>Chọn Nhiều Mục hoặc Dropdown</strong>';
              }
            }

            if (badgesEl) {
              badgesEl.innerHTML = vals.map(v => `<span class="filter-sample-pill">${escapeHtml(v)}</span>`).join('') || '<span style="color:#94a3b8;font-size:11px;">Không có giá trị</span>';
            }

            renderFilterLiveMockup(selectedCol, typeHiddenInp.value, titleInp.value, vals);
          };

          // Khởi tạo trạng thái ban đầu
          updateModalOps(curType, curOp);
          refreshColumnInsights(curCol, curSourceKey);

          // Sự kiện đổi Nguồn Dữ Liệu & Sheet
          sourceSel?.addEventListener('change', () => {
            const nextSourceKey = sourceSel.value;
            const nextCols = DataStore.getColumns(nextSourceKey);
            if (colSel) {
              colSel.innerHTML = nextCols.map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');
              if (nextCols.length > 0) {
                colSel.value = nextCols[0];
                if (titleInp && (!existing || titleInp.value.startsWith('Lọc: '))) {
                  titleInp.value = `Lọc: ${nextCols[0]}`;
                }
                const nextRows = DataStore.getData(nextSourceKey, nextCols[0]);
                if (isProbableDateColumn(nextCols[0], nextRows)) {
                  cardsGrid?.querySelectorAll('.filter-type-card').forEach(c => {
                    c.classList.toggle('active', c.dataset.type === 'date');
                  });
                  typeHiddenInp.value = 'date';
                  updateModalOps('date', 'date_range');
                }
                refreshColumnInsights(nextCols[0], nextSourceKey);
              } else {
                if (badgesEl) badgesEl.innerHTML = '<span style="color:#94a3b8;font-size:11px;">Sheet không có dữ liệu cột</span>';
                if (uniqueCountEl) uniqueCountEl.textContent = '0 giá trị';
              }
            }
          });

          // Sự kiện bấm chọn Card loại bộ lọc
          cardsGrid?.querySelectorAll('.filter-type-card').forEach(card => {
            card.addEventListener('click', () => {
              cardsGrid.querySelectorAll('.filter-type-card').forEach(c => c.classList.remove('active'));
              card.classList.add('active');
              const nextType = card.dataset.type;
              typeHiddenInp.value = nextType;

              let defOp = '=';
              if (nextType === 'range') defOp = 'between';
              else if (nextType === 'multiselect') defOp = 'in';
              else if (nextType === 'date') defOp = 'date_range';
              else if (nextType === 'text') defOp = 'contains';
              else if (nextType === 'comparison') defOp = '>';
              updateModalOps(nextType, defOp);

              const rows = DataStore.getData(sourceSel?.value, colSel.value);
              const vals = Array.from(new Set(rows.map(r => r[colSel.value]).filter(v => v !== null && v !== undefined && v !== ''))).slice(0, 10);
              renderFilterLiveMockup(colSel.value, nextType, titleInp.value, vals);
            });
          });

          // Sự kiện đổi cột
          colSel?.addEventListener('change', () => {
            const selectedCol = colSel.value;
            if (titleInp && (!existing || titleInp.value.startsWith('Lọc: '))) {
              titleInp.value = `Lọc: ${selectedCol}`;
            }
            const rows = DataStore.getData(sourceSel?.value, selectedCol);
            if (isProbableDateColumn(selectedCol, rows)) {
              // Tự động chuyển card sang date
              cardsGrid?.querySelectorAll('.filter-type-card').forEach(c => {
                c.classList.toggle('active', c.dataset.type === 'date');
              });
              typeHiddenInp.value = 'date';
              updateModalOps('date', 'date_range');
            }
            refreshColumnInsights(selectedCol, sourceSel?.value);
          });

          // Sự kiện bấm nút Đặt lại tên theo cột
          autoTitleBtn?.addEventListener('click', () => {
            if (titleInp && colSel) {
              titleInp.value = `Lọc: ${colSel.value}`;
              const rows = DataStore.getData(sourceSel?.value, colSel.value);
              const vals = Array.from(new Set(rows.map(r => r[colSel.value]).filter(v => v !== null && v !== undefined && v !== ''))).slice(0, 10);
              renderFilterLiveMockup(colSel.value, typeHiddenInp.value, titleInp.value, vals);
            }
          });

          // Cập nhật live preview khi gõ tên hoặc toggle ẩn header
          titleInp?.addEventListener('input', () => {
            const rows = DataStore.getData(sourceSel?.value, colSel.value);
            const vals = Array.from(new Set(rows.map(r => r[colSel.value]).filter(v => v !== null && v !== undefined && v !== ''))).slice(0, 10);
            renderFilterLiveMockup(colSel.value, typeHiddenInp.value, titleInp.value, vals);
          });

          hideHCheck?.addEventListener('change', () => {
            const rows = DataStore.getData(sourceSel?.value, colSel.value);
            const vals = Array.from(new Set(rows.map(r => r[colSel.value]).filter(v => v !== null && v !== undefined && v !== ''))).slice(0, 10);
            renderFilterLiveMockup(colSel.value, typeHiddenInp.value, titleInp.value, vals);
          });
        }, 30);
      } else if (type === 'kpi') {
        titleEl.textContent = existing ? 'Chỉnh Sửa Thẻ KPI' : 'Thêm Chỉ Số Đo Lường (KPI)';
        bodyEl.innerHTML = `
          <div class="ai-form-group">
            <label class="ai-label">Chọn Cột Số Liệu:</label>
            <select class="ce-select" id="wmKpiColSelect" style="width:100%;">
              ${cols.map(c => `<option value="${c}" ${existing && existing.metricColumn === c ? 'selected' : ''}>${c}</option>`).join('')}
            </select>
          </div>
          <div class="ai-form-group">
            <label class="ai-label">Phương Thức Tính Toán:</label>
            <select class="ce-select" id="wmKpiAggSelect" style="width:100%;">
              <option value="sum" ${existing && existing.aggregation === 'sum' ? 'selected' : ''}>Tổng Cộng (Sum)</option>
              <option value="avg" ${existing && existing.aggregation === 'avg' ? 'selected' : ''}>Trung Bình (Average)</option>
              <option value="count" ${existing && existing.aggregation === 'count' ? 'selected' : ''}>Đếm Dòng (Count)</option>
              <option value="max" ${existing && existing.aggregation === 'max' ? 'selected' : ''}>Lớn Nhất (Max)</option>
              <option value="min" ${existing && existing.aggregation === 'min' ? 'selected' : ''}>Nhỏ Nhất (Min)</option>
            </select>
          </div>
          <div class="ai-form-group">
            <label class="ai-label">Tên Chỉ Số KPI:</label>
            <input type="text" class="ce-input" id="wmKpiTitleInput" value="${existing ? existing.name : 'Tổng ' + (cols[0] || 'Chỉ Số')}" />
          </div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;">
            <div class="ai-form-group">
              <label class="ai-label">Tiền Tố (Prefix):</label>
              <input type="text" class="ce-input" id="wmKpiPrefixInput" placeholder="Ví dụ: $" value="${existing ? (existing.prefix || '') : ''}" />
            </div>
            <div class="ai-form-group">
              <label class="ai-label">Hậu Tố (Suffix):</label>
              <input type="text" class="ce-input" id="wmKpiSuffixInput" placeholder="Ví dụ: đ, VND, %" value="${existing ? (existing.suffix || '') : ' đ'}" />
            </div>
          </div>
        `;
      } else if (type === 'heading') {
        titleEl.textContent = existing ? 'Chỉnh Sửa Tiêu Đề' : 'Thêm Tiêu Đề / Heading';
        const curText = existing ? (existing.text !== undefined ? existing.text : existing.name) : 'Báo Cáo Phân Tích Hoạt Động';
        const curFont = existing?.fontFamily || 'Inter, sans-serif';
        const curSizeNum = parseInt(existing?.fontSize || '24', 10) || 24;
        const curColor = existing?.color || '#0f172a';
        const isBold = existing ? (existing.fontWeight === 'bold' || existing.fontWeight === '700' || existing.fontWeight === 800) : true;
        const isItalic = existing ? (existing.fontStyle === 'italic') : false;
        const isUnderline = existing ? (existing.textDecoration === 'underline') : false;
        const curAlign = existing?.textAlign || 'left';

        bodyEl.innerHTML = `
          <div class="ai-form-group">
            <label class="ai-label">Nội Dung Tiêu Đề:</label>
            <input type="text" class="ce-input" id="wmHeadingTextInput" value="${escapeHtml(curText)}" style="font-size:14px;padding:8px 10px;" />
          </div>
          <div class="ai-form-group">
            <label class="ai-label">Định Dạng Chữ (Canva Style - Phím tắt Ctrl + B, I, U):</label>
            ${buildCanvaBarHtml('wmHeading', curFont, curSizeNum, curColor, isBold, isItalic, isUnderline, curAlign)}
          </div>
          <div class="ai-form-group">
            <label class="ai-label">Xem Trước Trực Tiếp (Live Preview):</label>
            <div class="wm-canva-preview-box" id="wmHeadingPreviewBox"></div>
          </div>
        `;
      } else if (type === 'text') {
        titleEl.textContent = existing ? 'Chỉnh Sửa Đoạn Văn Bản' : 'Thêm Đoạn Văn Bản / Ghi Chú';
        const curText = existing ? (existing.text !== undefined ? existing.text : existing.name) : 'Ghi chú số liệu phân tích: Báo cáo cập nhật theo thời gian thực từ file Excel nguồn.';
        const curFont = existing?.fontFamily || 'Inter, sans-serif';
        const curSizeNum = parseInt(existing?.fontSize || '14', 10) || 14;
        const curColor = existing?.color || '#0f172a';
        const isBold = existing ? (existing.fontWeight === 'bold' || existing.fontWeight === '700' || existing.fontWeight === 800) : false;
        const isItalic = existing ? (existing.fontStyle === 'italic') : false;
        const isUnderline = existing ? (existing.textDecoration === 'underline') : false;
        const curAlign = existing?.textAlign || 'left';

        bodyEl.innerHTML = `
          <div class="ai-form-group">
            <label class="ai-label">Nội Dung Văn Bản / Ghi Chú:</label>
            <textarea class="ce-textarea" id="wmTextContent" rows="4" style="height:90px;font-size:13px;padding:8px 10px;">${escapeHtml(curText)}</textarea>
          </div>
          <div class="ai-form-group">
            <label class="ai-label">Định Dạng Chữ (Canva Style - Phím tắt Ctrl + B, I, U):</label>
            ${buildCanvaBarHtml('wmText', curFont, curSizeNum, curColor, isBold, isItalic, isUnderline, curAlign)}
          </div>
          <div class="ai-form-group">
            <label class="ai-label">Xem Trước Trực Tiếp (Live Preview):</label>
            <div class="wm-canva-preview-box" id="wmTextPreviewBox"></div>
          </div>
        `;
      } else if (type === 'shape') {
        titleEl.textContent = existing ? 'Chỉnh Sửa Hình Khối' : 'Cấu Hình Hình Khối (Canva / Figma)';
        bodyEl.innerHTML = `
          <div class="ai-form-group">
            <label class="ai-label">Loại Hình Khối:</label>
            <select class="ce-select" id="wmShapeTypeSelect" style="width:100%;">
              <option value="rectangle" ${!existing || existing.shapeType === 'rectangle' ? 'selected' : ''}>Hình Chữ Nhật / Card Khối</option>
              <option value="circle" ${existing && existing.shapeType === 'circle' ? 'selected' : ''}>Hình Tròn / Elip</option>
              <option value="line" ${existing && existing.shapeType === 'line' ? 'selected' : ''}>Đường Kẻ Phân Cách (Line)</option>
            </select>
          </div>
          <div class="ai-form-group">
            <label class="ai-label">Tên Hình Khối:</label>
            <input type="text" class="ce-input" id="wmShapeNameInput" value="${existing ? existing.name : 'Hình Khối'}" />
          </div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;">
            <div class="ai-form-group">
              <label class="ai-label">Màu Nền (Background):</label>
              <input type="color" class="ce-input" id="wmShapeBgColor" value="${existing?.bgColor && existing.bgColor.startsWith('#') ? existing.bgColor : '#ffffff'}" style="height:38px;cursor:pointer;padding:2px;" />
            </div>
            <div class="ai-form-group">
              <label class="ai-label">Màu Viền (Border):</label>
              <input type="color" class="ce-input" id="wmShapeBorderColor" value="${existing?.borderColor && existing.borderColor.startsWith('#') ? existing.borderColor : '#cbd5e1'}" style="height:38px;cursor:pointer;padding:2px;" />
            </div>
          </div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;">
            <div class="ai-form-group">
              <label class="ai-label">Độ Dày Viền (px):</label>
              <input type="number" class="ce-input" id="wmShapeBorderWidth" min="0" max="10" value="${existing?.borderWidth !== undefined ? existing.borderWidth : 1}" />
            </div>
            <div class="ai-form-group">
              <label class="ai-label">Bo Góc (Border Radius px):</label>
              <input type="number" class="ce-input" id="wmShapeRadius" min="0" max="100" value="${existing?.borderRadius !== undefined ? existing.borderRadius : 8}" />
            </div>
          </div>
        `;
      } else if (type === 'button') {
        titleEl.textContent = existing ? 'Chỉnh Sửa Nút Bấm Tương Tác' : 'Thêm Nút Bấm Tương Tác (Action Button)';
        const tabs = DashboardState.present.tabs || [];
        bodyEl.innerHTML = `
          <div class="ai-form-group">
            <label class="ai-label">Nhãn Nút Bấm (Button Text):</label>
            <input type="text" class="ce-input" id="wmBtnNameInput" value="${existing ? existing.name : '👉 Xem Chi Tiết'}" />
          </div>
          <div class="ai-form-group">
            <label class="ai-label">Hành Động Khi Nhấp (Action):</label>
            <select class="ce-select" id="wmBtnActionSelect" style="width:100%;">
              <option value="tab" ${!existing || existing.btnAction === 'tab' ? 'selected' : ''}>Chuyển Sang Tab Khác</option>
              <option value="url" ${existing && existing.btnAction === 'url' ? 'selected' : ''}>Mở Đường Dẫn Liên Kết (URL)</option>
            </select>
          </div>
          <div class="ai-form-group" id="wmBtnTabGroup">
            <label class="ai-label">Chọn Tab Đích:</label>
            <select class="ce-select" id="wmBtnTabSelect" style="width:100%;">
              ${tabs.map(t => `<option value="${t.id}" ${existing && existing.btnTarget === t.id ? 'selected' : ''}>${t.name}</option>`).join('')}
            </select>
          </div>
          <div class="ai-form-group" id="wmBtnUrlGroup" style="display:none;">
            <label class="ai-label">Nhập Đường Dẫn URL (https://...):</label>
            <input type="text" class="ce-input" id="wmBtnUrlInput" placeholder="https://..." value="${existing?.btnAction === 'url' ? (existing.btnTarget || '') : 'https://'}" />
          </div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;">
            <div class="ai-form-group">
              <label class="ai-label">Màu Nền Nút:</label>
              <input type="color" class="ce-input" id="wmBtnBgColor" value="${existing?.btnBg && existing.btnBg.startsWith('#') ? existing.btnBg : '#004ea8'}" style="height:38px;cursor:pointer;padding:2px;" />
            </div>
            <div class="ai-form-group">
              <label class="ai-label">Màu Chữ Nút:</label>
              <input type="color" class="ce-input" id="wmBtnTextColor" value="${existing?.btnColor && existing.btnColor.startsWith('#') ? existing.btnColor : '#ffffff'}" style="height:38px;cursor:pointer;padding:2px;" />
            </div>
          </div>
        `;
        setTimeout(() => {
          const actionSelect = document.getElementById('wmBtnActionSelect');
          const tabGroup = document.getElementById('wmBtnTabGroup');
          const urlGroup = document.getElementById('wmBtnUrlGroup');
          function updateBtnActionUI() {
            if (actionSelect.value === 'url') {
              tabGroup.style.display = 'none';
              urlGroup.style.display = 'block';
            } else {
              tabGroup.style.display = 'block';
              urlGroup.style.display = 'none';
            }
          }
          actionSelect?.addEventListener('change', updateBtnActionUI);
          updateBtnActionUI();
        }, 20);
      } else if (type === 'frame') {
        titleEl.textContent = existing ? 'Chỉnh Sửa Khung Phân Vùng' : 'Thêm Khung Phân Vùng (Frame)';
        bodyEl.innerHTML = `
          <div class="ai-form-group">
            <label class="ai-label">Tên Khung Phân Vùng:</label>
            <input type="text" class="ce-input" id="wmFrameNameInput" value="${existing ? existing.name : 'Khu Vực Phân Tích'}" />
          </div>
        `;
      }

      // Thêm tùy chọn Ẩn Header cho mọi loại widget
      const hideHeaderHtml = `
        <div class="ai-form-group" style="margin-top:8px;padding-top:10px;border-top:1px dashed var(--db-card-border);">
          <label style="display:flex;align-items:center;gap:8px;font-size:12px;font-weight:600;color:var(--db-text-main);cursor:pointer;user-select:none;" title="Tắt thanh tiêu đề của widget trên Canvas để chỉ hiển thị nội dung bên trong">
            <input type="checkbox" id="wmHideHeaderCheckbox" ${existing?.hideHeader ? 'checked' : ''} style="cursor:pointer;" />
            <span>Ẩn thanh tiêu đề (Header) của widget (Chỉ hiển thị nội dung bên trong)</span>
          </label>
        </div>
      `;
      bodyEl.insertAdjacentHTML('beforeend', hideHeaderHtml);

      if (modal) modal.classList.add('open');

      setTimeout(() => {
        if (type === 'heading') {
          bindCanvaBarEvents('wmHeading', 'wmHeadingTextInput', 'wmHeadingPreviewBox');
        } else if (type === 'text') {
          bindCanvaBarEvents('wmText', 'wmTextContent', 'wmTextPreviewBox');
        }
      }, 20);
    },

    close() {
      const modal = document.getElementById('widgetConfigModal');
      if (modal) {
        modal.classList.remove('open');
        const modalBox = modal.querySelector('.db-simple-box');
        if (modalBox) {
          modalBox.style.maxWidth = '540px';
          modalBox.style.maxHeight = '90vh';
        }
      }
    },

    save() {
      syncItemPositionsFromDOM();
      DashboardState.pushSnapshot();
      const hideHeader = Boolean(document.getElementById('wmHideHeaderCheckbox')?.checked);

      if (this.currentType === 'filter') {
        const sourceKey = document.getElementById('wmFilterSourceSelect')?.value || '';
        const col = document.getElementById('wmFilterColSelect').value;
        const fType = document.getElementById('wmFilterTypeSelect')?.value || 'range';
        const parentId = document.getElementById('wmFilterParentSelect')?.value || '';
        let op = document.getElementById('wmFilterOpSelect')?.value;
        if (!op) {
          if (fType === 'range') op = 'between';
          else if (fType === 'multiselect') op = 'in';
          else if (fType === 'date') op = 'date_range';
          else if (fType === 'text') op = 'contains';
          else op = '=';
        }
        const name = document.getElementById('wmFilterTitleInput').value.trim() || `Lọc: ${col}`;
        if (this.editingItemId) {
          const item = DashboardState.present.items.find(i => i.id === this.editingItemId);
          if (item) {
            item.dataSourceKey = sourceKey;
            item.filterColumn = col;
            item.filterType = fType;
            item.parentFilterId = parentId;
            item.operator = op;
            item.name = name;
            item.hideHeader = hideHeader;
          }
        } else {
          const newId = 'filter_' + Date.now();
          const initW = (fType === 'range' || fType === 'date') ? 340 : 290;
          DashboardState.present.items.push({
            id: newId,
            type: 'filter',
            name: name,
            dataSourceKey: sourceKey,
            filterColumn: col,
            filterType: fType,
            parentFilterId: parentId,
            operator: op,
            hideHeader: hideHeader,
            tabId: DashboardState.present.activeTabId,
            pos: getNextAvailablePosition(initW, 105)
          });
          DashboardState.selectedItemId = newId;
        }
      } else if (this.currentType === 'kpi') {
        const col = document.getElementById('wmKpiColSelect').value;
        const agg = document.getElementById('wmKpiAggSelect').value;
        const name = document.getElementById('wmKpiTitleInput').value.trim() || `Tổng ${col}`;
        const prefix = document.getElementById('wmKpiPrefixInput').value;
        const suffix = document.getElementById('wmKpiSuffixInput').value;
        if (this.editingItemId) {
          const item = DashboardState.present.items.find(i => i.id === this.editingItemId);
          if (item) { item.metricColumn = col; item.aggregation = agg; item.name = name; item.prefix = prefix; item.suffix = suffix; item.hideHeader = hideHeader; }
        } else {
          const newId = 'kpi_' + Date.now();
          DashboardState.present.items.push({
            id: newId,
            type: 'kpi',
            name: name,
            metricColumn: col,
            aggregation: agg,
            prefix: prefix,
            suffix: suffix,
            hideHeader: hideHeader,
            tabId: DashboardState.present.activeTabId,
            pos: getNextAvailablePosition(260, 90)
          });
          DashboardState.selectedItemId = newId;
        }
      } else if (this.currentType === 'heading') {
        const text = document.getElementById('wmHeadingTextInput')?.value.trim() || 'Tiêu Đề';
        const fontSize = (document.getElementById('wmHeadingFontSize')?.value ? document.getElementById('wmHeadingFontSize').value + 'px' : '24px');
        const fontFamily = document.getElementById('wmHeadingFontFamily')?.value || 'Inter, sans-serif';
        const fontWeight = document.getElementById('wmHeadingBoldBtn')?.classList.contains('active') ? 'bold' : 'normal';
        const fontStyle = document.getElementById('wmHeadingItalicBtn')?.classList.contains('active') ? 'italic' : 'normal';
        const textDecoration = document.getElementById('wmHeadingUnderlineBtn')?.classList.contains('active') ? 'underline' : 'none';
        const color = document.getElementById('wmHeadingColor')?.value || '#0f172a';
        const textAlign = document.getElementById('wmHeadingAlign')?.value || 'left';

        if (this.editingItemId) {
          const item = DashboardState.present.items.find(i => i.id === this.editingItemId);
          if (item) {
            item.text = text;
            item.name = text;
            item.fontSize = fontSize;
            item.fontFamily = fontFamily;
            item.fontWeight = fontWeight;
            item.fontStyle = fontStyle;
            item.textDecoration = textDecoration;
            item.color = color;
            item.textAlign = textAlign;
            item.hideHeader = hideHeader;
          }
        } else {
          const newId = 'heading_' + Date.now();
          DashboardState.present.items.push({
            id: newId,
            type: 'heading',
            name: text,
            text: text,
            fontSize: fontSize,
            fontFamily: fontFamily,
            fontWeight: fontWeight,
            fontStyle: fontStyle,
            textDecoration: textDecoration,
            color: color,
            textAlign: textAlign,
            hideHeader: hideHeader,
            tabId: DashboardState.present.activeTabId,
            pos: getNextAvailablePosition(560, 70)
          });
          DashboardState.selectedItemId = newId;
        }
      } else if (this.currentType === 'text') {
        const text = document.getElementById('wmTextContent')?.value || 'Ghi chú';
        const fontSize = (document.getElementById('wmTextFontSize')?.value ? document.getElementById('wmTextFontSize').value + 'px' : '14px');
        const fontFamily = document.getElementById('wmTextFontFamily')?.value || 'Inter, sans-serif';
        const fontWeight = document.getElementById('wmTextBoldBtn')?.classList.contains('active') ? 'bold' : 'normal';
        const fontStyle = document.getElementById('wmTextItalicBtn')?.classList.contains('active') ? 'italic' : 'normal';
        const textDecoration = document.getElementById('wmTextUnderlineBtn')?.classList.contains('active') ? 'underline' : 'none';
        const color = document.getElementById('wmTextColor')?.value || '#0f172a';
        const textAlign = document.getElementById('wmTextAlign')?.value || 'left';

        if (this.editingItemId) {
          const item = DashboardState.present.items.find(i => i.id === this.editingItemId);
          if (item) {
            item.text = text;
            item.name = text.substring(0, 20);
            item.fontSize = fontSize;
            item.fontFamily = fontFamily;
            item.fontWeight = fontWeight;
            item.fontStyle = fontStyle;
            item.textDecoration = textDecoration;
            item.color = color;
            item.textAlign = textAlign;
            item.hideHeader = hideHeader;
          }
        } else {
          const newId = 'text_' + Date.now();
          DashboardState.present.items.push({
            id: newId,
            type: 'text',
            name: text.substring(0, 20),
            text: text,
            fontSize: fontSize,
            fontFamily: fontFamily,
            fontWeight: fontWeight,
            fontStyle: fontStyle,
            textDecoration: textDecoration,
            color: color,
            textAlign: textAlign,
            hideHeader: hideHeader,
            tabId: DashboardState.present.activeTabId,
            pos: getNextAvailablePosition(380, 120)
          });
          DashboardState.selectedItemId = newId;
        }
      } else if (this.currentType === 'shape') {
        const shapeType = document.getElementById('wmShapeTypeSelect').value;
        const name = document.getElementById('wmShapeNameInput').value.trim() || 'Hình Khối';
        const bgColor = document.getElementById('wmShapeBgColor').value;
        const borderColor = document.getElementById('wmShapeBorderColor').value;
        const borderWidth = Number(document.getElementById('wmShapeBorderWidth').value) || 1;
        const borderRadius = Number(document.getElementById('wmShapeRadius').value) || 0;
        if (this.editingItemId) {
          const item = DashboardState.present.items.find(i => i.id === this.editingItemId);
          if (item) { item.shapeType = shapeType; item.name = name; item.bgColor = bgColor; item.borderColor = borderColor; item.borderWidth = borderWidth; item.borderRadius = borderRadius; item.hideHeader = hideHeader; }
        } else {
          const newId = 'shape_' + Date.now();
          DashboardState.present.items.push({
            id: newId,
            type: 'shape',
            shapeType: shapeType,
            name: name,
            bgColor: bgColor,
            borderColor: borderColor,
            borderWidth: borderWidth,
            borderRadius: borderRadius,
            hideHeader: hideHeader,
            tabId: DashboardState.present.activeTabId,
            pos: getNextAvailablePosition(320, 180)
          });
          DashboardState.selectedItemId = newId;
        }
      } else if (this.currentType === 'button') {
        const name = document.getElementById('wmBtnNameInput').value.trim() || 'Nút Bấm';
        const btnAction = document.getElementById('wmBtnActionSelect').value;
        const btnTarget = btnAction === 'url' ? document.getElementById('wmBtnUrlInput').value.trim() : document.getElementById('wmBtnTabSelect').value;
        const btnBg = document.getElementById('wmBtnBgColor').value;
        const btnColor = document.getElementById('wmBtnTextColor').value;
        if (this.editingItemId) {
          const item = DashboardState.present.items.find(i => i.id === this.editingItemId);
          if (item) { item.name = name; item.btnAction = btnAction; item.btnTarget = btnTarget; item.btnBg = btnBg; item.btnColor = btnColor; item.hideHeader = hideHeader; }
        } else {
          const newId = 'btn_' + Date.now();
          DashboardState.present.items.push({
            id: newId,
            type: 'button',
            name: name,
            btnAction: btnAction,
            btnTarget: btnTarget,
            btnBg: btnBg,
            btnColor: btnColor,
            hideHeader: hideHeader,
            tabId: DashboardState.present.activeTabId,
            pos: getNextAvailablePosition(200, 50)
          });
          DashboardState.selectedItemId = newId;
        }
      } else if (this.currentType === 'frame') {
        const name = document.getElementById('wmFrameNameInput').value.trim() || 'Khung Phân Vùng';
        if (this.editingItemId) {
          const item = DashboardState.present.items.find(i => i.id === this.editingItemId);
          if (item) { item.name = name; item.hideHeader = hideHeader; }
        } else {
          const newId = 'frame_' + Date.now();
          DashboardState.present.items.push({
            id: newId,
            type: 'frame',
            name: name,
            hideHeader: hideHeader,
            tabId: DashboardState.present.activeTabId,
            pos: getNextAvailablePosition(1160, 420)
          });
          DashboardState.selectedItemId = newId;
        }
      }

      const justAddedId = !this.editingItemId ? DashboardState.selectedItemId : null;
      this.close();
      DashboardState.isDirty = true;
      renderCanvas();
      triggerDebouncedAutoSave();
      if (justAddedId) {
        showToast('Đã thêm thành phần vào bên dưới Canvas!', 'success');
        setTimeout(() => {
          const el = document.getElementById(`item_${justAddedId}`);
          if (el) {
            el.scrollIntoView({ behavior: 'smooth', block: 'center' });
          }
        }, 60);
      } else {
        showToast('Đã lưu cấu hình thành công!', 'success');
      }
    }
  };

  // ===========================================================================
  // 8. SAVED DASHBOARDS CATALOG (NO DATABASE - JSON LOCAL PERSISTENCE)
  // ===========================================================================
  const SavedDashboardsCatalog = {
    STORAGE_KEY: 'mb_dashboard_configs_catalog',

    getAll() {
      try {
        if (window.JsonStorage) {
          const res = window.JsonStorage.get(this.STORAGE_KEY, []);
          if (Array.isArray(res) && res.length > 0) return res;
        }
        const local = localStorage.getItem(this.STORAGE_KEY);
        return local ? JSON.parse(local) : [];
      } catch (e) {
        return [];
      }
    },

    saveCurrent(notify = true) {
      const list = this.getAll();
      prepareDashboardConfigJson();
      const current = JSON.parse(JSON.stringify(DashboardState.present));
      current.updatedAt = new Date().toISOString();

      const existingIdx = list.findIndex(d => d.id === current.id || d.title === current.title);
      if (existingIdx >= 0) {
        list[existingIdx] = current;
      } else {
        list.unshift(current);
      }

      if (list.length > 20) list.pop();

      try {
        localStorage.setItem(this.STORAGE_KEY, JSON.stringify(list));
        if (window.JsonStorage) window.JsonStorage.set(this.STORAGE_KEY, list);
        if (notify) showToast(`Đã lưu bản sao Dashboard "${current.title}" vào kho cục bộ!`, 'success');
      } catch (e) {
        console.warn('Storage save error:', e);
      }
    },

    loadById(id) {
      const list = this.getAll();
      const item = list.find(d => d.id === id);
      if (!item) return;

      loadDashboardFromConfig(item, {
        successMsg: `✓ Đã nạp Dashboard "${item.title}" từ danh sách đã lưu!`
      });
      this.closeModal();
    },

    deleteById(id) {
      let list = this.getAll();
      list = list.filter(d => d.id !== id);
      try {
        localStorage.setItem(this.STORAGE_KEY, JSON.stringify(list));
        if (window.JsonStorage) window.JsonStorage.set(this.STORAGE_KEY, list);
      } catch (e) {}
      this.renderList();
      showToast('Đã xóa bản lưu.', 'info');
    },

    openModal() {
      this.renderList();
      const modal = document.getElementById('savedDashboardsModal');
      if (modal) modal.classList.add('open');
    },

    closeModal() {
      const modal = document.getElementById('savedDashboardsModal');
      if (modal) modal.classList.remove('open');
    },

    renderList() {
      const container = document.getElementById('savedDashboardsList');
      if (!container) return;
      const list = this.getAll();

      if (list.length === 0) {
        container.innerHTML = '<div style="text-align:center;padding:20px;color:var(--db-text-sub);">Chưa có bản Dashboard nào được lưu trên máy. Bạn có thể nhấn nút "Lưu Bản Hiện Tại" hoặc tải file JSON về máy.</div>';
        return;
      }

      container.innerHTML = '';
      list.forEach(item => {
        const card = document.createElement('div');
        card.className = 'saved-db-card';
        const dateStr = item.updatedAt ? new Date(item.updatedAt).toLocaleString('vi-VN') : '';
        const itemsCount = (item.items || []).length;
        const tabsCount = (item.tabs || []).length;

        card.innerHTML = `
          <div style="flex:1;min-width:0;">
            <div class="saved-db-title">${escapeHtml(item.title || 'Untitled Dashboard')}</div>
            <div class="saved-db-meta">
              <span>📅 ${dateStr}</span>
              <span>📊 ${itemsCount} widget / biểu đồ</span>
              <span>📑 ${tabsCount} tab</span>
            </div>
          </div>
          <div class="saved-db-actions">
            <button class="db-btn-tool primary" data-load-id="${item.id}" style="padding:4px 10px;font-size:12px;">📂 Mở</button>
            <button class="db-btn-tool" data-export-id="${item.id}" style="padding:4px 8px;font-size:12px;" title="Tải file JSON về máy">💾 JSON</button>
            <button class="db-btn-tool" data-del-id="${item.id}" style="padding:4px 8px;font-size:12px;color:#ef4444;" title="Xóa">🗑️</button>
          </div>
        `;
        container.appendChild(card);
      });
    }
  };

  // ===========================================================================
  // 9. STANDALONE HTML VIEWER GENERATOR (EXACT DASHBOARD WITH EXCEL LOADER)
  // ===========================================================================
  function generateStandaloneDashboardHTML() {
    prepareDashboardConfigJson();
    const config = DashboardState.present;
    const configJson = JSON.stringify(config).replace(/<\/script/gi, '<\\/script');
    const allFilesObj = DataStore.getAllFilesObjectCompact();
    const embeddedDataJson = JSON.stringify(allFilesObj).replace(/<\/script/gi, '<\\/script');
    const title = escapeHtml(config.title || 'Báo Cáo Dashboard');
    const pos = config.tabBarConfig?.position || 'top';
    const vWidth = config.tabBarConfig?.verticalWidth || 230;
    const canvasBg = config.theme?.canvasBg || 'dots';
    const accentColor = config.tabBarConfig?.accentColor || '#0284c7';

    return `<!DOCTYPE html>
<html lang="vi" data-theme="light">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${title} - Superset Dashboard BI Studio</title>

  <!-- Google Fonts -->
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Montserrat:wght@400;500;600;700;800&family=Playfair+Display:ital,wght@0,600;0,700;1,400&family=Roboto:ital,wght@0,400;0,500;0,700;1,400&display=swap" rel="stylesheet">

  <!-- Fallback to local styles if opened inside project directory -->
  <link rel="stylesheet" href="../../css/theme.css">
  <link rel="stylesheet" href="../../css/dashboard-builder.css">

  <!-- Apache ECharts Library (Offline Local + CDN Fallback) -->
  <script src="js/echarts.min.js"><\/script>
  <script>
    if (typeof echarts === 'undefined') {
      document.write('<script src="../../js/echarts.min.js"><\\/script>');
    }
  <\/script>
  <script>
    if (typeof echarts === 'undefined') {
      document.write('<script src="https://cdn.jsdelivr.net/npm/echarts@5.5.0/dist/echarts.min.js"><\\/script>');
    }
  <\/script>

  <!-- SheetJS Library for Excel Processing (Offline Local + CDN Fallback) -->
  <script src="js/xlsx.full.min.js"><\/script>
  <script>
    if (typeof XLSX === 'undefined') {
      document.write('<script src="../../js/xlsx.full.min.js"><\\/script>');
    }
  <\/script>
  <script>
    if (typeof XLSX === 'undefined') {
      document.write('<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"><\\/script>');
    }
  <\/script>

  <!-- Standalone Viewer Stylesheet (100% Replicating Dashboard Edit Screen) -->
  <style>
    :root, [data-theme="light"] {
      --db-header-height: 56px;
      --db-toolbar-height: 48px;
      --db-accent: ${accentColor};
      --db-accent-hover: #0369a1;
      --db-accent-orange: #ea580c;
      --db-card-bg: #ffffff;
      --db-card-border: #e2e8f0;
      --db-card-radius: 12px;
      --db-canvas-bg: #f8fafc;
      --db-text-main: #0f172a;
      --db-text-sub: #64748b;
      --bg-primary: #f8fafc;
      --bg-secondary: #ffffff;
      --bg-card: #ffffff;
      --font-family: 'Inter', system-ui, -apple-system, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
    }

    * { box-sizing: border-box; }
    html, body {
      margin: 0; padding: 0;
      font-family: var(--font-family);
      background-color: var(--db-canvas-bg);
      color: var(--db-text-main);
      min-height: 100vh;
      overflow-x: hidden !important;
      user-select: none;
    }

    /* 1. TOP NAVBAR */
    .db-header {
      position: sticky; top: 0; z-index: 100;
      height: var(--db-header-height);
      background: #ffffff;
      border-bottom: 1px solid #e2e8f0;
      display: flex; align-items: center; justify-content: space-between;
      padding: 0 24px;
      box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04);
    }
    .db-header-left {
      display: flex; align-items: center; gap: 12px; min-width: 0;
    }
    .db-title-container {
      display: flex; align-items: center; gap: 10px; min-width: 0;
    }
    .db-title-heading {
      font-size: 17px; font-weight: 700; color: #0f172a; margin: 0;
      white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
    }
    .db-header-right {
      display: flex; align-items: center; gap: 10px; flex-shrink: 0;
    }
    .db-btn-tool {
      display: inline-flex; align-items: center; gap: 6px;
      padding: 6px 14px; font-size: 13px; font-weight: 600;
      color: #0f172a; background: #f8fafc;
      border: 1px solid #cbd5e1; border-radius: 8px;
      cursor: pointer; text-decoration: none; transition: all 0.2s ease;
      white-space: nowrap; outline: none;
    }
    .db-btn-tool:hover {
      background: #f1f5f9; border-color: #94a3b8;
    }

    /* 2. WORKSPACE CONTAINER (TABS BAR + CANVAS) */
    .db-studio-container {
      display: flex; flex: 1;
      min-height: calc(100vh - var(--db-header-height));
      position: relative; width: 100%;
      box-sizing: border-box; overflow: hidden;
    }
    .db-studio-container.tabs-pos-top { flex-direction: column; }
    .db-studio-container.tabs-pos-bottom { flex-direction: column-reverse; }
    .db-studio-container.tabs-pos-left { flex-direction: row; }
    .db-studio-container.tabs-pos-right { flex-direction: row-reverse; }

    /* Tabs Bar */
    .db-tabs-bar {
      background: #ffffff;
      border-bottom: 1px solid #e2e8f0;
      display: flex; align-items: center; justify-content: flex-start;
      padding: 0 20px; min-height: 52px; z-index: 50;
      box-sizing: border-box;
    }
    .tabs-pos-bottom .db-tabs-bar {
      border-top: 1px solid #e2e8f0; border-bottom: none;
    }
    .tabs-pos-left .db-tabs-bar {
      border-right: 1px solid #e2e8f0; border-bottom: none;
      flex-direction: column; align-items: stretch; justify-content: flex-start;
      padding: 16px 12px; height: 100%; min-width: ${vWidth}px; max-width: ${vWidth}px;
    }
    .tabs-pos-right .db-tabs-bar {
      border-left: 1px solid #e2e8f0; border-bottom: none;
      flex-direction: column; align-items: stretch; justify-content: flex-start;
      padding: 16px 12px; height: 100%; min-width: ${vWidth}px; max-width: ${vWidth}px;
    }

    .db-tabs-list {
      display: flex; align-items: center; gap: 8px;
      overflow-x: auto; overflow-y: hidden; flex: 1;
      scrollbar-width: none; -ms-overflow-style: none;
    }
    .db-tabs-list::-webkit-scrollbar { display: none; }
    .tabs-pos-left .db-tabs-list,
    .tabs-pos-right .db-tabs-list {
      flex-direction: column; align-items: stretch;
      overflow-y: auto; overflow-x: hidden; gap: 6px;
    }

    /* Tab Pills */
    .db-tab-pill {
      display: inline-flex; align-items: center; gap: 8px;
      padding: 7px 16px; border-radius: 8px; font-size: 13px; font-weight: 600;
      color: #64748b; background: transparent; border: 1px solid transparent;
      cursor: pointer; transition: all 0.15s ease; white-space: nowrap; line-height: 1.4;
      outline: none;
    }
    .db-tab-pill:hover {
      background: rgba(0, 0, 0, 0.04); color: #0f172a;
    }
    .db-tab-pill.active {
      background: var(--db-accent, #0284c7);
      color: #ffffff; font-weight: 700;
      box-shadow: 0 2px 6px rgba(2, 132, 199, 0.25);
      border-color: var(--db-accent, #0284c7);
    }
    .tab-index-badge {
      font-size: 11px; font-weight: 700;
      background: rgba(0, 0, 0, 0.08); border-radius: 4px;
      padding: 1px 6px; line-height: 1.2;
    }
    .db-tab-pill.active .tab-index-badge {
      background: rgba(255, 255, 255, 0.25); color: #ffffff;
    }
    .tab-name-label {
      display: inline-block;
    }

    /* Tab Styles */
    .tab-style-cards .db-tab-pill {
      border: 1px solid #e2e8f0; background: #f8fafc; border-radius: 8px 8px 0 0;
    }
    .tab-style-cards .db-tab-pill.active {
      background: #ffffff; color: var(--db-accent, #0284c7);
      border-color: var(--db-accent, #0284c7); border-bottom-color: #ffffff;
    }
    .tab-style-underline .db-tab-pill {
      border-radius: 0; border-bottom: 2px solid transparent; background: transparent;
    }
    .tab-style-underline .db-tab-pill.active {
      background: transparent; color: var(--db-accent, #0284c7);
      border-bottom-color: var(--db-accent, #0284c7); box-shadow: none;
    }
    .tab-size-compact .db-tab-pill { padding: 4px 10px; font-size: 12px; }
    .tab-size-comfortable .db-tab-pill { padding: 10px 22px; font-size: 14px; }

    /* 3. CANVAS WRAPPER & CANVAS STUDIO */
    .db-canvas-wrapper {
      flex: 1; overflow-x: hidden !important; overflow-y: auto !important;
      background-color: #f1f5f9;
      display: flex; justify-content: center;
      padding: 24px 16px 80px 16px; box-sizing: border-box;
      width: 100%;
    }
    .canvas-bg-white {
      background-color: #ffffff !important; background-image: none !important;
    }
    .canvas-bg-dots {
      background-color: #f8fafc !important;
      background-image: radial-gradient(#94a3b8 1.4px, transparent 1.4px) !important;
      background-size: 20px 20px !important;
    }
    .canvas-bg-stripes {
      background-color: #ffffff !important;
      background-image: linear-gradient(90deg, #e2e8f0 1.5px, transparent 1.5px) !important;
      background-size: 24px 24px !important;
    }
    .canvas-bg-horizontal {
      background-color: #ffffff !important;
      background-image: linear-gradient(#e2e8f0 1.5px, transparent 1.5px) !important;
      background-size: 24px 24px !important;
    }
    .canvas-bg-grid {
      background-color: #ffffff !important;
      background-image: linear-gradient(#e2e8f0 1.2px, transparent 1.2px), linear-gradient(90deg, #e2e8f0 1.2px, transparent 1.2px) !important;
      background-size: 22px 22px !important;
    }

    .db-canvas-viewport {
      padding: 0; margin: 0 auto; width: 100%;
      display: flex; justify-content: center; align-items: flex-start;
      overflow: visible;
      min-height: 100%; height: auto; flex: 1;
    }

    .db-canvas {
      width: 1360px !important; min-width: 1360px !important; max-width: 1360px !important;
      margin: 0 auto !important; position: relative !important;
      overflow: visible !important; box-sizing: border-box !important;
      background-color: #ffffff !important;
      border: 2px solid #2563eb !important;
      border-radius: 10px !important;
      box-shadow: 0 12px 36px rgba(15, 23, 42, 0.12), 0 0 0 1px rgba(37, 99, 235, 0.2) !important;
      padding: 24px !important; min-height: 800px;
      transform-origin: top center; transition: transform 0.1s ease-out;
    }

    /* 2 Boundary Guides */
    .db-canvas-boundary-guide {
      position: absolute; top: 0; bottom: 0; width: 0;
      pointer-events: none; z-index: 5;
      border-left: 1px dashed rgba(37, 99, 235, 0.35);
    }
    .db-canvas-boundary-guide.is-left { left: 24px; }
    .db-canvas-boundary-guide.is-right { left: 1336px; }

    /* 4. ITEMS & CARDS (MATCHING EDIT SCREEN PREVIEW MODE) */
    .db-item-box {
      position: absolute; background: #ffffff;
      border: 1px solid #e2e8f0; border-radius: 12px;
      box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08);
      display: flex; flex-direction: column; overflow: hidden;
      box-sizing: border-box; cursor: default;
      transition: box-shadow 0.2s ease;
    }
    .db-item-box.is-type-filter,
    .db-top-filter-bar,
    .db-top-filter-body,
    .db-top-filter-item-wrapper {
      overflow: visible !important;
    }
    .db-item-box.is-type-filter .db-item-body,
    .db-item-box.is-type-filter .db-filter-container,
    .db-item-box.is-type-filter .db-filter-controls-row,
    .db-item-box.is-type-filter .db-filter-multiselect-container,
    .db-top-filter-item-wrapper .db-item-body,
    .db-top-filter-item-wrapper .db-filter-container,
    .db-top-filter-item-wrapper .db-filter-controls-row,
    .db-top-filter-item-wrapper .db-filter-multiselect-container {
      overflow: visible !important;
    }
    .db-top-filter-item-wrapper .db-filter-label {
      display: none !important;
    }
    .db-item-box.is-type-filter.dropdown-open,
    .db-item-box.is-type-filter:has(.db-filter-multiselect-dropdown.open),
    .db-top-filter-item-wrapper.dropdown-open,
    .db-top-filter-item-wrapper:has(.db-filter-multiselect-dropdown.open),
    .db-top-filter-body:has(.db-filter-multiselect-dropdown.open),
    .db-top-filter-bar.dropdown-open,
    .db-top-filter-bar:has(.db-filter-multiselect-dropdown.open) {
      z-index: 1000000 !important;
    }
    .db-item-box:hover {
      box-shadow: 0 6px 24px rgba(0, 0, 0, 0.12);
    }
    .db-item-box.is-frame {
      background: rgba(255, 255, 255, 0.4) !important;
      border: 2px dashed #cbd5e1 !important;
      box-shadow: none !important;
      z-index: 1;
    }
    .db-item-box.is-frameless-element,
    .db-item-box.is-shape-line {
      background: transparent !important;
      border: none !important;
      box-shadow: none !important;
      border-radius: 0 !important;
    }
    .db-item-box.is-frameless-element .db-item-header,
    .db-item-box.is-shape-line .db-item-header,
    .db-item-box.hide-header .db-item-header {
      display: none !important;
    }
    .db-item-box.is-frameless-element .db-item-body,
    .db-item-box.is-shape-line .db-item-body,
    .db-item-box.hide-header .db-item-body {
      height: 100% !important; padding: 0 !important;
    }

    .db-item-header {
      height: 38px; padding: 0 12px;
      display: flex; align-items: center; justify-content: flex-start;
      font-size: 13px; font-weight: 700;
      border-bottom: 1px solid #f1f5f9; background: rgba(255, 255, 255, 0.02);
      color: #0f172a; flex-shrink: 0; user-select: none;
    }
    .db-item-header-left {
      display: flex; align-items: center; gap: 6px; min-width: 0;
    }
    .db-item-title {
      font-size: 13px; font-weight: 700; color: #0f172a;
      overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
    }
    .db-item-body {
      flex: 1; padding: 10px; position: relative; min-height: 0;
      display: flex; flex-direction: column; box-sizing: border-box;
    }

    .db-chart-canvas {
      width: 100%; height: 100%; min-height: 80px;
    }
    .db-chart-error-banner {
      color: #ef4444; font-size: 12px; padding: 12px; text-align: center;
    }

    .db-shape-container {
      width: 100%; height: 100%; box-sizing: border-box;
    }
    .db-frame-container {
      width: 100%; height: 100%; box-sizing: border-box;
    }
    .db-text-container {
      box-sizing: border-box;
    }

    .db-action-btn-widget {
      width: 100%; height: 100%; display: flex; align-items: center; justify-content: center;
      cursor: pointer; user-select: none; transition: transform 0.1s, opacity 0.15s;
      box-sizing: border-box;
    }
    .db-action-btn-widget:hover { opacity: 0.92; transform: translateY(-1px); }
    .db-action-btn-widget:active { transform: translateY(0); }

    .db-image-widget {
      width: 100%; height: 100%; display: flex; align-items: center; justify-content: center;
      overflow: hidden; box-sizing: border-box;
    }
    .db-image-widget img {
      max-width: 100%; max-height: 100%; display: block;
    }

    /* 5. BI PLATFORM FILTERS (ALL TYPES) */
    .db-filter-container {
      display: flex; flex-direction: column; gap: 6px; justify-content: center;
      height: 100%; padding: 6px 10px; box-sizing: border-box;
    }
    .db-filter-header-row {
      display: flex; align-items: center; justify-content: space-between; gap: 6px;
    }
    .db-filter-label {
      font-size: 11.5px; font-weight: 700; color: #334155;
      white-space: nowrap; overflow: hidden; text-overflow: ellipsis; flex: 1;
    }
    .db-filter-reset-btn {
      background: transparent; border: none; color: #64748b; font-size: 11px;
      cursor: pointer; padding: 1px 5px; border-radius: 4px; line-height: 1;
      transition: all 0.15s;
    }
    .db-filter-reset-btn:hover { background: #fee2e2; color: #ef4444; }
    .db-filter-active-badge {
      font-size: 10px; background: #e0f2fe; color: #0369a1; padding: 1px 5px;
      border-radius: 4px; font-weight: 600; white-space: nowrap;
    }
    .db-filter-parent-badge {
      font-size: 10px; background: #f0fdf4; color: #15803d; border: 1px solid #bbf7d0;
      padding: 1px 5px; border-radius: 4px; font-weight: 600; white-space: nowrap;
    }
    .db-filter-controls-row {
      display: flex; align-items: center; gap: 6px;
    }
    .db-filter-op-select {
      width: 82px; background: #f8fafc; border: 1px solid #cbd5e1;
      border-radius: 6px; padding: 5px 6px; font-size: 12px; font-weight: 700;
      color: #0284c7; outline: none; cursor: pointer; flex-shrink: 0;
    }
    .db-filter-val-input,
    .db-filter-val-select {
      flex: 1; min-width: 0; background: #ffffff; border: 1px solid #cbd5e1;
      border-radius: 6px; padding: 5px 8px; font-size: 12px; color: #1e293b;
      outline: none; transition: border-color 0.15s;
    }
    .db-filter-val-input:focus,
    .db-filter-val-select:focus { border-color: #0284c7; }
    .db-filter-range-wrap,
    .db-filter-date-range-wrap {
      display: flex; align-items: center; gap: 4px; flex: 1; min-width: 0;
    }
    .db-filter-range-input,
    .db-filter-date-input {
      flex: 1; min-width: 0; background: #ffffff; border: 1px solid #cbd5e1;
      border-radius: 6px; padding: 4px 6px; font-size: 11px; color: #1e293b;
      outline: none; font-family: inherit; transition: border-color 0.15s, box-shadow 0.15s;
    }
    .db-filter-range-input:focus,
    .db-filter-date-input:focus {
      border-color: #0284c7; box-shadow: 0 0 0 2px rgba(2, 132, 199, 0.15);
    }
    .db-filter-range-sep,
    .db-filter-date-sep {
      font-size: 11px; color: #64748b; font-weight: 700; flex-shrink: 0;
    }
    .db-filter-multiselect-container {
      position: relative; flex: 1; min-width: 0;
    }
    .db-filter-multiselect-btn {
      width: 100%; display: flex; align-items: center; justify-content: space-between;
      background: #ffffff; border: 1px solid #cbd5e1; border-radius: 6px;
      padding: 4px 8px; font-size: 11px; color: #1e293b; cursor: pointer;
      outline: none; text-align: left; white-space: nowrap; overflow: hidden;
      text-overflow: ellipsis;
    }
    .db-filter-multiselect-btn:hover { border-color: #94a3b8; }
    .db-filter-multiselect-dropdown {
      position: absolute; top: calc(100% + 4px); left: 0;
      min-width: 230px; width: 100%; max-width: 360px;
      background: #ffffff; border: 1px solid #cbd5e1;
      border-radius: 8px; box-shadow: 0 12px 36px rgba(0, 0, 0, 0.25), 0 4px 12px rgba(0, 0, 0, 0.1);
      z-index: 999999 !important; padding: 8px; display: none; box-sizing: border-box;
    }
    .db-filter-multiselect-dropdown.open {
      display: flex !important; flex-direction: column !important; gap: 6px !important;
      height: 250px !important; box-sizing: border-box !important;
    }
    .db-filter-multiselect-search {
      width: 100%; box-sizing: border-box; border: 1px solid #cbd5e1; border-radius: 6px;
      padding: 5px 8px; font-size: 11.5px; outline: none; flex-shrink: 0;
    }
    .db-filter-multiselect-actions {
      display: flex; justify-content: space-between; gap: 6px; font-size: 10.5px; flex-shrink: 0; padding: 2px 0;
    }
    .db-filter-ms-action-btn {
      flex: 1; background: #f1f5f9; border: 1px solid #e2e8f0; border-radius: 4px;
      padding: 3px 6px; font-size: 10.5px; cursor: pointer; color: #334155; font-weight: 600; text-align: center;
      transition: all 0.15s;
    }
    .db-filter-ms-action-btn:hover { background: #e2e8f0; color: #0f172a; }
    .db-filter-multiselect-list {
      flex: 1 1 0% !important; min-height: 0 !important; max-height: none !important;
      overflow-y: auto !important; overflow-x: hidden !important; display: flex !important; flex-direction: column !important;
      gap: 2px !important; padding-right: 4px;
    }
    .db-filter-multiselect-item {
      display: flex; align-items: center; gap: 6px; padding: 4px 6px;
      border-radius: 4px; font-size: 11.5px; color: #1e293b; cursor: pointer; user-select: none; flex-shrink: 0;
    }
    /* Sticky Top Filter Bar (Thanh Bộ Lọc Cố Định Trên Cùng) */
    .db-top-filter-bar {
      width: 100%;
      margin: 0 0 16px 0;
      background: #ffffff;
      border: 1px solid #e2e8f0;
      border-radius: 12px;
      box-shadow: 0 2px 10px rgba(0, 0, 0, 0.05);
      transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
      box-sizing: border-box;
      overflow: visible !important;
    }
    .db-top-filter-bar.is-sticky {
      position: relative;
      z-index: 950;
      background: rgba(255, 255, 255, 0.96);
      backdrop-filter: blur(12px);
      -webkit-backdrop-filter: blur(12px);
      box-shadow: 0 4px 16px rgba(15, 23, 42, 0.08);
      border-color: #cbd5e1;
    }
    .db-top-filter-bar.is-pinned-floating {
      margin-top: 0 !important;
      z-index: 999990 !important;
      box-shadow: 0 12px 32px rgba(15, 23, 42, 0.18), 0 2px 6px rgba(0, 0, 0, 0.08) !important;
      border-color: #3b82f6 !important;
      background: rgba(255, 255, 255, 0.98) !important;
    }
    .db-top-filter-bar.is-collapsed.collapse-style-circle {
      width: 48px !important;
      height: 48px !important;
      min-width: 48px !important;
      max-width: 48px !important;
      padding: 0 !important;
      border-radius: 50% !important;
      background: transparent !important;
      border: none !important;
      box-shadow: none !important;
      overflow: visible !important;
    }
    .db-top-filter-circle-btn {
      display: none;
      width: 48px;
      height: 48px;
      border-radius: 50%;
      background: #ffffff;
      border: 2px solid #3b82f6;
      box-shadow: 0 4px 16px rgba(15, 23, 42, 0.16), 0 2px 6px rgba(59, 130, 246, 0.2);
      align-items: center;
      justify-content: center;
      position: relative;
      cursor: pointer;
      transition: transform 0.2s cubic-bezier(0.34, 1.56, 0.64, 1), box-shadow 0.2s ease, border-color 0.2s ease;
      user-select: none;
      box-sizing: border-box;
    }
    .db-top-filter-bar.is-collapsed.collapse-style-circle .db-top-filter-circle-btn {
      display: flex !important;
    }
    .db-top-filter-bar.is-collapsed.collapse-style-circle .db-top-filter-expanded-wrap {
      display: none !important;
    }
    .db-top-filter-circle-btn:hover {
      transform: scale(1.1);
      box-shadow: 0 6px 22px rgba(59, 130, 246, 0.35);
      border-color: #2563eb;
    }
    .db-top-filter-circle-btn:active {
      transform: scale(0.95);
    }
    .db-filter-circle-icon {
      font-size: 20px;
      line-height: 1;
    }
    .db-filter-circle-badge {
      position: absolute;
      top: -4px;
      right: -4px;
      background: #ef4444;
      color: #ffffff;
      font-size: 11px;
      font-weight: 700;
      min-width: 18px;
      height: 18px;
      padding: 0 4px;
      border-radius: 10px;
      display: flex;
      align-items: center;
      justify-content: center;
      box-shadow: 0 2px 5px rgba(0, 0, 0, 0.25);
      border: 1.5px solid #ffffff;
    }
    .db-top-filter-bar:not(.is-collapsed) .db-top-filter-circle-btn,
    .db-top-filter-bar:not(.collapse-style-circle) .db-top-filter-circle-btn {
      display: none !important;
    }
    .db-top-filter-bar:has(.db-filter-multiselect-dropdown.open),
    .db-top-filter-bar.dropdown-open {
      z-index: 1000000 !important;
    }
    .db-top-filter-item-wrapper.dropdown-open,
    .db-top-filter-item-wrapper:has(.db-filter-multiselect-dropdown.open) {
      z-index: 999999 !important;
    }
    .db-top-filter-bar.is-movable {
      position: absolute;
      box-shadow: 0 8px 24px rgba(15, 23, 42, 0.08);
      z-index: 900;
    }
    .db-top-filter-header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 10px 16px;
      background: #f8fafc;
      border-bottom: 1px solid #e2e8f0;
      border-radius: 12px 12px 0 0;
      user-select: none;
    }
    .db-top-filter-bar.is-collapsed .db-top-filter-header {
      border-bottom: none;
      border-radius: 12px;
    }
    .db-top-filter-title-group {
      display: flex;
      align-items: center;
      gap: 8px;
    }
    .db-top-filter-icon {
      font-size: 15px;
    }
    .db-top-filter-title {
      font-size: 13px;
      font-weight: 700;
      color: #0f172a;
      letter-spacing: 0.3px;
    }
    .db-top-filter-count-badge {
      font-size: 11px;
      font-weight: 600;
      background: #e0f2fe;
      color: #0284c7;
      padding: 2px 8px;
      border-radius: 999px;
    }
    .db-top-filter-actions {
      display: flex;
      align-items: center;
      gap: 8px;
    }
    .db-top-filter-btn {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      padding: 5px 12px;
      font-size: 12px;
      font-weight: 600;
      color: #475569;
      background: #ffffff;
      border: 1px solid #cbd5e1;
      border-radius: 6px;
      cursor: pointer;
      transition: all 0.2s ease;
    }
    .db-top-filter-btn:hover {
      background: #f1f5f9;
      color: #0f172a;
      border-color: #94a3b8;
    }
    .db-top-filter-btn.primary {
      background: var(--db-accent, #0284c7);
      color: #ffffff;
      border-color: transparent;
    }
    .db-top-filter-btn.primary:hover {
      background: var(--db-accent-hover, #0369a1);
    }
    .db-top-filter-body {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
      gap: 12px;
      padding: 14px 16px;
      background: #ffffff;
      border-radius: 0 0 12px 12px;
      transition: max-height 0.3s ease, opacity 0.25s ease, padding 0.3s ease;
      overflow: visible;
    }
    .db-top-filter-bar.is-collapsed .db-top-filter-body {
      display: none !important;
    }
    .db-top-filter-item-wrapper {
      background: #f8fafc;
      border: 1px solid #e2e8f0;
      border-radius: 8px;
      padding: 10px 12px;
      display: flex;
      flex-direction: column;
      gap: 8px;
      position: relative;
    }
    .db-top-filter-item-header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 8px;
    }
    .db-top-filter-item-title {
      font-size: 12px;
      font-weight: 700;
      color: #1e293b;
    }
  </style>
</head>

<body class="dashboard-app-body preview-mode">
  <!-- 1. Header Web Dashboard (Tái hiện thanh tiêu đề) -->
  <header class="db-header">
    <div class="db-header-left">
      <div class="db-title-container">
        <span style="font-size:22px;">📊</span>
        <h1 class="db-title-heading" id="dashboardTitleHeading">${title}</h1>
      </div>
    </div>
    <div class="db-header-right">
      <button class="db-btn-tool" id="btnViewerUploadExcel" title="Cập nhật nguồn dữ liệu Excel cho Dashboard">
        📤 Nạp Excel Mới
      </button>
      <input type="file" id="viewerExcelFileInput" accept=".xlsx,.xls,.csv" style="display:none;" />
    </div>
  </header>

  <!-- 2. Studio Workspace Container (Tabs Bar + Canvas) -->
  <div class="db-studio-container tabs-pos-${pos}" id="dbStudioContainer">
    <!-- Thanh Tab Báo Cáo -->
    <section class="db-tabs-bar" id="dbTabsBar">
      <div class="db-tabs-list" id="dbTabsList"></div>
    </section>

    <!-- Canvas Studio -->
    <main class="db-canvas-wrapper canvas-bg-${canvasBg}" id="canvasWrapper">
      <div class="db-canvas-viewport" id="canvasViewport">
        <div class="db-canvas" id="dashboardCanvas"></div>
      </div>
    </main>
  </div>

  <script>
    const PARQUET_THRESHOLD_BYTES = 50 * 1024 * 1024; // 50MB
    const ParquetEngine = {
      isParquet(obj) {
        return Boolean(obj && (obj.__isParquet || (obj.__table && obj.__table.__isParquet)));
      },
      encodeToParquet(rows, metadata = {}) {
        if (!Array.isArray(rows) || rows.length === 0) return rows;
        if (this.isParquet(rows)) return rows;

        const colNames = Object.keys(rows[0] || {});
        const rowCount = rows.length;
        const columns = {};
        const columnTypes = {};

        for (let c = 0; c < colNames.length; c++) {
          const col = colNames[c];
          let isAllNumeric = true;
          let hasValue = false;
          const uniqueMap = new Map();
          let uniqueCounter = 0;
          const indices = new Array(rowCount);
          const rawNumbers = new Array(rowCount);

          for (let i = 0; i < rowCount; i++) {
            const val = rows[i][col];
            if (val !== null && val !== undefined && val !== '') {
              hasValue = true;
              const num = Number(val);
              if (isNaN(num)) {
                isAllNumeric = false;
              } else {
                rawNumbers[i] = num;
              }
            } else {
              rawNumbers[i] = 0;
            }

            const strKey = String(val !== undefined && val !== null ? val : '');
            if (!uniqueMap.has(strKey)) {
              uniqueMap.set(strKey, uniqueCounter++);
            }
            indices[i] = uniqueMap.get(strKey);
          }

          if (isAllNumeric && hasValue) {
            columnTypes[col] = 'numeric';
            columns[col] = {
              type: 'numeric',
              data: rawNumbers
            };
          } else {
            const dictArray = new Array(uniqueMap.size);
            uniqueMap.forEach((idx, key) => {
              dictArray[idx] = key;
            });
            columnTypes[col] = 'dictionary';
            columns[col] = {
              type: 'dictionary',
              dictionary: dictArray,
              indices: indices
            };
          }
        }

        const parquetTable = {
          __isParquet: true,
          rowCount,
          colNames,
          columnTypes,
          columns,
          metadata: {
            ...metadata,
            encodedAt: new Date().toISOString()
          },
          getValue(col, rowIdx) {
            const colData = this.columns[col];
            if (!colData) return undefined;
            if (colData.type === 'numeric') {
              return colData.data[rowIdx];
            }
            const dictIdx = colData.indices[rowIdx];
            return colData.dictionary[dictIdx];
          },
          hasColumn(col) {
            return Boolean(this.columns[col]);
          },
          extractRowObject(rowIdx) {
            const rowObj = {};
            for (let i = 0; i < this.colNames.length; i++) {
              const col = this.colNames[i];
              rowObj[col] = this.getValue(col, rowIdx);
            }
            return rowObj;
          }
        };

        return parquetTable;
      },

      createVirtualArray(table) {
        if (!table) return [];
        if (!table.__isParquet && Array.isArray(table)) return table;

        const tbl = table;
        const len = tbl.rowCount;

        class VirtualArray extends Array {
          constructor() {
            super(len);
            Object.defineProperty(this, '__table', { value: tbl, enumerable: false, writable: true });
            Object.defineProperty(this, '__isParquet', { value: true, enumerable: false });
            Object.defineProperty(this, '__columns', { value: tbl.colNames, enumerable: false });
            Object.defineProperty(this, 'length', { value: len, enumerable: false, writable: false });
          }

          map(callback, thisArg) {
            const res = new Array(len);
            let curIdx = 0;
            const flyweight = new Proxy({}, {
              get(_, prop) {
                if (prop === '__rawIndex') return curIdx;
                return tbl.getValue(prop, curIdx);
              },
              has(_, prop) {
                return tbl.hasColumn(prop);
              },
              ownKeys() {
                return tbl.colNames;
              },
              getOwnPropertyDescriptor() {
                return { enumerable: true, configurable: true };
              }
            });

            for (let i = 0; i < len; i++) {
              curIdx = i;
              res[i] = callback.call(thisArg, flyweight, i, this);
            }
            return res;
          }

          forEach(callback, thisArg) {
            let curIdx = 0;
            const flyweight = new Proxy({}, {
              get(_, prop) {
                if (prop === '__rawIndex') return curIdx;
                return tbl.getValue(prop, curIdx);
              },
              has(_, prop) {
                return tbl.hasColumn(prop);
              },
              ownKeys() {
                return tbl.colNames;
              },
              getOwnPropertyDescriptor() {
                return { enumerable: true, configurable: true };
              }
            });

            for (let i = 0; i < len; i++) {
              curIdx = i;
              callback.call(thisArg, flyweight, i, this);
            }
          }

          filter(callback, thisArg) {
            const res = [];
            let curIdx = 0;
            const flyweight = new Proxy({}, {
              get(_, prop) {
                if (prop === '__rawIndex') return curIdx;
                return tbl.getValue(prop, curIdx);
              },
              has(_, prop) {
                return tbl.hasColumn(prop);
              },
              ownKeys() {
                return tbl.colNames;
              },
              getOwnPropertyDescriptor() {
                return { enumerable: true, configurable: true };
              }
            });

            for (let i = 0; i < len; i++) {
              curIdx = i;
              if (callback.call(thisArg, flyweight, i, this)) {
                res.push(tbl.extractRowObject(i));
              }
            }
            return res;
          }

          slice(start = 0, end = len) {
            const actualStart = Math.max(0, start < 0 ? len + start : start);
            const actualEnd = Math.min(len, end < 0 ? len + end : end);
            const res = [];
            for (let i = actualStart; i < actualEnd; i++) {
              res.push(tbl.extractRowObject(i));
            }
            return res;
          }

          reduce(callback, initialValue) {
            let accumulator = initialValue;
            let startIndex = 0;
            if (accumulator === undefined) {
              if (len === 0) throw new TypeError('Reduce of empty array with no initial value');
              accumulator = tbl.extractRowObject(0);
              startIndex = 1;
            }
            let curIdx = 0;
            const flyweight = new Proxy({}, {
              get(_, prop) {
                return tbl.getValue(prop, curIdx);
              }
            });
            for (let i = startIndex; i < len; i++) {
              curIdx = i;
              accumulator = callback(accumulator, flyweight, i, this);
            }
            return accumulator;
          }

          toRawArray() {
            const res = new Array(len);
            for (let i = 0; i < len; i++) {
              res[i] = tbl.extractRowObject(i);
            }
            return res;
          }
        }

        const instance = new VirtualArray();
        return new Proxy(instance, {
          get(target, prop) {
            if (typeof prop === 'string') {
              const idx = Number(prop);
              if (!isNaN(idx) && Number.isInteger(idx) && idx >= 0 && idx < len) {
                return tbl.extractRowObject(idx);
              }
            }
            return target[prop];
          }
        });
      }
    };

    const CONFIG = ${configJson};
    let FILES = ${embeddedDataJson};

    // Tự động giải nén cấu trúc Compact Matrix Data hoặc kích hoạt Parquet Virtual Array
    function uncompactSheet(sheetData) {
      if (!sheetData) return [];
      if (sheetData.__isParquet) {
        return ParquetEngine.createVirtualArray(sheetData);
      }
      if (Array.isArray(sheetData)) return sheetData;
      if (sheetData.__c && Array.isArray(sheetData.cols) && Array.isArray(sheetData.data)) {
        const cols = sheetData.cols;
        const len = cols.length;
        return sheetData.data.map(arr => {
          const row = {};
          for (let i = 0; i < len; i++) row[cols[i]] = arr[i];
          return row;
        });
      }
      return sheetData;
    }

    for (const fn in FILES) {
      if (FILES[fn] && typeof FILES[fn] === 'object') {
        for (const sn in FILES[fn]) {
          FILES[fn][sn] = uncompactSheet(FILES[fn][sn]);
        }
      }
    }

    let activeTabId = CONFIG.activeTabId || (CONFIG.tabs && CONFIG.tabs[0] ? CONFIG.tabs[0].id : 'default');
    const tabFilters = JSON.parse(JSON.stringify(CONFIG.tabFilters || {}));
    let activeFilters = { ...((tabFilters && tabFilters[activeTabId]) || CONFIG.activeFilters || {}) };
    let openDropdownFilterId = null;

    function switchDashboardTab(targetTabId) {
      if (!targetTabId || targetTabId === activeTabId) return;
      tabFilters[activeTabId] = { ...activeFilters };
      activeTabId = targetTabId;
      activeFilters = { ...(tabFilters[targetTabId] || {}) };
      openDropdownFilterId = null;
      renderAllTabs();
      renderCanvas();
    }

    function escapeHtml(str) {
      if (str === null || str === undefined) return '';
      return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
    }

    function parseDateToTimestamp(val) {
      if (val === null || val === undefined || val === '') return null;
      if (val instanceof Date) return isNaN(val.getTime()) ? null : new Date(val.getFullYear(), val.getMonth(), val.getDate()).getTime();
      if (typeof val === 'number') {
        if (val > 1000 && val < 100000) {
          const utcDays = Math.floor(val - 25569);
          const d = new Date(utcDays * 86400 * 1000);
          return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
        }
        if (val > 1000000000000) {
          const d = new Date(val);
          return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
        }
      }
      const s = String(val).trim();
      if (!s) return null;

      // YYYY-MM-DD or YYYY/MM/DD or YYYY.MM.DD (ưu tiên chuẩn ISO trước)
      const ymd = s.match(/^(\\d{4})[\\/\\-\\.](\\d{1,2})[\\/\\-\\.](\\d{1,2})/);
      if (ymd) {
        const year = parseInt(ymd[1], 10);
        const month = parseInt(ymd[2], 10) - 1;
        const day = parseInt(ymd[3], 10);
        const d = new Date(year, month, day);
        if (!isNaN(d.getTime())) return d.getTime();
      }

      // DD/MM/YYYY or DD-MM-YYYY or DD.MM.YYYY
      const dmy = s.match(/^(\\d{1,2})[\\/\\-\\.](\\d{1,2})[\\/\\-\\.](\\d{4})/);
      if (dmy) {
        const day = parseInt(dmy[1], 10);
        const month = parseInt(dmy[2], 10) - 1;
        const year = parseInt(dmy[3], 10);
        const d = new Date(year, month, day);
        if (!isNaN(d.getTime())) return d.getTime();
      }

      if (/^\\d{1,7}$/.test(s) || /^\\d{9,}$/.test(s)) return null;

      const p = Date.parse(s);
      if (!isNaN(p)) {
        const d = new Date(p);
        return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
      }
      return null;
    }

    function formatDateToYMD(val) {
      if (!val) return '';
      const s = String(val).trim();
      if (/^\\d{4}-\\d{2}-\\d{2}$/.test(s)) return s;
      const ts = parseDateToTimestamp(val);
      if (!ts) return '';
      const d = new Date(ts);
      return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
    }

    function isProbableDateColumn(colName, rows = []) {
      if (!colName) return false;
      const lower = String(colName).toLowerCase();
      if (lower.includes('ngày') || lower.includes('ngay') || lower.includes('date') || lower.includes('time') || lower.includes('thời gian') || lower.includes('thoi_gian')) {
        return true;
      }
      if (!rows || rows.length === 0) return false;
      let validDateCount = 0;
      let checkedCount = 0;
      for (let i = 0; i < Math.min(rows.length, 20); i++) {
        const v = rows[i][colName];
        if (v !== undefined && v !== null && v !== '') {
          checkedCount++;
          if (parseDateToTimestamp(v) !== null) validDateCount++;
        }
      }
      return checkedCount > 0 && (validDateCount / checkedCount) >= 0.6;
    }

    function getData(sourceKey, fallbackCol = null) {
      if (!sourceKey && fallbackCol) {
        for (const f in FILES) {
          for (const s in FILES[f]) {
            const rows = FILES[f][s];
            if (rows && rows.length > 0 && rows[0] && rows[0][fallbackCol] !== undefined) {
              return rows;
            }
          }
        }
      }
      if (!sourceKey) {
        const firstFile = Object.keys(FILES)[0];
        if (!firstFile) return [];
        const firstSheet = Object.keys(FILES[firstFile])[0];
        return FILES[firstFile][firstSheet] || [];
      }
      const parts = String(sourceKey).split('::');
      const fileName = parts[0];
      const sheetName = parts[1];
      if (fileName && sheetName && FILES[fileName] && FILES[fileName][sheetName]) {
        return FILES[fileName][sheetName];
      }
      if (FILES[fileName]) {
        const firstS = Object.keys(FILES[fileName])[0];
        if (firstS && FILES[fileName][firstS]) return FILES[fileName][firstS];
      }
      for (const f in FILES) {
        if (FILES[f][sheetName || sourceKey]) return FILES[f][sheetName || sourceKey];
      }
      const firstF = Object.keys(FILES)[0];
      if (firstF) {
        const firstS = Object.keys(FILES[firstF])[0];
        return FILES[firstF][firstS] || [];
      }
      return [];
    }

    function getFilteredData(sourceKey, excludeCols = [], targetTabId = null) {
      let rows = getData(sourceKey);
      if (!rows || rows.length === 0) return [];
      const tabId = targetTabId || activeTabId || 'default';
      const filters = tabId === activeTabId ? activeFilters : (tabFilters[tabId] || {});
      const filterKeys = Object.keys(filters).filter(k => !excludeCols.includes(k));
      if (filterKeys.length === 0) return rows;

      return rows.filter(row => {
        for (const col of filterKeys) {
          const filterVal = filters[col];
          if (!filterVal || row[col] === undefined) continue;

          let op = '=';
          let target = '';
          let startDate = '';
          let endDate = '';
          let minVal = '';
          let maxVal = '';
          let valuesList = null;
          let isDateFilter = false;

          if (typeof filterVal === 'object' && filterVal !== null) {
            op = filterVal.op || '=';
            target = filterVal.val !== undefined ? filterVal.val : '';
            startDate = filterVal.start || '';
            endDate = filterVal.end || '';
            minVal = filterVal.min !== undefined ? filterVal.min : (filterVal.start !== undefined ? filterVal.start : '');
            maxVal = filterVal.max !== undefined ? filterVal.max : (filterVal.end !== undefined ? filterVal.end : '');
            valuesList = Array.isArray(filterVal.values) ? filterVal.values : (Array.isArray(filterVal.val) ? filterVal.val : null);
            isDateFilter = Boolean(filterVal.isDate || op === 'date_range' || op.startsWith('date_'));
          } else {
            target = filterVal;
          }

          const cell = row[col];
          if (cell === undefined || cell === null) return false;

          if (op === 'between' || op === 'range' || op === 'date_range') {
            if (isDateFilter || op === 'date_range') {
              const cellTs = parseDateToTimestamp(cell);
              if (cellTs === null) return false;
              let startTs = startDate ? parseDateToTimestamp(startDate) : (minVal ? parseDateToTimestamp(minVal) : null);
              let endTs = endDate ? parseDateToTimestamp(endDate) : (maxVal ? parseDateToTimestamp(maxVal) : null);
              if (startTs !== null && endTs !== null && startTs > endTs) {
                const tmp = startTs;
                startTs = endTs;
                endTs = tmp;
              }
              if (startTs !== null && cellTs < startTs) return false;
              if (endTs !== null && cellTs > endTs) return false;
              continue;
            }

            const sMin = minVal !== undefined && minVal !== null ? String(minVal).trim() : '';
            const sMax = maxVal !== undefined && maxVal !== null ? String(maxVal).trim() : '';
            if (sMin === '' && sMax === '') continue;

            const numCell = Number(cell);
            const hasNumCell = !isNaN(numCell);
            let numMin = sMin !== '' ? Number(sMin) : null;
            let numMax = sMax !== '' ? Number(sMax) : null;

            if (numMin !== null && numMax !== null && !isNaN(numMin) && !isNaN(numMax) && numMin > numMax) {
              const tmp = numMin;
              numMin = numMax;
              numMax = tmp;
            }

            if (hasNumCell && ((numMin !== null && !isNaN(numMin)) || (numMax !== null && !isNaN(numMax)))) {
              if (numMin !== null && !isNaN(numMin) && numCell < numMin) return false;
              if (numMax !== null && !isNaN(numMax) && numCell > numMax) return false;
              continue;
            } else {
              let strMin = sMin !== '' ? sMin.toLowerCase() : null;
              let strMax = sMax !== '' ? sMax.toLowerCase() : null;
              if (strMin && strMax && strMin > strMax) {
                const tmp = strMin;
                strMin = strMax;
                strMax = tmp;
              }
              if (strMin !== null && String(cell).toLowerCase() < strMin) return false;
              if (strMax !== null && String(cell).toLowerCase() > strMax) return false;
              continue;
            }
          }

          if (op === 'in' || valuesList !== null) {
            const list = valuesList || [];
            if (!list || list.length === 0 || list.includes('all')) continue;
            const strCell = String(cell).trim().toLowerCase();
            const match = list.some(item => String(item).trim().toLowerCase() === strCell);
            if (!match) return false;
            continue;
          }

          if (target === undefined || target === null || target === '' || target === 'all') {
            continue;
          }

          if (isDateFilter) {
            const cellTs = parseDateToTimestamp(cell);
            if (cellTs === null) return false;
            const targetTs = parseDateToTimestamp(target);
            if (targetTs === null) continue;

            if (op === 'date_eq' || op === '=') {
              if (cellTs !== targetTs) return false;
            } else if (op === 'date_gte' || op === '>=') {
              if (cellTs < targetTs) return false;
            } else if (op === 'date_lte' || op === '<=') {
              if (cellTs > targetTs) return false;
            } else if (op === '>') {
              if (cellTs <= targetTs) return false;
            } else if (op === '<') {
              if (cellTs >= targetTs) return false;
            }
            continue;
          }

          const strCell = String(cell).toLowerCase();
          const strTarget = String(target).toLowerCase();

          if (op === 'contains') {
            if (!strCell.includes(strTarget)) return false;
            continue;
          } else if (op === 'starts_with') {
            if (!strCell.startsWith(strTarget)) return false;
            continue;
          } else if (op === 'ends_with') {
            if (!strCell.endsWith(strTarget)) return false;
            continue;
          } else if (op === 'not_contains') {
            if (strCell.includes(strTarget)) return false;
            continue;
          }

          const numCell = Number(cell);
          const numTarget = Number(target);
          const isBothNumber = !isNaN(numCell) && !isNaN(numTarget) && String(target).trim() !== '';

          if (op === '>') {
            if (isBothNumber ? !(numCell > numTarget) : !(String(cell) > String(target))) return false;
          } else if (op === '>=') {
            if (isBothNumber ? !(numCell >= numTarget) : !(String(cell) >= String(target))) return false;
          } else if (op === '<') {
            if (isBothNumber ? !(numCell < numTarget) : !(String(cell) < String(target))) return false;
          } else if (op === '<=') {
            if (isBothNumber ? !(numCell <= numTarget) : !(String(cell) <= String(target))) return false;
          } else if (op === '!=') {
            if (isBothNumber ? (numCell === numTarget) : (String(cell).trim().toLowerCase() === String(target).trim().toLowerCase())) return false;
          } else {
            if (isBothNumber ? (numCell !== numTarget) : (String(cell).trim().toLowerCase() !== String(target).trim().toLowerCase())) return false;
          }
        }
        return true;
      });
    }

    function getFilteredFilesObject(targetTabId = null) {
      const filtered = {};
      const tabId = targetTabId || activeTabId || 'default';
      for (const f in FILES) {
        filtered[f] = {};
        for (const s in FILES[f]) {
          filtered[f][s] = getFilteredData(f + '::' + s, [], tabId);
        }
      }
      return filtered;
    }

    function optimizeChartOptionForLargeData(option) {
      if (!option || typeof option !== 'object') return option;

      let hasHeavyData = false;
      const seriesList = Array.isArray(option.series) 
        ? option.series 
        : (option.series && typeof option.series === 'object' ? [option.series] : []);

      seriesList.forEach(s => {
        if (!s || typeof s !== 'object') return;
        const dataLen = Array.isArray(s.data) ? s.data.length : 0;
        if (dataLen >= 300) {
          hasHeavyData = true;
          if (s.large === undefined && (s.type === 'line' || s.type === 'bar' || s.type === 'scatter')) {
            s.large = true;
            s.largeThreshold = 300;
          }
          if (s.type === 'line' && !s.sampling) {
            s.sampling = 'lttb';
          }
          if (dataLen >= 1500 && s.progressive === undefined) {
            s.progressive = 800;
            s.progressiveThreshold = 1500;
          }
        }
      });

      if (hasHeavyData && option.animation === undefined) {
        option.animation = false;
      }

      return option;
    }

    function enhanceInteractiveGraphic(option) {
      if (!option || typeof option !== 'object' || !option.graphic) return option;
      const processItem = (g) => {
        if (!g || typeof g !== 'object') return;
        const hasInteraction = typeof g.onclick === 'function' || typeof g.onmouseover === 'function' || typeof g.onmousedown === 'function';
        if (hasInteraction) {
          if (g.cursor === undefined) g.cursor = 'pointer';
          if (!g.style) g.style = {};
          if (g.style.cursor === undefined) g.style.cursor = 'pointer';
          if (g.z === undefined) g.z = 100;
          if (g.silent === true) g.silent = false;
        }
        if (Array.isArray(g.children)) {
          g.children.forEach(processItem);
        }
      };

      if (Array.isArray(option.graphic)) {
        option.graphic.forEach(processItem);
      } else if (typeof option.graphic === 'object') {
        processItem(option.graphic);
      }
      return option;
    }

    function evaluateEChartsCode(code, optionalData = null, context = {}, chartInstance = null) {
      if (!code || !code.trim()) {
        return { option: {}, error: 'Chưa có code JavaScript cấu hình.' };
      }
      try {
        const itemTabId = context.tabId || activeTabId || 'default';
        const allFiles = FILES;
        const filteredFiles = getFilteredFilesObject(itemTabId);
        let defaultData = optionalData;
        if (defaultData === null || defaultData === undefined) {
          defaultData = getFilteredData(null, [], itemTabId);
        }
        const getSheet = (sheetName, fileName = null) => {
          if (fileName && filteredFiles[fileName] && filteredFiles[fileName][sheetName]) {
            return filteredFiles[fileName][sheetName];
          }
          for (const f in filteredFiles) {
            if (filteredFiles[f][sheetName]) return filteredFiles[f][sheetName];
          }
          return [];
        };
        const getRawSheet = (sheetName, fileName = null) => {
          if (fileName && allFiles[fileName] && allFiles[fileName][sheetName]) {
            return allFiles[fileName][sheetName];
          }
          for (const f in allFiles) {
            if (allFiles[f][sheetName]) return allFiles[f][sheetName];
          }
          return [];
        };
        const getFile = (fileName) => {
          return filteredFiles[fileName] || null;
        };

        const setFilter = (col, val) => {
          if (!col) return;
          if (!tabFilters[itemTabId]) tabFilters[itemTabId] = {};
          if (val === undefined || val === null || val === '' || val === 'all') {
            delete tabFilters[itemTabId][col];
            if (itemTabId === activeTabId) delete activeFilters[col];
          } else {
            tabFilters[itemTabId][col] = val;
            if (itemTabId === activeTabId) activeFilters[col] = val;
          }
          renderCanvas();
        };

        const resetFilter = (col = null) => {
          if (col) {
            setFilter(col, null);
          } else {
            tabFilters[itemTabId] = {};
            if (itemTabId === activeTabId) activeFilters = {};
            renderCanvas();
          }
        };

        const targetFilters = itemTabId === activeTabId ? activeFilters : (tabFilters[itemTabId] || {});
        const safeContext = {
          setFilter,
          resetFilter,
          activeFilters: targetFilters,
          rawFiles: allFiles,
          getRawSheet,
          tabId: itemTabId,
          myChart: chartInstance,
          chart: chartInstance,
          ...context
        };

        const evalFunc = new Function(
          'data', 'files', 'getSheet', 'getFile', 'context', 'echarts', 'XLSX', 'console', 'myChart', 'chart', 'setFilter', 'resetFilter',
          'let option = null; try { ' + code + ' } catch (e) { return { __error: e.stack || e.message }; } if (typeof option === "object" && option !== null) return option;'
        );

        let result = evalFunc(defaultData, filteredFiles, getSheet, getFile, safeContext, window.echarts, window.XLSX, window.console, chartInstance, chartInstance, setFilter, resetFilter);
        if (result && result.__error) {
          return { option: {}, error: result.__error };
        }
        if ((!result || typeof result !== 'object') && chartInstance && chartInstance.getOption) {
          const curOpt = chartInstance.getOption();
          if (curOpt && Object.keys(curOpt).length > 0) {
            result = curOpt;
          }
        }
        if (!result || typeof result !== 'object') {
          return { option: {}, error: 'Code phải return về một object option ECharts hoặc gán myChart.setOption(option);' };
        }
        enhanceInteractiveGraphic(result);
        const optimizedOption = optimizeChartOptionForLargeData(result);
        return { option: optimizedOption, error: null };
      } catch (err) {
        return { option: {}, error: err.message || String(err) };
      }
    }

    function renderAllTabs() {
      const tabsList = document.getElementById('dbTabsList');
      if (!tabsList) return;
      tabsList.innerHTML = '';
      const cfg = CONFIG.tabBarConfig || {};
      const accent = cfg.accentColor || '#0284c7';

      (CONFIG.tabs || []).forEach((tab, index) => {
        const pill = document.createElement('button');
        const isActive = tab.id === activeTabId;
        pill.className = 'db-tab-pill ' + (isActive ? 'active' : '');
        if (isActive && accent) {
          pill.style.backgroundColor = accent;
          pill.style.borderColor = accent;
        }
        let inner = '';
        if (cfg.showIndex) {
          inner += '<span class="tab-index-badge">' + (index + 1) + '</span>';
        }
        inner += '<span class="tab-name-label">' + escapeHtml(tab.name) + '</span>';
        pill.innerHTML = inner;

        pill.addEventListener('click', () => {
          switchDashboardTab(tab.id);
        });
        tabsList.appendChild(pill);
      });
    }

    function renderCanvas() {
      const canvas = document.getElementById('dashboardCanvas');
      if (!canvas) return;
      canvas.innerHTML = '';

      // 2 Vạch Ranh Giới Dashboard
      const leftGuide = document.createElement('div');
      leftGuide.className = 'db-canvas-boundary-guide is-left';
      canvas.appendChild(leftGuide);

      const rightGuide = document.createElement('div');
      rightGuide.className = 'db-canvas-boundary-guide is-right';
      canvas.appendChild(rightGuide);

      const currentTabItems = (CONFIG.items || []).filter(i => i.tabId === activeTabId);
      const pinnedFilterItems = currentTabItems.filter(i => i.type === 'filter' && (i.isPinnedTop || i.container === 'sticky_top'));
      const freeCanvasItems = currentTabItems.filter(i => !(i.type === 'filter' && (i.isPinnedTop || i.container === 'sticky_top')));

      const barConfig = (CONFIG.topFilterBarConfig && CONFIG.topFilterBarConfig[activeTabId]) || {
        enabled: false,
        isSticky: true,
        isCollapsed: false,
        title: 'BỘ LỌC TỔNG HỢP',
        pos: { x: 24, y: 24, w: 1312 }
      };

      // 1. Render Thanh Bộ Lọc Tổng Hợp (nếu có bộ lọc được ghim)
      if (pinnedFilterItems.length > 0) {
        const topBarEl = document.createElement('div');
        const isCollapsed = Boolean(barConfig.isCollapsed);
        const collapseStyle = barConfig.collapseStyle || 'circle';
        const isCircle = collapseStyle === 'circle';
        topBarEl.id = 'dbTopFilterBar';

        topBarEl.className = 'db-top-filter-bar is-movable ' + (isCollapsed ? 'is-collapsed ' : '') + (isCircle ? 'collapse-style-circle' : 'collapse-style-bar');
        topBarEl.style.position = 'absolute';
        topBarEl.style.left = (barConfig.pos?.x || 24) + 'px';
        topBarEl.style.top = (barConfig.pos?.y || 24) + 'px';
        topBarEl.style.width = isCollapsed && isCircle ? '48px' : ((barConfig.pos?.w || 1312) + 'px');
        topBarEl.style.zIndex = '900';

        const countLabel = pinnedFilterItems.length + ' bộ lọc';
        topBarEl.innerHTML = '<div class="db-top-filter-circle-btn" id="btnTopFilterCircleOpen" title="🔍 Mở Bộ Lọc Tổng Hợp (' + countLabel + ')">' +
            '<span class="db-filter-circle-icon">🔍</span>' +
            '<span class="db-filter-circle-badge">' + pinnedFilterItems.length + '</span>' +
          '</div>' +
          '<div class="db-top-filter-expanded-wrap">' +
            '<div class="db-top-filter-header">' +
              '<div class="db-top-filter-title-group">' +
                '<span class="db-top-filter-icon">🔍</span>' +
                '<span class="db-top-filter-title">' + escapeHtml(barConfig.title || 'BỘ LỌC TỔNG HỢP') + '</span>' +
                '<span class="db-top-filter-count-badge">📌 ' + countLabel + '</span>' +
              '</div>' +
              '<div class="db-top-filter-actions">' +
                '<button type="button" class="db-top-filter-btn primary" id="btnToggleCollapseTopBar" title="Mở rộng hoặc Thu gọn">' +
                  (isCircle ? '▲ Thu Gọn (Nút Tròn)' : (isCollapsed ? '▼ Mở Bộ Lọc' : '▲ Thu Gọn')) +
                '</button>' +
              '</div>' +
            '</div>' +
            '<div class="db-top-filter-body" id="dbTopFilterBody"></div>' +
          '</div>';

        const topBodyEl = topBarEl.querySelector('#dbTopFilterBody');

        pinnedFilterItems.forEach(item => {
          const wrapper = document.createElement('div');
          wrapper.className = 'db-top-filter-item-wrapper';
          wrapper.id = 'pinned_item_' + item.id;
          wrapper.innerHTML = '<div class="db-top-filter-item-header">' +
            '<span class="db-top-filter-item-title">' + escapeHtml(item.name || item.filterColumn) + '</span>' +
          '</div>' +
          '<div class="db-item-body" id="item_body_' + item.id + '"></div>';
          topBodyEl.appendChild(wrapper);
        });

        canvas.appendChild(topBarEl);

        topBarEl.dataset.origTop = topBarEl.offsetTop;

        pinnedFilterItems.forEach(item => {
          renderItemContent(item);
        });

        topBarEl.querySelector('#btnTopFilterCircleOpen')?.addEventListener('click', (e) => {
          e.stopPropagation();
          barConfig.isCollapsed = false;
          topBarEl.classList.remove('is-collapsed');
          topBarEl.style.width = (barConfig.pos?.w || 1312) + 'px';
        });

        topBarEl.querySelector('#btnToggleCollapseTopBar')?.addEventListener('click', (e) => {
          e.stopPropagation();
          barConfig.isCollapsed = !barConfig.isCollapsed;
          topBarEl.classList.toggle('is-collapsed', barConfig.isCollapsed);
          if (isCircle) {
            topBarEl.style.width = barConfig.isCollapsed ? '48px' : ((barConfig.pos?.w || 1312) + 'px');
          }
          const btn = topBarEl.querySelector('#btnToggleCollapseTopBar');
          if (btn) btn.textContent = isCircle ? '▲ Thu Gọn (Nút Tròn)' : (barConfig.isCollapsed ? '▼ Mở Bộ Lọc' : '▲ Thu Gọn');
        });
      }

      let maxBottom = 800;
      freeCanvasItems.forEach(item => {
        const b = (item.pos?.y || 0) + (item.pos?.h || 0);
        if (b > maxBottom) maxBottom = b;
      });
      const finalH = maxBottom + 60;
      canvas.style.minHeight = finalH + 'px';
      canvas.dataset.originalHeight = finalH;

      freeCanvasItems.forEach(item => {
        const el = document.createElement('div');
        const isHideHeader = Boolean(item.hideHeader);
        const isFramelessType = ['shape', 'heading', 'text', 'image'].includes(item.type);
        const shapeSubclass = item.type === 'shape' ? (item.shapeType === 'line' ? 'is-shape-line' : 'is-shape') : '';

        el.className = 'db-item-box ' +
          (item.type === 'frame' ? 'is-frame ' : '') +
          (shapeSubclass ? shapeSubclass + ' ' : '') +
          (isFramelessType ? 'is-frameless-element ' : '') +
          ('is-type-' + item.type + ' ') +
          (isHideHeader ? 'hide-header ' : '');

        el.id = 'item_' + item.id;
        el.style.left = (item.pos?.x || 24) + 'px';
        el.style.top = (item.pos?.y || 24) + 'px';
        el.style.width = (item.pos?.w || 400) + 'px';
        el.style.height = (item.pos?.h || 300) + 'px';
        if (item.pos?.zIndex) el.style.zIndex = item.pos.zIndex;

        if (isFramelessType) {
          el.innerHTML = '<div class="db-item-body" id="item_body_' + item.id + '" style="padding:0;height:100%;"></div>';
        } else {
          el.innerHTML = '<div class="db-item-header"><div class="db-item-header-left"><span class="db-item-title">' + escapeHtml(item.name || item.type) + '</span></div></div><div class="db-item-body" id="item_body_' + item.id + '"></div>';
        }

        canvas.appendChild(el);
        renderItemContent(item);
      });

      updateCanvasTransform();
    }

    function renderItemContent(item) {
      const body = document.getElementById('item_body_' + item.id);
      if (!body) return;
      body.innerHTML = '';

      if (item.type === 'chart') {
        const chartDiv = document.createElement('div');
        chartDiv.className = 'db-chart-canvas';
        chartDiv.id = 'chart_el_' + item.id;
        body.appendChild(chartDiv);

        setTimeout(() => {
          let chart = echarts.getInstanceByDom(chartDiv);
          if (chart) chart.dispose();
          chart = echarts.init(chartDiv, null);

          const itemTabId = item.tabId || activeTabId || 'default';
          const data = getFilteredData(item.dataSourceKey, [], itemTabId);
          const { option, error } = evaluateEChartsCode(item.code, data, { tabId: itemTabId }, chart);

          if (error) {
            chartDiv.innerHTML = '<div class="db-chart-error-banner">⚠️ Lỗi render ECharts:<br/>' + escapeHtml(error) + '</div>';
          } else if (option && Object.keys(option).length > 0) {
            chart.setOption(option, true);
          }
        }, 30);
      } else if (item.type === 'shape') {
        const sType = item.shapeType || 'rectangle';
        const bg = item.bgColor || '#ffffff';
        const borderW = item.borderWidth !== undefined ? item.borderWidth : 1;
        const borderS = item.borderStyle || 'solid';
        const borderC = item.borderColor || '#cbd5e1';
        const border = borderW > 0 && borderS !== 'none' ? (borderW + 'px ' + borderS + ' ' + borderC) : 'none';
        const opacity = item.opacity !== undefined ? (item.opacity / 100) : 1;
        const shadow = item.boxShadow || 'none';

        let radiusCss = (item.borderRadius !== undefined ? item.borderRadius : 8) + 'px';
        if (item.radiusTL !== undefined) {
          radiusCss = (item.radiusTL || 0) + 'px ' + (item.radiusTR || 0) + 'px ' + (item.radiusBR || 0) + 'px ' + (item.radiusBL || 0) + 'px';
        }
        if (sType === 'circle') radiusCss = '50%';

        if (sType === 'line') {
          const thickness = item.borderWidth || 2;
          const color = item.borderColor || '#94a3b8';
          const style = item.borderStyle || 'solid';
          const cap = item.lineCap === 'round' ? '999px' : '0px';
          const isVert = item.lineOrientation === 'vertical';

          if (isVert) {
            body.innerHTML = '<div style="width:100%;height:100%;display:flex;justify-content:center;align-items:center;padding:4px 0;"><div style="width:' + thickness + 'px;height:100%;background:' + (style === 'solid' ? color : 'transparent') + ';border-left:' + (style !== 'solid' ? (thickness + 'px ' + style + ' ' + color) : 'none') + ';border-radius:' + cap + ';box-shadow:' + shadow + ';opacity:' + opacity + ';"></div></div>';
          } else {
            body.innerHTML = '<div style="width:100%;height:100%;display:flex;align-items:center;padding:0 4px;"><div style="width:100%;height:' + thickness + 'px;background:' + (style === 'solid' ? color : 'transparent') + ';border-top:' + (style !== 'solid' ? (thickness + 'px ' + style + ' ' + color) : 'none') + ';border-radius:' + cap + ';box-shadow:' + shadow + ';opacity:' + opacity + ';"></div></div>';
          }
        } else {
          body.innerHTML = '<div class="db-shape-container" style="background:' + bg + ';border:' + border + ';border-radius:' + radiusCss + ';opacity:' + opacity + ';box-shadow:' + shadow + ';"></div>';
        }
      } else if (item.type === 'text') {
        const fFamily = item.fontFamily || 'Inter, sans-serif';
        const fSize = item.fontSize || '14px';
        const fWeight = item.fontWeight || 'normal';
        const fStyle = item.fontStyle || 'normal';
        const tDeco = item.textDecoration || 'none';
        const tColor = item.color || '#0f172a';
        const tAlign = item.textAlign || 'left';
        const textVal = item.text !== undefined ? item.text : (item.name || '');

        body.innerHTML = '<div style="height:100%;width:100%;padding:4px;overflow:hidden;display:flex;flex-direction:column;justify-content:flex-start;"><div class="db-text-container" style="font-family:' + fFamily + ';font-size:' + fSize + ';font-weight:' + fWeight + ';font-style:' + fStyle + ';text-decoration:' + tDeco + ';color:' + tColor + ';text-align:' + tAlign + ';white-space:pre-wrap;word-break:break-word;line-height:1.5;width:100%;min-height:100%;">' + escapeHtml(textVal) + '</div></div>';
      } else if (item.type === 'heading') {
        const fFamily = item.fontFamily || 'Inter, sans-serif';
        const fSize = item.fontSize || '24px';
        const fWeight = item.fontWeight || '700';
        const fStyle = item.fontStyle || 'normal';
        const tDeco = item.textDecoration || 'none';
        const tColor = item.color || '#0f172a';
        const tAlign = item.textAlign || 'left';
        const textVal = item.text !== undefined ? item.text : (item.name || 'Tiêu Đề');

        body.innerHTML = '<div style="height:100%;width:100%;display:flex;align-items:center;padding:4px;overflow:hidden;"><h2 style="margin:0;width:100%;font-family:' + fFamily + ';font-size:' + fSize + ';font-weight:' + fWeight + ';font-style:' + fStyle + ';text-decoration:' + tDeco + ';color:' + tColor + ';text-align:' + tAlign + ';word-break:break-word;line-height:1.3;">' + escapeHtml(textVal) + '</h2></div>';
      } else if (item.type === 'kpi') {
        const itemTabId = item.tabId || activeTabId || 'default';
        const data = getFilteredData(item.dataSourceKey, [], itemTabId);
        const numVals = data.map(r => Number(r[item.metricColumn])).filter(n => !isNaN(n));
        let sum = 0;
        if (item.aggregation === 'avg') {
          sum = numVals.length > 0 ? (numVals.reduce((a, b) => a + b, 0) / numVals.length) : 0;
        } else if (item.aggregation === 'count') {
          sum = numVals.length;
        } else if (item.aggregation === 'max') {
          sum = numVals.length > 0 ? Math.max(...numVals) : 0;
        } else if (item.aggregation === 'min') {
          sum = numVals.length > 0 ? Math.min(...numVals) : 0;
        } else {
          sum = numVals.reduce((a, b) => a + b, 0);
        }

        const kpiFs = item.fontSize || '30px';
        const kpiColor = item.color || '#0f172a';
        body.innerHTML = '<div style="display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;text-align:center;"><div style="font-size:11px;text-transform:uppercase;color:var(--db-text-sub);font-weight:700;margin-bottom:4px;">' + escapeHtml(item.name || 'KPI') + '</div><div style="font-size:' + escapeHtml(kpiFs) + ';font-weight:800;color:' + escapeHtml(kpiColor) + ';">' + escapeHtml(item.prefix || '') + new Intl.NumberFormat('vi-VN').format(Math.round(sum)) + escapeHtml(item.suffix || '') + '</div></div>';
      } else if (item.type === 'button') {
        const btnEl = document.createElement('div');
        btnEl.className = 'db-action-btn-widget';
        btnEl.style.cssText = 'background:' + (item.btnBg || '#004ea8') + ';color:' + (item.btnColor || '#ffffff') + ';border-radius:' + (item.borderRadius || 8) + 'px;font-size:' + (item.fontSize || '14px') + ';font-weight:700;box-shadow:0 2px 6px rgba(0,0,0,0.08);';
        btnEl.innerHTML = '<span>' + escapeHtml(item.name || 'Nút Bấm') + '</span>';
        btnEl.addEventListener('click', () => {
          if (item.btnAction === 'url' && item.btnTarget) {
            window.open(item.btnTarget, '_blank');
          } else if (item.btnAction === 'tab' && item.btnTarget) {
            switchDashboardTab(item.btnTarget);
          }
        });
        body.appendChild(btnEl);
      } else if (item.type === 'frame') {
        const bg = item.bgColor || 'rgba(255,255,255,0.4)';
        const borderW = item.borderWidth !== undefined ? item.borderWidth : 2;
        const borderS = item.borderStyle || 'dashed';
        const borderC = item.borderColor || '#cbd5e1';
        const border = borderW > 0 && borderS !== 'none' ? (borderW + 'px ' + borderS + ' ' + borderC) : 'none';
        const radius = (item.borderRadius !== undefined ? item.borderRadius : 12) + 'px';
        const shadow = item.boxShadow || 'none';
        const opacity = item.opacity !== undefined ? (item.opacity / 100) : 1;

        body.innerHTML = '<div class="db-frame-container" style="background:' + bg + ';border:' + border + ';border-radius:' + radius + ';box-shadow:' + shadow + ';opacity:' + opacity + ';padding:8px;font-size:12px;color:var(--db-text-sub);font-style:italic;">[Khung phân vùng: ' + escapeHtml(item.name) + ']</div>';
      } else if (item.type === 'image') {
        body.innerHTML = '<div class="db-image-widget"><img src="' + (item.imageUrl || '') + '" alt="' + escapeHtml(item.name || 'Ảnh') + '" style="border-radius:' + (item.borderRadius || 6) + 'px;object-fit:' + (item.fit || 'contain') + ';" /></div>';
      } else if (item.type === 'filter') {
        const itemTabId = item.tabId || activeTabId || 'default';
        const col = item.filterColumn;
        if (!item.dataSourceKey && col) {
          for (const f in FILES) {
            for (const s in FILES[f]) {
              const r = FILES[f][s];
              if (r && r.length > 0 && r[0] && r[0][col] !== undefined) {
                item.dataSourceKey = f + '::' + s;
                break;
              }
            }
            if (item.dataSourceKey) break;
          }
        }
        const allRows = getData(item.dataSourceKey, col);
        const isDateCol = isProbableDateColumn(col, allRows);

        let fType = item.filterType;
        if (!fType) {
          if (item.operator === 'between' || item.operator === 'range') fType = 'range';
          else if (item.operator === 'in') fType = 'multiselect';
          else if (item.operator === 'date_range' || String(item.operator).startsWith('date_') || isDateCol) fType = 'date';
          else if (['contains', 'starts_with', 'ends_with', 'not_contains'].includes(item.operator)) fType = 'text';
          else if (['>', '>=', '<', '<=', '!='].includes(item.operator)) fType = 'comparison';
          else fType = isDateCol ? 'date' : 'range';
        }

        let parentItem = null;
        if (item.parentFilterId) {
          parentItem = (CONFIG.items || []).find(i => i.id === item.parentFilterId);
        }

        if (!tabFilters[itemTabId]) tabFilters[itemTabId] = {};
        const currentTabFilters = itemTabId === activeTabId ? activeFilters : (tabFilters[itemTabId] || {});
        const filterState = currentTabFilters[col];

        const setTabFilterState = (valObj, keepDropdownOpen = false) => {
          if (!tabFilters[itemTabId]) tabFilters[itemTabId] = {};
          if (valObj === undefined || valObj === null || valObj === '' || valObj === 'all') {
            delete tabFilters[itemTabId][col];
            if (itemTabId === activeTabId) delete activeFilters[col];
          } else {
            tabFilters[itemTabId][col] = valObj;
            if (itemTabId === activeTabId) activeFilters[col] = valObj;
          }
          if (keepDropdownOpen) {
            openDropdownFilterId = item.id;
          } else {
            openDropdownFilterId = null;
          }
          renderCanvas();
        };

        let currentOp = item.operator;
        if (!currentOp) {
          if (fType === 'range') currentOp = 'between';
          else if (fType === 'multiselect') currentOp = 'in';
          else if (fType === 'date') currentOp = 'date_range';
          else if (fType === 'text') currentOp = 'contains';
          else if (fType === 'comparison') currentOp = '>';
          else currentOp = '=';
        }

        let currentVal = 'all';
        let minVal = '';
        let maxVal = '';
        let startDate = '';
        let endDate = '';
        let selectedValues = [];

        if (filterState) {
          if (typeof filterState === 'object' && filterState !== null) {
            currentOp = filterState.op || currentOp;
            currentVal = filterState.val !== undefined ? filterState.val : 'all';
            minVal = filterState.min !== undefined ? filterState.min : (filterState.start !== undefined ? filterState.start : '');
            maxVal = filterState.max !== undefined ? filterState.max : (filterState.end !== undefined ? filterState.end : '');
            startDate = filterState.start || filterState.min || '';
            endDate = filterState.end || filterState.max || '';
            selectedValues = Array.isArray(filterState.values) ? filterState.values : (Array.isArray(filterState.val) ? filterState.val : (filterState.val && filterState.val !== 'all' ? [filterState.val] : []));
          } else {
            currentVal = filterState;
          }
        }

        let filterSourceRows = allRows;
        if (parentItem && parentItem.filterColumn) {
          const parentCol = parentItem.filterColumn;
          const parentVal = (tabFilters[itemTabId] || {})[parentCol] || activeFilters[parentCol];
          if (parentVal) {
            filterSourceRows = getFilteredData(item.dataSourceKey, Object.keys(activeFilters).filter(k => k !== parentCol), itemTabId);
          }
        }
        const uniqueVals = Array.from(new Set((filterSourceRows || allRows).map(r => r[col]).filter(v => v !== null && v !== undefined && v !== ''))).sort((a, b) => {
          const na = Number(a), nb = Number(b);
          if (!isNaN(na) && !isNaN(nb)) return na - nb;
          return String(a).localeCompare(String(b));
        });

        if (Array.isArray(selectedValues)) {
          selectedValues.forEach(sv => {
            if (sv !== null && sv !== undefined && sv !== '' && !uniqueVals.map(String).includes(String(sv))) {
              uniqueVals.push(sv);
            }
          });
        }

        if (fType === 'select' && currentVal !== 'all' && currentVal !== '' && typeof currentVal === 'string') {
          if (uniqueVals.length > 0 && !uniqueVals.some(v => String(v) === String(currentVal))) {
            setTabFilterState(null);
            currentVal = 'all';
          }
        }

        const isDate = fType === 'date' || (fType === 'range' && isDateCol);
        const isNumberCol = !isDate && uniqueVals.some(v => !isNaN(Number(v)) && String(v).trim() !== '');

        let isFiltering = false;
        let activeBadgeHtml = '';

        if (fType === 'range') {
          if (isDate) {
            isFiltering = Boolean(startDate || endDate);
            let bStart = startDate;
            let bEnd = endDate;
            if (startDate && endDate) {
              const sTs = parseDateToTimestamp(startDate);
              const eTs = parseDateToTimestamp(endDate);
              if (sTs !== null && eTs !== null && sTs > eTs) {
                bStart = endDate;
                bEnd = startDate;
              }
            }
            if (isFiltering) activeBadgeHtml = '<span class="db-filter-active-badge">📅 [' + (bStart || '...') + ' ➔ ' + (bEnd || '...') + ']</span>';
          } else {
            isFiltering = Boolean(minVal !== '' || maxVal !== '');
            let bMin = minVal;
            let bMax = maxVal;
            if (minVal !== '' && maxVal !== '' && !isNaN(Number(minVal)) && !isNaN(Number(maxVal)) && Number(minVal) > Number(maxVal)) {
              bMin = maxVal;
              bMax = minVal;
            }
            if (isFiltering) activeBadgeHtml = '<span class="db-filter-active-badge">↔️ [' + (bMin !== '' ? bMin : '-∞') + ' ➔ ' + (bMax !== '' ? bMax : '+∞') + ']</span>';
          }
        } else if (fType === 'multiselect') {
          isFiltering = Boolean(selectedValues.length > 0 && selectedValues.length < uniqueVals.length);
          if (isFiltering) activeBadgeHtml = '<span class="db-filter-active-badge">☑️ (' + selectedValues.length + ' mục)</span>';
        } else if (fType === 'date') {
          if (currentOp === 'date_range') {
            isFiltering = Boolean(startDate || endDate);
            let bStart = startDate;
            let bEnd = endDate;
            if (startDate && endDate) {
              const sTs = parseDateToTimestamp(startDate);
              const eTs = parseDateToTimestamp(endDate);
              if (sTs !== null && eTs !== null && sTs > eTs) {
                bStart = endDate;
                bEnd = startDate;
              }
            }
            if (isFiltering) activeBadgeHtml = '<span class="db-filter-active-badge">📅 [' + (bStart || '...') + ' ➔ ' + (bEnd || '...') + ']</span>';
          } else {
            isFiltering = Boolean(currentVal !== 'all' && currentVal !== '');
            if (isFiltering) activeBadgeHtml = '<span class="db-filter-active-badge">📅 ' + currentOp + ' ' + escapeHtml(currentVal) + '</span>';
          }
        } else {
          isFiltering = Boolean(currentVal !== 'all' && currentVal !== '' && currentVal !== null && currentVal !== undefined);
          if (isFiltering) activeBadgeHtml = '<span class="db-filter-active-badge">' + escapeHtml(currentOp) + ' ' + escapeHtml(currentVal) + '</span>';
        }

        let parentBadgeHtml = '';
        if (parentItem) {
          parentBadgeHtml = '<span class="db-filter-parent-badge" title="Bộ lọc phụ thuộc vào: ' + escapeHtml(parentItem.name || parentItem.filterColumn) + '">↳ Con: ' + escapeHtml(parentItem.name || parentItem.filterColumn) + '</span>';
        }

        let controlsHtml = '';
        if (fType === 'range') {
          if (isDate) {
            controlsHtml = '<div class="db-filter-date-range-wrap">' +
              '<input type="date" class="db-filter-date-input" id="filter_date_start_' + item.id + '" value="' + formatDateToYMD(startDate) + '" title="Từ ngày" />' +
              '<span class="db-filter-date-sep">➔</span>' +
              '<input type="date" class="db-filter-date-input" id="filter_date_end_' + item.id + '" value="' + formatDateToYMD(endDate) + '" title="Đến ngày" />' +
              '</div>';
          } else {
            controlsHtml = '<div class="db-filter-range-wrap">' +
              '<input type="' + (isNumberCol ? 'number' : 'text') + '" class="db-filter-range-input" id="filter_range_min_' + item.id + '" placeholder="Từ (Min)..." value="' + escapeHtml(minVal) + '" title="Giá trị nhỏ nhất" />' +
              '<span class="db-filter-range-sep">➔</span>' +
              '<input type="' + (isNumberCol ? 'number' : 'text') + '" class="db-filter-range-input" id="filter_range_max_' + item.id + '" placeholder="Đến (Max)..." value="' + escapeHtml(maxVal) + '" title="Giá trị lớn nhất" />' +
              '</div>';
          }
        } else if (fType === 'comparison') {
          controlsHtml = '<select class="db-filter-op-select" id="filter_op_' + item.id + '" title="Chọn phép so sánh" style="width:84px;">' +
            '<option value=">" ' + (currentOp === '>' ? 'selected' : '') + '>&gt; Lớn hơn</option>' +
            '<option value=">=" ' + (currentOp === '>=' ? 'selected' : '') + '>&ge; Lớn/bằng</option>' +
            '<option value="<" ' + (currentOp === '<' ? 'selected' : '') + '>&lt; Bé hơn</option>' +
            '<option value="<=" ' + (currentOp === '<=' ? 'selected' : '') + '>&le; Bé/bằng</option>' +
            '<option value="=" ' + (currentOp === '=' ? 'selected' : '') + '>= Bằng</option>' +
            '<option value="!=" ' + (currentOp === '!=' ? 'selected' : '') + '>≠ Khác</option>' +
            '</select>' +
            '<div style="flex:1;display:flex;gap:4px;min-width:0;">' +
            '<input type="' + (isNumberCol ? 'number' : 'text') + '" class="db-filter-val-input" id="filter_val_' + item.id + '" placeholder="' + (isNumberCol ? 'Nhập số mốc...' : 'Giá trị...') + '" value="' + (currentVal !== 'all' ? escapeHtml(currentVal) : '') + '" list="filter_datalist_' + item.id + '" />' +
            '<datalist id="filter_datalist_' + item.id + '">' +
            uniqueVals.slice(0, 30).map(v => '<option value="' + escapeHtml(v) + '"></option>').join('') +
            '</datalist></div>';
        } else if (fType === 'multiselect') {
          const countLabel = selectedValues.length === 0 ? ('Tất cả (' + uniqueVals.length + ')') : ('Đã chọn (' + selectedValues.length + ')');
          controlsHtml = '<div class="db-filter-multiselect-container" id="filter_ms_container_' + item.id + '">' +
            '<button type="button" class="db-filter-multiselect-btn" id="filter_ms_btn_' + item.id + '" title="Nhấp để chọn một hoặc nhiều mục">' +
            '<span id="filter_ms_label_' + item.id + '">☑️ ' + escapeHtml(countLabel) + '</span>' +
            '<span>▾</span></button>' +
            '<div class="db-filter-multiselect-dropdown" id="filter_ms_drop_' + item.id + '">' +
            '<input type="text" class="db-filter-multiselect-search" id="filter_ms_search_' + item.id + '" placeholder="🔍 Tìm giá trị..." />' +
            '<div class="db-filter-multiselect-actions">' +
            '<button type="button" class="db-filter-ms-action-btn" id="filter_ms_all_' + item.id + '">Chọn tất cả</button>' +
            '<button type="button" class="db-filter-ms-action-btn" id="filter_ms_none_' + item.id + '">Bỏ chọn</button>' +
            '</div><div class="db-filter-multiselect-list" id="filter_ms_list_' + item.id + '">' +
            uniqueVals.map(v => '<label class="db-filter-multiselect-item"><input type="checkbox" value="' + escapeHtml(v) + '" ' + (selectedValues.map(String).includes(String(v)) ? 'checked' : '') + ' class="filter_ms_chk_' + item.id + '" /><span>' + escapeHtml(v) + '</span></label>').join('') +
            '</div></div></div>';
        } else if (fType === 'text') {
          controlsHtml = '<select class="db-filter-op-select" id="filter_op_' + item.id + '" title="Chọn điều kiện tìm kiếm" style="width:84px;">' +
            '<option value="contains" ' + (currentOp === 'contains' ? 'selected' : '') + '>Chứa</option>' +
            '<option value="starts_with" ' + (currentOp === 'starts_with' ? 'selected' : '') + '>Bắt đầu</option>' +
            '<option value="ends_with" ' + (currentOp === 'ends_with' ? 'selected' : '') + '>Kết thúc</option>' +
            '<option value="not_contains" ' + (currentOp === 'not_contains' ? 'selected' : '') + '>K.chứa</option>' +
            '<option value="=" ' + (currentOp === '=' ? 'selected' : '') + '>= Chuẩn</option>' +
            '</select>' +
            '<div style="flex:1;display:flex;gap:4px;min-width:0;">' +
            '<input type="text" class="db-filter-val-input" id="filter_val_' + item.id + '" placeholder="Nhập từ khóa..." value="' + (currentVal !== 'all' ? escapeHtml(currentVal) : '') + '" />' +
            '</div>';
        } else if (fType === 'date') {
          controlsHtml = '<select class="db-filter-op-select" id="filter_op_' + item.id + '" title="Chọn kiểu lọc ngày" style="width:105px;">' +
            '<option value="date_range" ' + (currentOp === 'date_range' ? 'selected' : '') + '>📅 Khoảng</option>' +
            '<option value="date_eq" ' + (currentOp === 'date_eq' ? 'selected' : '') + '>🎯 Đúng ngày</option>' +
            '<option value="date_gte" ' + (currentOp === 'date_gte' ? 'selected' : '') + '>⏩ Từ ngày</option>' +
            '<option value="date_lte" ' + (currentOp === 'date_lte' ? 'selected' : '') + '>⏪ Đến ngày</option>' +
            '</select>' +
            (currentOp === 'date_range' ? (
              '<div class="db-filter-date-range-wrap">' +
              '<input type="date" class="db-filter-date-input" id="filter_date_start_' + item.id + '" value="' + formatDateToYMD(startDate) + '" title="Từ ngày" />' +
              '<span class="db-filter-date-sep">➔</span>' +
              '<input type="date" class="db-filter-date-input" id="filter_date_end_' + item.id + '" value="' + formatDateToYMD(endDate) + '" title="Đến ngày" />' +
              '</div>'
            ) : (
              '<div class="db-filter-date-range-wrap">' +
              '<input type="date" class="db-filter-date-input" id="filter_date_val_' + item.id + '" value="' + formatDateToYMD(currentVal !== 'all' ? currentVal : '') + '" title="Chọn ngày" />' +
              '</div>'
            ));
        } else {
          controlsHtml = '<select class="db-filter-val-select" id="filter_val_' + item.id + '" title="Chọn giá trị từ Excel">' +
            '<option value="all">-- Tất cả (' + uniqueVals.length + ') --</option>' +
            uniqueVals.map(v => '<option value="' + escapeHtml(v) + '" ' + (String(v) === String(currentVal) ? 'selected' : '') + '>' + escapeHtml(v) + '</option>').join('') +
            '</select>';
        }

        body.innerHTML = '<div class="db-filter-container">' +
          '<div class="db-filter-header-row">' +
          '<span class="db-filter-label" title="' + escapeHtml(item.name || col) + '">' + escapeHtml(item.name || col) + '</span>' +
          '<div style="display:flex;align-items:center;gap:4px;flex-shrink:0;">' +
          parentBadgeHtml +
          activeBadgeHtml +
          (isFiltering ? ('<button class="db-filter-reset-btn" id="filter_reset_' + item.id + '" title="Xóa bỏ điều kiện lọc này">✕</button>') : '') +
          '</div></div>' +
          '<div class="db-filter-controls-row">' +
          controlsHtml +
          '</div></div>';

        setTimeout(() => {
          const resetBtn = document.getElementById('filter_reset_' + item.id);
          resetBtn?.addEventListener('click', (e) => {
            e.stopPropagation();
            setTabFilterState(null);
          });

          if (fType === 'range') {
            if (isDate) {
              const startInp = document.getElementById('filter_date_start_' + item.id);
              const endInp = document.getElementById('filter_date_end_' + item.id);
              const handleDateRange = () => {
                const sVal = startInp ? startInp.value : '';
                const eVal = endInp ? endInp.value : '';
                if (!sVal && !eVal) setTabFilterState(null);
                else setTabFilterState({ op: 'date_range', start: sVal, end: eVal, isDate: true });
              };
              startInp?.addEventListener('change', handleDateRange);
              endInp?.addEventListener('change', handleDateRange);
            } else {
              const minInp = document.getElementById('filter_range_min_' + item.id);
              const maxInp = document.getElementById('filter_range_max_' + item.id);
              const handleRange = () => {
                const minV = minInp ? minInp.value.trim() : '';
                const maxV = maxInp ? maxInp.value.trim() : '';
                if (minV === '' && maxV === '') setTabFilterState(null);
                else setTabFilterState({ op: 'between', min: minV, max: maxV, start: minV, end: maxV });
              };
              minInp?.addEventListener('change', handleRange);
              maxInp?.addEventListener('change', handleRange);
              minInp?.addEventListener('keydown', (e) => { if (e.key === 'Enter') handleRange(); });
              maxInp?.addEventListener('keydown', (e) => { if (e.key === 'Enter') handleRange(); });
            }
          } else if (fType === 'comparison') {
            const opSel = document.getElementById('filter_op_' + item.id);
            const valEl = document.getElementById('filter_val_' + item.id);
            const applyComp = () => {
              const op = opSel ? opSel.value : '=';
              item.operator = op;
              const val = valEl ? valEl.value.trim() : '';
              if (val === '' || val === 'all') setTabFilterState(null);
              else setTabFilterState({ op, val });
            };
            opSel?.addEventListener('change', applyComp);
            valEl?.addEventListener('change', applyComp);
            valEl?.addEventListener('keydown', (e) => {
              if (e.key === 'Enter') {
                e.preventDefault();
                valEl.blur();
                applyComp();
              }
            });
          } else if (fType === 'multiselect') {
            const btn = document.getElementById('filter_ms_btn_' + item.id);
            const drop = document.getElementById('filter_ms_drop_' + item.id);
            const searchInp = document.getElementById('filter_ms_search_' + item.id);
            const chks = document.querySelectorAll('.filter_ms_chk_' + item.id);
            const btnAll = document.getElementById('filter_ms_all_' + item.id);
            const btnNone = document.getElementById('filter_ms_none_' + item.id);

            // Khôi phục trạng thái mở nếu trước đó dropdown này vừa được tương tác
            if (openDropdownFilterId === item.id && drop) {
              drop.classList.add('open');
              const parentBox = drop.closest('.db-item-box');
              if (parentBox) {
                if (parentBox.dataset.originalZIndex === undefined) {
                  parentBox.dataset.originalZIndex = parentBox.style.zIndex || '';
                }
                parentBox.classList.add('dropdown-open');
                parentBox.style.zIndex = '999999';
              }
            }

            btn?.addEventListener('click', (e) => {
              e.stopPropagation();
              document.querySelectorAll('.db-filter-multiselect-dropdown.open').forEach(d => {
                if (d !== drop) {
                  d.classList.remove('open');
                  const pBox = d.closest('.db-item-box');
                  if (pBox) {
                    pBox.classList.remove('dropdown-open');
                    if (pBox.dataset.originalZIndex !== undefined) {
                      pBox.style.zIndex = pBox.dataset.originalZIndex;
                    }
                  }
                }
              });
              const isOpen = drop?.classList.toggle('open');
              const parentBox = drop?.closest('.db-item-box');
              if (parentBox) {
                if (isOpen) {
                  if (parentBox.dataset.originalZIndex === undefined) {
                    parentBox.dataset.originalZIndex = parentBox.style.zIndex || '';
                  }
                  parentBox.classList.add('dropdown-open');
                  parentBox.style.zIndex = '999999';
                  openDropdownFilterId = item.id;
                } else {
                  parentBox.classList.remove('dropdown-open');
                  if (parentBox.dataset.originalZIndex !== undefined) {
                    parentBox.style.zIndex = parentBox.dataset.originalZIndex;
                  }
                  openDropdownFilterId = null;
                }
              }
              if (drop?.classList.contains('open')) searchInp?.focus();
            });

            drop?.addEventListener('click', (e) => e.stopPropagation());

            searchInp?.addEventListener('input', (e) => {
              const term = e.target.value.toLowerCase().trim();
              document.querySelectorAll('#filter_ms_list_' + item.id + ' .db-filter-multiselect-item').forEach(label => {
                const text = label.textContent.toLowerCase();
                label.style.display = text.includes(term) ? 'flex' : 'none';
              });
            });

            const applyMultiSelect = () => {
              const checked = Array.from(document.querySelectorAll('.filter_ms_chk_' + item.id + ':checked')).map(c => c.value);
              if (checked.length === 0 || checked.length === uniqueVals.length) setTabFilterState(null, true);
              else setTabFilterState({ op: 'in', values: checked, val: checked }, true);
            };

            chks.forEach(c => c.addEventListener('change', applyMultiSelect));
            btnAll?.addEventListener('click', () => { setTabFilterState(null, true); });
            btnNone?.addEventListener('click', () => { setTabFilterState({ op: 'in', values: [], val: [] }, true); });
          } else if (fType === 'text') {
            const opSel = document.getElementById('filter_op_' + item.id);
            const valEl = document.getElementById('filter_val_' + item.id);
            const applyText = () => {
              const op = opSel ? opSel.value : 'contains';
              item.operator = op;
              const val = valEl ? valEl.value.trim() : '';
              if (val === '' || val === 'all') setTabFilterState(null);
              else setTabFilterState({ op, val });
            };
            opSel?.addEventListener('change', applyText);
            valEl?.addEventListener('change', applyText);
            valEl?.addEventListener('keydown', (e) => {
              if (e.key === 'Enter') {
                e.preventDefault();
                valEl.blur();
                applyText();
              }
            });
          } else if (fType === 'date') {
            const opSel = document.getElementById('filter_op_' + item.id);
            const startInp = document.getElementById('filter_date_start_' + item.id);
            const endInp = document.getElementById('filter_date_end_' + item.id);
            const singleInp = document.getElementById('filter_date_val_' + item.id);

            const handleDateChange = () => {
              const op = opSel ? opSel.value : 'date_range';
              if (op === 'date_range') {
                const sVal = startInp ? startInp.value : '';
                const eVal = endInp ? endInp.value : '';
                if (!sVal && !eVal) setTabFilterState(null);
                else setTabFilterState({ op: 'date_range', start: sVal, end: eVal, isDate: true });
              } else {
                const dVal = singleInp ? singleInp.value : '';
                if (!dVal) setTabFilterState(null);
                else setTabFilterState({ op, val: dVal, isDate: true });
              }
            };

            startInp?.addEventListener('change', handleDateChange);
            endInp?.addEventListener('change', handleDateChange);
            singleInp?.addEventListener('change', handleDateChange);
            opSel?.addEventListener('change', (e) => {
              item.operator = e.target.value;
              renderCanvas();
            });
          } else {
            const valSel = document.getElementById('filter_val_' + item.id);
            valSel?.addEventListener('change', (e) => {
              const v = e.target.value;
              if (v === 'all') setTabFilterState(null);
              else setTabFilterState({ op: '=', val: v });
            });
          }
        }, 0);
      }
    }

    function updateCanvasTransform() {
      const canvas = document.getElementById('dashboardCanvas');
      if (!canvas) return;
      const wrapper = canvas.parentElement || document.body;
      const availW = wrapper.clientWidth || window.innerWidth;
      const safePaddingX = 48;
      const scale = Math.min(1.0, Math.max(0.2, (availW - safePaddingX) / 1360));
      const origH = parseInt(canvas.dataset.originalHeight || '800', 10);

      if (scale < 1.0) {
        canvas.style.transform = 'scale(' + scale + ')';
        canvas.style.transformOrigin = 'top center';
        canvas.style.marginBottom = '-' + Math.round((origH * (1 - scale))) + 'px';
      } else {
        canvas.style.transform = 'none';
        canvas.style.transformOrigin = 'top center';
        canvas.style.marginBottom = '0px';
      }
      document.querySelectorAll('.db-chart-canvas').forEach(el => {
        const c = echarts.getInstanceByDom(el);
        if (c) c.resize();
      });
    }

    // Dynamic Excel File Upload Handler
    const viewerUploadBtn = document.getElementById('btnViewerUploadExcel');
    const viewerFileInput = document.getElementById('viewerExcelFileInput');
    viewerUploadBtn?.addEventListener('click', () => viewerFileInput?.click());
    viewerFileInput?.addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (!file) return;

      const isLargeFile = file.size >= PARQUET_THRESHOLD_BYTES;
      const reader = new FileReader();
      reader.onload = (evt) => {
        try {
          const data = new Uint8Array(evt.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const sheetsData = {};

          workbook.SheetNames.forEach(sName => {
            const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sName], { defval: '' });
            if (isLargeFile) {
              const parquetTable = ParquetEngine.encodeToParquet(rows, { fileName: file.name, sheetName: sName });
              sheetsData[sName] = ParquetEngine.createVirtualArray(parquetTable);
            } else {
              sheetsData[sName] = rows;
            }
          });

          FILES[file.name] = sheetsData;
          renderCanvas();
          const note = isLargeFile ? ' (Đã kích hoạt nén Parquet Columnar tiết kiệm RAM)' : '';
          alert('✓ Đã nạp thành công file Excel "' + file.name + '"' + note + '! Toàn bộ biểu đồ đã được cập nhật.');
        } catch (err) {
          alert('Lỗi đọc file Excel: ' + err.message);
        }
      };
      reader.readAsArrayBuffer(file);
    });

    function updateTopFilterBarStickyPosition() {
      const topBar = document.getElementById('dbTopFilterBar');
      if (!topBar) return;
      const canvas = document.getElementById('dashboardCanvas');
      if (!canvas) return;

      const canvasRect = canvas.getBoundingClientRect();
      const effectiveScale = canvas.offsetWidth > 0 ? (canvasRect.width / canvas.offsetWidth) : 1.0;
      let initialLocalY = 24;
      if (topBar.dataset.origTop !== undefined && topBar.dataset.origTop !== '') {
        initialLocalY = parseFloat(topBar.dataset.origTop) || 24;
      } else {
        initialLocalY = topBar.offsetTop || 24;
        topBar.dataset.origTop = initialLocalY;
      }

      const headerEl = document.querySelector('.db-header');
      let headerBottom = 0;
      if (headerEl && headerEl.offsetHeight > 0) {
        const headerRect = headerEl.getBoundingClientRect();
        if (headerRect.top <= 2 && headerRect.bottom > 0) {
          headerBottom = headerRect.bottom;
        }
      }

      const tabsBar = document.getElementById('dbTabsBar');
      let tabsBottom = 0;
      if (tabsBar && tabsBar.offsetHeight > 0) {
        const tabsRect = tabsBar.getBoundingClientRect();
        if (tabsRect.top <= headerBottom + 2 && tabsRect.bottom > headerBottom) {
          tabsBottom = tabsRect.bottom;
        }
      }

      const targetScreenY = Math.max(headerBottom, tabsBottom);
      const naturalScreenY = canvasRect.top + (initialLocalY * effectiveScale);

      if (naturalScreenY <= targetScreenY) {
        const neededTranslateY = (targetScreenY - canvasRect.top) / effectiveScale - initialLocalY;
        const maxTranslateY = Math.max(0, canvas.offsetHeight - topBar.offsetHeight - 20);
        const finalY = Math.min(maxTranslateY, Math.max(0, neededTranslateY));
        topBar.style.transform = 'translate3d(0, ' + finalY + 'px, 0)';
        topBar.classList.add('is-pinned-floating');
      } else {
        topBar.style.transform = '';
        topBar.classList.remove('is-pinned-floating');
      }
    }

    let isTickingSticky = false;
    const handleScrollSticky = () => {
      if (!isTickingSticky) {
        window.requestAnimationFrame(() => {
          updateTopFilterBarStickyPosition();
          isTickingSticky = false;
        });
        isTickingSticky = true;
      }
    };
    window.addEventListener('scroll', handleScrollSticky, { passive: true });
    document.addEventListener('scroll', handleScrollSticky, { passive: true });
    const canvasWrapperEl = document.getElementById('canvasWrapper');
    if (canvasWrapperEl) {
      canvasWrapperEl.addEventListener('scroll', handleScrollSticky, { passive: true });
    }

    window.addEventListener('resize', () => {
      updateCanvasTransform();
      updateTopFilterBarStickyPosition();
    });

    window.addEventListener('click', (e) => {
      if (!e.target.closest('.db-filter-multiselect-dropdown') && !e.target.closest('.db-filter-multiselect-btn')) {
        openDropdownFilterId = null;
        document.querySelectorAll('.db-filter-multiselect-dropdown.open').forEach(d => {
          d.classList.remove('open');
          const pBox = d.closest('.db-item-box') || d.closest('.db-top-filter-item-wrapper');
          if (pBox) {
            pBox.classList.remove('dropdown-open');
            if (pBox.dataset.originalZIndex !== undefined) {
              pBox.style.zIndex = pBox.dataset.originalZIndex;
            }
          }
        });
      }
    });

    // Initial render
    renderAllTabs();
    renderCanvas();
    updateTopFilterBarStickyPosition();
  <\/script>
</body>
</html>`;
  }
  function escapeHtml(str) {
    if (!str) return '';
    return String(str).replace(/[&<>"']/g, m => ({
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;'
    })[m]);
  }

  // ===========================================================================
  // 10. SAVE DASHBOARD HANDLERS (SEPARATE JSON, HTML, AND TRIAD BUNDLE)
  // ===========================================================================
  function prepareDashboardConfigJson() {
    // Luôn đồng bộ toạ độ và kích thước thực tế từ DOM vào State trước khi tạo JSON
    syncItemPositionsFromDOM();

    // 0. Đồng bộ tiêu đề từ ô input trên thanh Header (nếu có)
    const titleInput = document.getElementById('dashboardTitleInput');
    if (titleInput && titleInput.value.trim()) {
      DashboardState.present.title = titleInput.value.trim();
    }

    const usedDataSources = new Set();
    const sheetsNeeded = {};
    (DashboardState.present.items || []).forEach(it => {
      if (it.dataSourceKey) {
        const parts = it.dataSourceKey.split('::');
        const fName = parts[0];
        const sName = parts[1] || '';
        if (fName) {
          usedDataSources.add(fName);
          if (!sheetsNeeded[fName]) sheetsNeeded[fName] = [];
          if (sName && !sheetsNeeded[fName].includes(sName)) sheetsNeeded[fName].push(sName);
        }
      }
    });

    Object.keys(DataStore.files).forEach(fn => {
      if (!fn.includes('Du_Lieu_Mau')) usedDataSources.add(fn);
    });

    // 1. Cập nhật thông tin nhận diện Dashboard
    DashboardState.present.id = DashboardState.present.id || ('dashboard_' + Date.now());
    DashboardState.present.version = '2.0';
    DashboardState.present.schema = 'vnd.antigravity.dashboard.v2';
    DashboardState.present.updatedAt = new Date().toISOString();
    if (!DashboardState.present.createdAt) {
      DashboardState.present.createdAt = new Date().toISOString();
    }

    // 2. Cập nhật cấu hình chủ đề & nền Canvas
    if (!DashboardState.present.theme) DashboardState.present.theme = {};
    DashboardState.present.theme.themeMode = 'light';
    const canvasWrap = document.getElementById('canvasWrapper');
    const matchedBg = (canvasWrap?.className.match(/canvas-bg-([a-zA-Z0-9_-]+)/) || [])[1];
    const cachedBg = (typeof localStorage !== 'undefined' && localStorage.getItem('mb_dashboard_canvas_bg')) || null;
    DashboardState.present.theme.canvasBg = matchedBg || DashboardState.present.theme.canvasBg || cachedBg || 'dots';
    try {
      if (typeof localStorage !== 'undefined' && DashboardState.present.theme.canvasBg) {
        localStorage.setItem('mb_dashboard_canvas_bg', DashboardState.present.theme.canvasBg);
      }
    } catch (_) {}

    // 3. ĐẢM BẢO CẤU HÌNH THANH TAB (VỊ TRÍ 4 HƯỚNG, STYLES, MÀU SẮC, ĐỘ RỘNG...) ĐƯỢC LƯU TRỌN VẸN
    if (typeof TabBarManager !== 'undefined' && TabBarManager.getConfig) {
      DashboardState.present.tabBarConfig = TabBarManager.getConfig();
    }

    // 4. Danh sách các Tab và Tab đang chọn
    if (!Array.isArray(DashboardState.present.tabs) || DashboardState.present.tabs.length === 0) {
      DashboardState.present.tabs = [{ id: 'tab_overview', name: '📊 Trang 1' }];
    }
    if (!DashboardState.present.activeTabId && DashboardState.present.tabs[0]) {
      DashboardState.present.activeTabId = DashboardState.present.tabs[0].id;
    }

    // 5. Trạng thái bộ lọc tương tác hiện tại
    const curActiveTab = DashboardState.present.activeTabId || 'default';
    if (!DashboardState.tabFilters) DashboardState.tabFilters = {};
    DashboardState.tabFilters[curActiveTab] = JSON.parse(JSON.stringify(DashboardState.activeFilters || {}));
    DashboardState.present.tabFilters = JSON.parse(JSON.stringify(DashboardState.tabFilters));
    DashboardState.present.activeFilters = { ...(DashboardState.activeFilters || {}) };

    // 6. Danh mục nguồn dữ liệu cần thiết
    DashboardState.present.dataSourcesNeeded = Array.from(usedDataSources);
    DashboardState.present.sheetsNeeded = sheetsNeeded;

    // 7. Nhúng trọn bộ dữ liệu Excel/CSV (Tự động nén Compact Matrix giảm 65-80% dung lượng)
    DashboardState.present.embeddedData = DataStore.getAllFilesObjectCompact();

    // 8. Đảm bảo toàn bộ thành phần (items) có đầy đủ các thông số chi tiết thiết kế
    DashboardState.present.items = (DashboardState.present.items || []).map(item => {
      const cleanItem = { ...item };
      // Chuẩn hóa toạ độ & kích thước số nguyên
      if (!cleanItem.pos) {
        cleanItem.pos = { x: 24, y: 24, w: 400, h: 300, zIndex: 10 };
      } else {
        cleanItem.pos = {
          x: Math.round(cleanItem.pos.x || 0),
          y: Math.round(cleanItem.pos.y || 0),
          w: Math.round(cleanItem.pos.w || 300),
          h: Math.round(cleanItem.pos.h || 200),
          zIndex: cleanItem.pos.zIndex || 10
        };
      }
      // Gán tabId mặc định nếu thiếu
      if (!cleanItem.tabId && DashboardState.present.activeTabId) {
        cleanItem.tabId = DashboardState.present.activeTabId;
      }
      return cleanItem;
    });

    return JSON.stringify(DashboardState.present, null, 2);
  }

  function saveDashboardConfigJson() {
    DashboardState.pushSnapshot();
    SavedDashboardsCatalog.saveCurrent(false);

    const jsonFileName = 'dashboard_config.json';
    const jsonContent = prepareDashboardConfigJson();

    downloadBlob(new Blob([jsonContent], { type: 'application/json;charset=utf-8' }), jsonFileName);
    DashboardState.isDirty = false;
    DashboardState.updateUI();
    closeSaveConfirmModal();
    showToast(`✓ Đã xuất file JSON cấu hình và dữ liệu nhúng "${jsonFileName}"!`, 'success');
  }

  function exportStandaloneHtml() {
    const htmlFileName = 'index.html';
    const htmlContent = generateStandaloneDashboardHTML();

    downloadBlob(new Blob([htmlContent], { type: 'text/html;charset=utf-8' }), htmlFileName);
    closeSaveConfirmModal();
    showToast(`✓ Đã xuất file HTML báo cáo độc lập "${htmlFileName}"!`, 'success');
  }

  async function exportDashboardZipPackage() {
    if (typeof JSZip === 'undefined') {
      showToast('Thư viện nén tệp (JSZip) đang khởi tạo, vui lòng bấm lại sau 1 giây...', 'warning');
      return false;
    }

    try {
      showToast('Đang đóng gói file ZIP trọn bộ (HTML, JSON & Excel)...', 'info');
      const zip = new JSZip();
      const cleanTitle = (DashboardState.present.title || 'Dashboard').trim().replace(/[/\\?%*:|"<>]/g, '_');

      const htmlContent = generateStandaloneDashboardHTML();
      const jsonContent = prepareDashboardConfigJson();

      // 1. File HTML tái hiện (index.html)
      zip.file('index.html', htmlContent);

      // 2. File JSON cấu hình (dashboard_config.json)
      zip.file('dashboard_config.json', jsonContent);

      // 3. File Excel / CSV nguồn
      let hasExcel = false;
      for (const fileName in DataStore.files) {
        hasExcel = true;
        const fileObj = DataStore.files[fileName];
        if (fileObj.rawBytes) {
          zip.file(fileName, fileObj.rawBytes);
        } else {
          const wb = XLSX.utils.book_new();
          for (const sName in fileObj.sheets) {
            const ws = XLSX.utils.json_to_sheet(fileObj.sheets[sName]);
            XLSX.utils.book_append_sheet(wb, ws, sName);
          }
          const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
          zip.file(fileName, new Uint8Array(wbout));
        }
      }

      if (!hasExcel) {
        const wb = XLSX.utils.book_new();
        const sampleRows = [
          { 'Tháng': 'T1', 'Doanh Thu': 120000000, 'Chi Phí': 80000000, 'Lợi Nhuận': 40000000, 'Nhóm Sản Phẩm': 'Điện tử' },
          { 'Tháng': 'T2', 'Doanh Thu': 150000000, 'Chi Phí': 95000000, 'Lợi Nhuận': 55000000, 'Nhóm Sản Phẩm': 'Gia dụng' },
          { 'Tháng': 'T3', 'Doanh Thu': 180000000, 'Chi Phí': 110000000, 'Lợi Nhuận': 70000000, 'Nhóm Sản Phẩm': 'Điện tử' },
          { 'Tháng': 'T4', 'Doanh Thu': 160000000, 'Chi Phí': 105000000, 'Lợi Nhuận': 55000000, 'Nhóm Sản Phẩm': 'Thời trang' },
          { 'Tháng': 'T5', 'Doanh Thu': 210000000, 'Chi Phí': 130000000, 'Lợi Nhuận': 80000000, 'Nhóm Sản Phẩm': 'Gia dụng' },
          { 'Tháng': 'T6', 'Doanh Thu': 240000000, 'Chi Phí': 145000000, 'Lợi Nhuận': 95000000, 'Nhóm Sản Phẩm': 'Điện tử' }
        ];
        const ws = XLSX.utils.json_to_sheet(sampleRows);
        XLSX.utils.book_append_sheet(wb, ws, 'Du_Lieu_Mau');
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        zip.file('Du_Lieu_Mau_Kinh_Doanh.xlsx', new Uint8Array(wbout));
      }

      // Tạo Blob ZIP nén
      const zipBlob = await zip.generateAsync({
        type: 'blob',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 }
      });

      const zipFileName = `dashboard_${cleanTitle}_package.zip`;
      downloadBlob(zipBlob, zipFileName);

      DashboardState.isDirty = false;
      DashboardState.updateUI();
      closeSaveConfirmModal();
      showToast(`✓ Đã tải về gói ZIP trọn bộ "${zipFileName}" (Chứa đúng 3 file: index.html, dashboard_config.json, Excel)!`, 'success');
      return true;
    } catch (err) {
      console.error('Lỗi đóng gói ZIP:', err);
      showToast('Lỗi tạo file ZIP: ' + err.message, 'error');
      return false;
    }
  }

  // ===========================================================================
  // 10. ACTIVE DIRECTORY HANDLE PERSISTENCE & DIRECT AUTO-SAVING
  // ===========================================================================
  let activeDashboardFolderHandle = null;
  let activeDashboardFolderName = '';

  function saveActiveWorkspaceHandle(dirHandle, folderName) {
    if (!dirHandle) return;
    try {
      const req = indexedDB.open('mb_dashboard_hub_db', 3);
      req.onupgradeneeded = (e) => {
        const db = e.target.result;
        if (!db.objectStoreNames.contains('pending_import')) {
          db.createObjectStore('pending_import', { keyPath: 'id' });
        }
        if (!db.objectStoreNames.contains('active_workspace')) {
          db.createObjectStore('active_workspace', { keyPath: 'id' });
        }
      };
      req.onsuccess = (e) => {
        try {
          const db = e.target.result;
          if (!db.objectStoreNames.contains('active_workspace')) return;
          const tx = db.transaction('active_workspace', 'readwrite');
          const store = tx.objectStore('active_workspace');
          store.put({ id: 'current_workspace', dirHandle, folderName: folderName || dirHandle.name, timestamp: Date.now() });
        } catch (_) {}
      };
    } catch (_) {}
  }

  function getActiveWorkspaceHandle() {
    return new Promise((resolve) => {
      try {
        const req = indexedDB.open('mb_dashboard_hub_db', 3);
        req.onupgradeneeded = (e) => {
          const db = e.target.result;
          if (!db.objectStoreNames.contains('pending_import')) {
            db.createObjectStore('pending_import', { keyPath: 'id' });
          }
          if (!db.objectStoreNames.contains('active_workspace')) {
            db.createObjectStore('active_workspace', { keyPath: 'id' });
          }
        };
        req.onsuccess = (e) => {
          try {
            const db = e.target.result;
            if (!db.objectStoreNames.contains('active_workspace')) {
              resolve(null);
              return;
            }
            const tx = db.transaction('active_workspace', 'readonly');
            const store = tx.objectStore('active_workspace');
            const getReq = store.get('current_workspace');
            getReq.onsuccess = () => resolve(getReq.result || null);
            getReq.onerror = () => resolve(null);
          } catch (_) {
            resolve(null);
          }
        };
        req.onerror = () => resolve(null);
      } catch (_) {
        resolve(null);
      }
      setTimeout(() => resolve(null), 600);
    });
  }

  function clearActiveWorkspaceHandle() {
    try {
      const req = indexedDB.open('mb_dashboard_hub_db', 3);
      req.onsuccess = (e) => {
        try {
          const db = e.target.result;
          if (!db.objectStoreNames.contains('active_workspace')) return;
          const tx = db.transaction('active_workspace', 'readwrite');
          tx.objectStore('active_workspace').delete('current_workspace');
        } catch (_) {}
      };
    } catch (_) {}
  }

  let isAutoSavingToDisk = false;
  async function autoSaveToActiveDirectory() {
    if (!activeDashboardFolderHandle || isAutoSavingToDisk) return;
    try {
      isAutoSavingToDisk = true;
      let hasPerm = false;
      if (activeDashboardFolderHandle.queryPermission) {
        const perm = await activeDashboardFolderHandle.queryPermission({ mode: 'readwrite' });
        hasPerm = (perm === 'granted');
      } else {
        hasPerm = true;
      }
      if (!hasPerm) return;

      syncItemPositionsFromDOM();
      const jsonContent = prepareDashboardConfigJson();
      const htmlContent = generateStandaloneDashboardHTML();

      // 1. Cập nhật file JSON cấu hình 'dashboard_config.json'
      const stdJsonHandle = await activeDashboardFolderHandle.getFileHandle('dashboard_config.json', { create: true });
      const stdJsonWritable = await stdJsonHandle.createWritable();
      await stdJsonWritable.write(jsonContent);
      await stdJsonWritable.close();

      // 2. Cập nhật file HTML tái hiện mới nhất 'index.html'
      const htmlHandle = await activeDashboardFolderHandle.getFileHandle('index.html', { create: true });
      const htmlWritable = await htmlHandle.createWritable();
      await htmlWritable.write(htmlContent);
      await htmlWritable.close();

      SavedDashboardsCatalog.saveCurrent(false);
      DashboardState.isDirty = false;

      const dirtyBadge = document.getElementById('dirtyBadge');
      if (dirtyBadge) {
        dirtyBadge.textContent = `✓ Đã tự động lưu vào ${activeDashboardFolderName || 'thư mục'}`;
        dirtyBadge.className = 'db-dirty-badge saved';
      }
    } catch (err) {
      console.warn('Lỗi tự động ghi file vào thư mục:', err);
    } finally {
      isAutoSavingToDisk = false;
    }
  }

  async function writePackageToDirectoryHandle(targetDirHandle, folderDisplayName) {
    const titleInput = document.getElementById('dashboardTitleInput');
    if (titleInput && titleInput.value.trim()) {
      DashboardState.present.title = titleInput.value.trim();
    }
    const htmlFileName = 'index.html';
    const htmlContent = generateStandaloneDashboardHTML();

    const jsonFileName = 'dashboard_config.json';
    const jsonContent = prepareDashboardConfigJson();

    // Chuẩn bị các file Excel nguồn
    const excelFiles = [];
    for (const fileName in DataStore.files) {
      const fileObj = DataStore.files[fileName];
      if (fileObj.rawBytes) {
        excelFiles.push({ name: fileName, data: fileObj.rawBytes });
      } else {
        const wb = XLSX.utils.book_new();
        for (const sName in fileObj.sheets) {
          const ws = XLSX.utils.json_to_sheet(fileObj.sheets[sName]);
          XLSX.utils.book_append_sheet(wb, ws, sName);
        }
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        excelFiles.push({ name: fileName, data: new Uint8Array(wbout) });
      }
    }

    if (excelFiles.length === 0) {
      const wb = XLSX.utils.book_new();
      const sampleRows = [
        { 'Tháng': 'T1', 'Doanh Thu': 120000000, 'Chi Phí': 80000000, 'Lợi Nhuận': 40000000, 'Nhóm Sản Phẩm': 'Điện tử' },
        { 'Tháng': 'T2', 'Doanh Thu': 150000000, 'Chi Phí': 95000000, 'Lợi Nhuận': 55000000, 'Nhóm Sản Phẩm': 'Gia dụng' },
        { 'Tháng': 'T3', 'Doanh Thu': 180000000, 'Chi Phí': 110000000, 'Lợi Nhuận': 70000000, 'Nhóm Sản Phẩm': 'Điện tử' },
        { 'Tháng': 'T4', 'Doanh Thu': 160000000, 'Chi Phí': 105000000, 'Lợi Nhuận': 55000000, 'Nhóm Sản Phẩm': 'Thời trang' },
        { 'Tháng': 'T5', 'Doanh Thu': 210000000, 'Chi Phí': 130000000, 'Lợi Nhuận': 80000000, 'Nhóm Sản Phẩm': 'Gia dụng' },
        { 'Tháng': 'T6', 'Doanh Thu': 240000000, 'Chi Phí': 145000000, 'Lợi Nhuận': 95000000, 'Nhóm Sản Phẩm': 'Điện tử' }
      ];
      const ws = XLSX.utils.json_to_sheet(sampleRows);
      XLSX.utils.book_append_sheet(wb, ws, 'Du_Lieu_Mau');
      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      excelFiles.push({ name: 'Du_Lieu_Mau_Kinh_Doanh.xlsx', data: new Uint8Array(wbout) });
    }

    SavedDashboardsCatalog.saveCurrent(false);

    // 1. Ghi đúng file HTML tái hiện 'index.html'
    const htmlHandle = await targetDirHandle.getFileHandle(htmlFileName, { create: true });
    const htmlWritable = await htmlHandle.createWritable();
    await htmlWritable.write(htmlContent);
    await htmlWritable.close();

    // 2. Ghi đúng file JSON cấu hình 'dashboard_config.json'
    const stdJsonHandle = await targetDirHandle.getFileHandle(jsonFileName, { create: true });
    const stdJsonWritable = await stdJsonHandle.createWritable();
    await stdJsonWritable.write(jsonContent);
    await stdJsonWritable.close();

    // 3. Ghi file Excel dữ liệu nguồn
    for (const ef of excelFiles) {
      const efHandle = await targetDirHandle.getFileHandle(ef.name, { create: true });
      const efWritable = await efHandle.createWritable();
      await efWritable.write(ef.data);
      await efWritable.close();
    }

    // Tự động dọn dẹp các file thừa từ phiên bản cũ (nếu có) để chỉ còn đúng 3 file trong thư mục
    try {
      const titleSlug = (DashboardState.present.title || 'dashboard').trim().toLowerCase().replace(/[^a-z0-9]+/g, '_');
      const oldSlugHtml = `dashboard_${titleSlug}.html`;
      const oldSlugJson = `dashboard_${titleSlug}_config.json`;
      if (oldSlugHtml !== htmlFileName) {
        await targetDirHandle.removeEntry(oldSlugHtml).catch(() => {});
      }
      if (oldSlugJson !== jsonFileName) {
        await targetDirHandle.removeEntry(oldSlugJson).catch(() => {});
      }
    } catch (_) {}

    DashboardState.isDirty = false;
    DashboardState.updateUI();
    closeSaveConfirmModal();
    const nameToShow = folderDisplayName || targetDirHandle.name || 'Dashboard';
    showToast(`✓ Đã lưu thành công 3 tệp (HTML, JSON, Excel) vào thư mục "${nameToShow}"!`, 'success');
    return true;
  }

  async function saveAndExportDashboardPackage(forcePickNewFolder = false) {
    // 0. Luôn luôn đồng bộ tiêu đề mới nhất từ ô input trước khi tiến hành lưu
    const titleInput = document.getElementById('dashboardTitleInput');
    if (titleInput && titleInput.value.trim()) {
      DashboardState.present.title = titleInput.value.trim();
    }

    // 1. Nếu đã xác định thư mục con rồi (đã lưu lần đầu hoặc đã mở từ thư mục), từ lần 2 trở đi chỉ cần bấm Lưu là tự động lưu ngay mà không cần chọn lại
    if (!forcePickNewFolder && activeDashboardFolderHandle) {
      try {
        let hasPerm = false;
        if (activeDashboardFolderHandle.queryPermission) {
          let perm = await activeDashboardFolderHandle.queryPermission({ mode: 'readwrite' });
          if (perm === 'granted') {
            hasPerm = true;
          } else if (activeDashboardFolderHandle.requestPermission) {
            perm = await activeDashboardFolderHandle.requestPermission({ mode: 'readwrite' });
            hasPerm = (perm === 'granted');
          }
        } else {
          hasPerm = true;
        }

        if (hasPerm) {
          const ok = await writePackageToDirectoryHandle(activeDashboardFolderHandle, activeDashboardFolderName);
          if (ok) {
            saveActiveWorkspaceHandle(activeDashboardFolderHandle, activeDashboardFolderName);
            saveLastEditedSessionToLocalStorage();
            return true;
          }
        }
      } catch (err) {
        console.warn('Lỗi ghi vào thư mục đã lưu trước đó, sẽ yêu cầu chọn lại thư mục cha:', err);
        activeDashboardFolderHandle = null;
      }
    }

    // 2. Lần đầu lưu: Mở hộp thoại để người dùng chọn thư mục cha trên máy tính
    if (window.showDirectoryPicker) {
      try {
        showToast('Vui lòng chọn thư mục cha trên máy tính để lưu Dashboard...', 'info');
        const parentDirHandle = await window.showDirectoryPicker({
          mode: 'readwrite'
        });

        // Lấy tên Dashboard trong cấu hình để tự động tạo thư mục con
        const titleInput = document.getElementById('dashboardTitleInput');
        if (titleInput && titleInput.value.trim()) {
          DashboardState.present.title = titleInput.value.trim();
        }
        const rawTitle = (DashboardState.present.title || 'Dashboard').trim();
        const subfolderName = rawTitle.replace(/[/\\?%*:|"<>]/g, '_').trim() || 'Dashboard';

        // Tự động tạo thư mục con theo tên cấu hình Dashboard bên trong thư mục cha
        const childDirHandle = await parentDirHandle.getDirectoryHandle(subfolderName, { create: true });

        // Ghi chính xác 3 file (HTML tái hiện, JSON cấu hình, Excel dữ liệu) vào thư mục con này
        const ok = await writePackageToDirectoryHandle(childDirHandle, subfolderName);
        if (ok) {
          activeDashboardFolderHandle = childDirHandle;
          activeDashboardFolderName = subfolderName;
          saveActiveWorkspaceHandle(childDirHandle, subfolderName);
          saveLastEditedSessionToLocalStorage();
          return true;
        }
      } catch (err) {
        if (err.name === 'AbortError') {
          showToast('Đã hủy chọn thư mục.', 'info');
          return false;
        }
        console.warn('Lỗi Directory Picker:', err);
        showToast('Không thể ghi thư mục trực tiếp, chuyển sang chế độ tải về gói ZIP...', 'warning');
      }
    }

    // Fallback nếu trình duyệt không hỗ trợ File System Access API
    return await exportDashboardZipPackage();
  }

  const executeSaveDashboardFiles = saveAndExportDashboardPackage;

  function downloadBlob(blob, fileName) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 200);
  }

  function openSaveConfirmModal(onAfterSaveAction = null) {
    const modal = document.getElementById('saveConfirmModal');
    if (!modal) return;
    document.getElementById('confirmHtmlName').textContent = 'index.html';
    document.getElementById('confirmJsonName').textContent = 'dashboard_config.json';

    const sources = Object.keys(DataStore.files);
    document.getElementById('confirmExcelName').textContent = sources.length > 0 ? sources.join(', ') : 'Du_Lieu_Mau_Kinh_Doanh.xlsx';

    modal._onAfterSave = onAfterSaveAction;
    modal.classList.add('open');
  }

  function closeSaveConfirmModal() {
    const modal = document.getElementById('saveConfirmModal');
    if (modal) modal.classList.remove('open');
  }

  function openExitConfirmModal() {
    if (!DashboardState.isDirty) {
      window.location.href = 'index.html';
      return;
    }
    openSaveConfirmModal(() => {
      window.location.href = 'index.html';
    });
  }

  async function createNewDashboard(force = false) {
    if (!force && DashboardState.isDirty) {
      const proceed = await AppDialog.confirm('Dashboard hiện tại có thay đổi chưa lưu. Bạn có chắc muốn tạo mới và hủy các thay đổi này không?', 'Tạo Dashboard Mới', { isDanger: true, confirmText: 'Tạo Mới & Hủy Thay Đổi' });
      if (!proceed) {
        return;
      }
    }

    DashboardState.pushSnapshot();
    const newId = 'dashboard_' + Date.now();
    DashboardState.present = {
      id: newId,
      title: 'Báo Cáo Dashboard Mới',
      description: 'Dashboard BI tùy biến bằng JavaScript ECharts và nguồn dữ liệu Excel',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      theme: { themeMode: 'light', canvasBg: (typeof localStorage !== 'undefined' && localStorage.getItem('mb_dashboard_canvas_bg')) || 'dots' },
      tabs: [
        { id: 'tab_page_1', name: '📊 Trang 1' }
      ],
      activeTabId: 'tab_page_1',
      items: []
    };
    DashboardState.isDirty = false;
    DashboardState.selectedItemId = null;
    DashboardState.activeFilters = {};
    activeDashboardFolderHandle = null;
    activeDashboardFolderName = '';
    clearActiveWorkspaceHandle();
    DashboardState.updateUI();

    const titleInput = document.getElementById('dashboardTitleInput');
    if (titleInput) titleInput.value = DashboardState.present.title;

    applyCanvasBackgroundUI(DashboardState.present.theme.canvasBg);
    renderAllTabs();
    renderCanvas();
    showToast('Đã khởi tạo Dashboard mới! Hãy thêm biểu đồ hoặc widget để bắt đầu chỉnh sửa.', 'success');
  }

  async function parseExcelFileAsync(file) {
    const isCsv = file.name.toLowerCase().endsWith('.csv');
    const isLargeFile = file.size >= PARQUET_THRESHOLD_BYTES;
    const sizeStr = file.size > 1048576 
      ? (file.size / 1048576).toFixed(1) + ' MB' 
      : Math.max(1, Math.round(file.size / 1024)) + ' KB';

    if (isCsv) {
      // Dùng Fast Streaming CSV Parser: Đọc trực tiếp thành Columnar Table không qua SheetJS
      const parquetTable = await FastCsvParser.parse(file);
      const sName = file.name.replace(/\.[^/.]+$/, '');
      const sheetsData = {
        [sName]: ParquetEngine.createVirtualArray(parquetTable)
      };

      return {
        name: file.name,
        sizeStr: isLargeFile ? `${sizeStr} (📦 Parquet)` : sizeStr,
        uploadedAt: new Date().toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit' }),
        sheets: sheetsData,
        isParquet: isLargeFile,
        rawBytes: null
      };
    } else {
      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(new Uint8Array(arrayBuffer), {
        type: 'array',
        dense: true,
        cellFormula: false,
        cellHTML: false,
        cellText: false
      });
      const sheetsData = {};
      const sheetNames = workbook.SheetNames || [];

      sheetNames.forEach(sName => {
        const worksheet = workbook.Sheets[sName];
        const rows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
        delete workbook.Sheets[sName]; // Giải phóng ngay sheet khỏi bộ nhớ
        if (isLargeFile && rows.length > 500) {
          const parquetTable = ParquetEngine.encodeToParquet(rows, { fileName: file.name, sheetName: sName });
          sheetsData[sName] = ParquetEngine.createVirtualArray(parquetTable);
        } else {
          sheetsData[sName] = rows;
        }
      });

      return {
        name: file.name,
        sizeStr: isLargeFile ? `${sizeStr} (📦 Parquet)` : sizeStr,
        uploadedAt: new Date().toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit' }),
        sheets: sheetsData,
        isParquet: isLargeFile,
        rawBytes: null
      };
    }
  }

  // ===========================================================================
  // 10.8 UNIFIED DASHBOARD IMPORTER & STATE REHYDRATION (JSON, HTML, ZIP, FOLDER)
  // ===========================================================================
  function extractDashboardFromHtml(htmlText) {
    if (!htmlText || typeof htmlText !== 'string') return null;
    let config = null;
    let embeddedData = null;

    // 1. Cắt chuỗi theo marker chính xác
    const cfgMarker = 'const CONFIG =';
    const filesMarker = 'let FILES =';
    const tabMarker = 'let activeTabId =';

    const cfgIdx = htmlText.indexOf(cfgMarker);
    const filesIdx = htmlText.indexOf(filesMarker);

    if (cfgIdx !== -1 && filesIdx !== -1 && filesIdx > cfgIdx) {
      let cfgStr = htmlText.substring(cfgIdx + cfgMarker.length, filesIdx).trim();
      if (cfgStr.endsWith(';')) cfgStr = cfgStr.slice(0, -1).trim();
      try {
        config = JSON.parse(cfgStr);
      } catch (err) {
        console.warn('Lỗi parse CONFIG từ HTML:', err);
      }

      let filesStr = '';
      const tabIdx = htmlText.indexOf(tabMarker, filesIdx);
      if (tabIdx !== -1 && tabIdx > filesIdx) {
        filesStr = htmlText.substring(filesIdx + filesMarker.length, tabIdx).trim();
      } else {
        const nextScript = htmlText.indexOf('</script>', filesIdx);
        if (nextScript !== -1) {
          filesStr = htmlText.substring(filesIdx + filesMarker.length, nextScript).trim();
        }
      }
      if (filesStr.endsWith(';')) filesStr = filesStr.slice(0, -1).trim();
      if (filesStr) {
        try {
          embeddedData = JSON.parse(filesStr);
        } catch (err) {
          console.warn('Lỗi parse FILES từ HTML:', err);
        }
      }
    }

    // 2. Regex fallback nếu cấu trúc cắt chuỗi bị xê dịch
    if (!config) {
      const mCfg = htmlText.match(/const\s+CONFIG\s*=\s*(\{[\s\S]*?\})\s*;\s*(?:let|const|var)\s+FILES/);
      if (mCfg && mCfg[1]) {
        try { config = JSON.parse(mCfg[1]); } catch (_) {}
      }
    }
    if (!embeddedData && config) {
      const mFiles = htmlText.match(/(?:let|const|var)\s+FILES\s*=\s*(\{[\s\S]*?\})\s*;\s*(?:let|const|var)\s+activeTabId/);
      if (mFiles && mFiles[1]) {
        try { embeddedData = JSON.parse(mFiles[1]); } catch (_) {}
      }
    }

    if (config) {
      if (embeddedData && typeof embeddedData === 'object' && Object.keys(embeddedData).length > 0) {
        config.embeddedData = embeddedData;
      }
      return config;
    }
    return null;
  }

  function loadDashboardFromConfig(config, options = {}) {
    if (!config || typeof config !== 'object' || (!config.items && !config.tabs && !config.title)) {
      showToast('Cấu trúc file Dashboard không hợp lệ hoặc bị hỏng!', 'error');
      return false;
    }

    // 1. Tái sinh toàn bộ dữ liệu nguồn từ embeddedData (nếu DataStore chưa có từ file Excel thực tế trong folder)
    if (config.embeddedData && typeof config.embeddedData === 'object') {
      const embedded = config.embeddedData;
      const fileNames = Object.keys(embedded);
      if (fileNames.length > 0) {
        for (const fName in embedded) {
          // Chỉ nạp từ embeddedData nếu file chưa được nạp từ thư mục (ưu tiên file Excel thật từ thư mục)
          if (!DataStore.files[fName] || !DataStore.files[fName].sheets) {
            const rawSheets = embedded[fName] || {};
            const uncompactedSheets = {};
            for (const sName in rawSheets) {
              uncompactedSheets[sName] = DataStore.uncompactRows(rawSheets[sName]);
            }
            DataStore.files[fName] = {
              name: fName,
              sizeStr: 'Nhúng sẵn trong file JSON',
              uploadedAt: new Date().toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit' }),
              sheets: uncompactedSheets,
              rawBytes: null
            };
          }
        }
        // Xóa file dữ liệu mẫu khi đã có dữ liệu thật
        if (DataStore.files['Du_Lieu_Mau_Kinh_Doanh.xlsx'] && Object.keys(DataStore.files).length > 1) {
          delete DataStore.files['Du_Lieu_Mau_Kinh_Doanh.xlsx'];
        }
      }
    }

    // 1.5. Đảm bảo cấu trúc Tabs hợp lệ
    if (!Array.isArray(config.tabs) || config.tabs.length === 0) {
      config.tabs = [{ id: 'tab_overview', name: '📊 Trang 1' }];
      config.activeTabId = 'tab_overview';
    } else if (!config.activeTabId || !config.tabs.some(t => t.id === config.activeTabId)) {
      config.activeTabId = config.tabs[0].id;
    }

    const validTabIds = new Set(config.tabs.map(t => t.id));
    const fallbackTabId = config.tabs[0].id;

    // 2. Bảo đảm toạ độ an toàn, tabId và tự động đối chiếu thông minh dataSourceKey
    const availableFiles = Object.keys(DataStore.files);
    if (Array.isArray(config.items)) {
      config.items.forEach(item => {
        if (item.pos) {
          const parsedX = typeof item.pos.x === 'number' ? item.pos.x : parseInt(item.pos.x, 10);
          const parsedY = typeof item.pos.y === 'number' ? item.pos.y : parseInt(item.pos.y, 10);
          const parsedW = typeof item.pos.w === 'number' ? item.pos.w : parseInt(item.pos.w, 10);
          const parsedH = typeof item.pos.h === 'number' ? item.pos.h : parseInt(item.pos.h, 10);
          item.pos.x = (!isNaN(parsedX) && parsedX >= 24) ? parsedX : 24;
          item.pos.y = (!isNaN(parsedY) && parsedY >= 24) ? parsedY : 24;
          item.pos.w = (!isNaN(parsedW) && parsedW >= 50) ? parsedW : 400;
          item.pos.h = (!isNaN(parsedH) && parsedH >= 30) ? parsedH : 300;
          if (item.pos.x + item.pos.w > 1336) {
            item.pos.x = Math.max(24, 1336 - item.pos.w);
          }
        } else {
          item.pos = { x: 24, y: 24, w: 400, h: 300, zIndex: 10 };
        }

        // Đảm bảo widget luôn thuộc một tab hợp lệ
        if (!item.tabId || !validTabIds.has(item.tabId)) {
          item.tabId = fallbackTabId;
        }

        if (item.dataSourceKey && availableFiles.length > 0) {
          const parts = item.dataSourceKey.split('::');
          const currentFileName = parts[0];
          const currentSheetName = parts[1] || '';

          if (!DataStore.files[currentFileName]) {
            if (availableFiles.length === 1) {
              const onlyFile = availableFiles[0];
              const targetSheet = (DataStore.files[onlyFile].sheets && DataStore.files[onlyFile].sheets[currentSheetName])
                ? currentSheetName
                : Object.keys(DataStore.files[onlyFile].sheets || {})[0] || '';
              item.dataSourceKey = `${onlyFile}::${targetSheet}`;
            } else {
              const matched = availableFiles.find(fn =>
                fn.toLowerCase() === currentFileName.toLowerCase() ||
                fn.toLowerCase().includes(currentFileName.toLowerCase()) ||
                currentFileName.toLowerCase().includes(fn.toLowerCase())
              );
              if (matched) {
                const targetSheet = (DataStore.files[matched].sheets && DataStore.files[matched].sheets[currentSheetName])
                  ? currentSheetName
                  : Object.keys(DataStore.files[matched].sheets || {})[0] || '';
                item.dataSourceKey = `${matched}::${targetSheet}`;
              }
            }
          }
        }
      });
    }

    // 3. Cập nhật DashboardState
    DashboardState.pushSnapshot();
    DashboardState.present = config;
    DashboardState.isDirty = false;
    DashboardState.selectedItemId = null;
    DashboardState.tabFilters = (config.tabFilters && typeof config.tabFilters === 'object')
      ? JSON.parse(JSON.stringify(config.tabFilters))
      : {};
    DashboardState.activeFilters = (config.activeFilters && typeof config.activeFilters === 'object')
      ? JSON.parse(JSON.stringify(config.activeFilters))
      : {};
    if (config.activeTabId && DashboardState.tabFilters[config.activeTabId]) {
      DashboardState.activeFilters = JSON.parse(JSON.stringify(DashboardState.tabFilters[config.activeTabId]));
    }
    DashboardState.updateUI();

    // 4. Đồng bộ tiêu đề trên thanh Header
    const titleInput = document.getElementById('dashboardTitleInput');
    if (titleInput && config.title) {
      titleInput.value = config.title;
    }

    // 5. Cập nhật Background Canvas nếu có
    const bgToApply = config.theme?.canvasBg || config.canvasBg || (typeof localStorage !== 'undefined' && localStorage.getItem('mb_dashboard_canvas_bg')) || 'dots';
    if (typeof applyCanvasBackgroundUI === 'function') {
      applyCanvasBackgroundUI(bgToApply);
    }

    // 6. Đồng bộ toàn bộ giao diện, Tab bar, Canvas và Quản lý dữ liệu
    if (config.tabBarConfig && typeof TabBarManager !== 'undefined') {
      TabBarManager.saveConfig(config.tabBarConfig);
      TabBarManager.apply(config.tabBarConfig);
    } else if (typeof TabBarManager !== 'undefined' && TabBarManager.apply) {
      TabBarManager.apply();
    }
    if (typeof FigmaPropertiesSidebar !== 'undefined' && FigmaPropertiesSidebar.syncTabUI) {
      FigmaPropertiesSidebar.syncTabUI();
    }
    renderAllTabs();
    renderCanvas();
    updateDataSourcesSummary();
    updateExcelButtonState();
    SavedDashboardsCatalog.saveCurrent(false);
    saveLastEditedSessionToLocalStorage();

    // Tự động cuộn lên đầu vùng làm việc
    const wrapper = document.getElementById('canvasWrapper');
    if (wrapper) wrapper.scrollTop = 0;

    // 7. Thông báo trạng thái người dùng
    const neededFiles = config.dataSourcesNeeded || [];
    const missingFiles = neededFiles.filter(fn => !DataStore.files[fn]);
    if (missingFiles.length > 0) {
      showToast(options.warnMsg || `✓ Đã nạp Dashboard! Lưu ý: Chưa có file dữ liệu "${missingFiles.join(', ')}", bạn có thể bấm "📤 Nạp File Excel" để bổ sung.`, 'info');
    } else {
      showToast(options.successMsg || '✓ Đã tái hiện 100% giao diện, bố cục và dữ liệu Dashboard!', 'success');
    }
    return true;
  }

  async function restoreDashboardFromFolderFiles(filesArray) {
    if (!filesArray || filesArray.length === 0) return false;

    let configJson = null;
    let configFileName = '';
    const excelFiles = [];
    const htmlFiles = [];
    const candidateJsonFiles = [];

    for (const file of filesArray) {
      const nameLower = file.name.toLowerCase();
      if (nameLower.endsWith('.json')) {
        let score = 0;
        if (nameLower === 'dashboard_config.json') score = 100;
        else if (nameLower.includes('dashboard') && nameLower.includes('config')) score = 80;
        else if (nameLower.includes('config')) score = 60;
        else if (nameLower.includes('dashboard')) score = 50;
        else score = 10;
        candidateJsonFiles.push({ file, score });
      } else if (nameLower.endsWith('.xlsx') || nameLower.endsWith('.xls') || nameLower.endsWith('.csv')) {
        excelFiles.push(file);
      } else if (nameLower.endsWith('.html') || nameLower.endsWith('.htm')) {
        htmlFiles.push(file);
      }
    }

    // Đọc tất cả các file JSON ứng viên và tìm file cấu hình mới nhất (theo updatedAt hoặc lastModified)
    const validJsonCandidates = [];
    for (const candidate of candidateJsonFiles) {
      try {
        const text = await candidate.file.text();
        const parsed = JSON.parse(text);
        if (parsed && (parsed.items || parsed.tabs || parsed.title)) {
          const modTime = (parsed.updatedAt ? new Date(parsed.updatedAt).getTime() : 0) || (candidate.file.lastModified || 0);
          validJsonCandidates.push({
            file: candidate.file,
            config: parsed,
            modTime: modTime,
            score: candidate.score
          });
        }
      } catch (err) {
        console.warn('Lỗi đọc file JSON cấu hình:', candidate.file.name, err);
      }
    }

    // Sắp xếp: Ưu tiên file được lưu gần nhất (modTime cao nhất)
    validJsonCandidates.sort((a, b) => {
      if (b.modTime !== a.modTime) return b.modTime - a.modTime;
      return b.score - a.score;
    });

    if (validJsonCandidates.length > 0) {
      configJson = validJsonCandidates[0].config;
      configFileName = validJsonCandidates[0].file.name;
    }

    // Nếu không có file .json, thử trích xuất cấu hình từ file .html (ví dụ index.html đã xuất trước đó)
    if (!configJson && htmlFiles.length > 0) {
      htmlFiles.sort((a, b) => (b.lastModified || 0) - (a.lastModified || 0));
      for (const hf of htmlFiles) {
        try {
          const htmlText = await hf.text();
          const extracted = extractDashboardFromHtml(htmlText);
          if (extracted) {
            configJson = extracted;
            configFileName = hf.name;
            break;
          }
        } catch (_) {}
      }
    }

    if (!configJson) {
      await AppDialog.alert(
        'Không tìm thấy file cấu hình Dashboard hợp lệ trong thư mục được chọn.\n\nThư mục cần chứa:\n1. File cấu hình (dashboard_config.json hoặc *_config.json) HOẶC file HTML báo cáo (index.html)\n2. File Excel nguồn dữ liệu (*.xlsx / *.csv)',
        'Không Tìm Thấy Cấu Hình Dashboard'
      );
      return false;
    }

    // Đọc và nạp đầy đủ các file Excel bất đồng bộ trước khi vẽ lại Dashboard
    let loadedExcelCount = 0;
    if (excelFiles.length > 0) {
      try {
        DataStore.files = {}; // Làm mới danh sách file để không bị lẫn dữ liệu từ dashboard cũ
        const parsedFiles = await Promise.all(excelFiles.map(f => parseExcelFileAsync(f)));
        parsedFiles.forEach(fileObj => {
          DataStore.files[fileObj.name] = fileObj;
          loadedExcelCount++;
        });

        if (DataStore.files['Du_Lieu_Mau_Kinh_Doanh.xlsx'] && Object.keys(DataStore.files).length > 1) {
          delete DataStore.files['Du_Lieu_Mau_Kinh_Doanh.xlsx'];
        }
      } catch (excelErr) {
        console.error('Lỗi phân tích file Excel từ thư mục:', excelErr);
        showToast('Có lỗi khi đọc file Excel: ' + excelErr.message, 'error');
      }
    }

    const ok = loadDashboardFromConfig(configJson, {
      successMsg: loadedExcelCount > 0
        ? `✓ Đã tái hiện 100% Dashboard từ thư mục! ("${configFileName}", ${loadedExcelCount} file Excel)`
        : `✓ Đã nạp thành công cấu hình Dashboard từ "${configFileName}"!`
    });

    if (ok) {
      saveLastEditedSessionToLocalStorage();
    }
    return ok;
  }

  async function importDashboardZipFile(file) {
    if (typeof JSZip === 'undefined') {
      showToast('Thư viện giải nén ZIP (JSZip) chưa sẵn sàng.', 'error');
      return false;
    }

    try {
      showToast('Đang giải nén gói ZIP...', 'info');
      const zip = await JSZip.loadAsync(file);
      let configFileContent = null;
      let htmlFileContent = null;
      const excelEntries = [];

      for (const relativePath in zip.files) {
        const zipEntry = zip.files[relativePath];
        if (zipEntry.dir) continue;
        const nameLower = zipEntry.name.toLowerCase();

        if (nameLower.endsWith('.json')) {
          const str = await zipEntry.async('string');
          if (!configFileContent || nameLower === 'dashboard_config.json' || nameLower.includes('config')) {
            configFileContent = str;
          }
        } else if (nameLower.endsWith('.xlsx') || nameLower.endsWith('.xls') || nameLower.endsWith('.csv')) {
          const arrayBuffer = await zipEntry.async('arraybuffer');
          excelEntries.push({ name: zipEntry.name.split('/').pop(), data: new Uint8Array(arrayBuffer) });
        } else if (nameLower.endsWith('.html') || nameLower.endsWith('.htm')) {
          htmlFileContent = await zipEntry.async('string');
        }
      }

      // 1. Phân tích các file Excel trong ZIP nếu có
      if (excelEntries.length > 0) {
        DataStore.files = {};
      }
      let loadedExcelCount = 0;
      for (const ef of excelEntries) {
        try {
          const isCsv = ef.name.toLowerCase().endsWith('.csv');
          let workbook;
          if (isCsv) {
            const decoder = new TextDecoder('utf-8');
            workbook = XLSX.read(decoder.decode(ef.data), { type: 'string', raw: false });
          } else {
            workbook = XLSX.read(ef.data, { type: 'array' });
          }
          const isLargeFile = ef.data.byteLength >= PARQUET_THRESHOLD_BYTES;
          const sheetsData = {};
          (workbook.SheetNames || []).forEach(sName => {
            const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sName], { defval: '' });
            if (isLargeFile) {
              const parquetTable = ParquetEngine.encodeToParquet(rows, { fileName: ef.name, sheetName: sName });
              sheetsData[sName] = ParquetEngine.createVirtualArray(parquetTable);
            } else {
              sheetsData[sName] = rows;
            }
          });
          const rawSizeStr = ef.data.byteLength >= 1048576 
            ? (ef.data.byteLength / 1048576).toFixed(1) + ' MB'
            : Math.max(1, Math.round(ef.data.byteLength / 1024)) + ' KB';
          DataStore.files[ef.name] = {
            name: ef.name,
            sizeStr: isLargeFile ? `${rawSizeStr} (📦 Parquet)` : rawSizeStr,
            uploadedAt: new Date().toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit' }),
            sheets: sheetsData,
            rawBytes: ef.data,
            isParquet: isLargeFile
          };
          loadedExcelCount++;
        } catch (exErr) {
          console.warn('Lỗi đọc Excel trong ZIP:', ef.name, exErr);
        }
      }

      if (DataStore.files['Du_Lieu_Mau_Kinh_Doanh.xlsx'] && Object.keys(DataStore.files).length > 1) {
        delete DataStore.files['Du_Lieu_Mau_Kinh_Doanh.xlsx'];
      }

      // 2. Tìm cấu hình từ JSON hoặc HTML
      let config = null;
      if (configFileContent) {
        try { config = JSON.parse(configFileContent); } catch (_) {}
      }
      if (!config && htmlFileContent) {
        config = extractDashboardFromHtml(htmlFileContent);
      }

      if (!config) {
        showToast('Không tìm thấy cấu hình Dashboard hợp lệ trong gói ZIP này.', 'error');
        return false;
      }

      loadDashboardFromConfig(config, {
        successMsg: `✓ Đã nạp thành công trọn bộ Dashboard từ file ZIP "${file.name}" (${loadedExcelCount} file Excel)!`
      });
      closeImportModal();
      return true;
    } catch (err) {
      console.error('Lỗi giải nén ZIP:', err);
      showToast('Lỗi giải nén file ZIP: ' + err.message, 'error');
      return false;
    }
  }

  function importDashboardHtmlFile(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const htmlText = e.target.result;
        const config = extractDashboardFromHtml(htmlText);
        if (!config) {
          showToast('Không tìm thấy cấu hình Dashboard bên trong file HTML này.', 'error');
          return;
        }
        loadDashboardFromConfig(config, {
          successMsg: `✓ Đã trích xuất cấu hình và tái hiện 100% Dashboard từ file HTML "${file.name}"!`
        });
        closeImportModal();
      } catch (err) {
        showToast('Lỗi đọc file HTML: ' + err.message, 'error');
      }
    };
    reader.readAsText(file);
  }

  function importDashboardJsonFile(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const config = JSON.parse(e.target.result);
        if (!config || (!config.items && !config.tabs && !config.title)) {
          showToast('Cấu trúc file JSON không hợp lệ!', 'error');
          return;
        }
        loadDashboardFromConfig(config, {
          successMsg: `✓ Đã nạp thành công Dashboard từ file JSON "${file.name}"!`
        });
        closeImportModal();
      } catch (err) {
        showToast('Lỗi đọc JSON: ' + err.message, 'error');
      }
    };
    reader.readAsText(file);
  }

  function openImportModal() {
    const modal = document.getElementById('dashboardImportModal');
    if (modal) modal.classList.add('open');
  }

  function closeImportModal() {
    const modal = document.getElementById('dashboardImportModal');
    if (modal) modal.classList.remove('open');
  }

  async function importDashboardFolder() {
    // 1. Thử dùng File System Access API (showDirectoryPicker)
    if (window.showDirectoryPicker) {
      try {
        const dirHandle = await window.showDirectoryPicker();
        const filesArray = [];

        for await (const entry of dirHandle.values()) {
          if (entry.kind === 'file') {
            const file = await entry.getFile();
            filesArray.push(file);
          } else if (entry.kind === 'directory') {
            try {
              for await (const subEntry of entry.values()) {
                if (subEntry.kind === 'file') {
                  const subFile = await subEntry.getFile();
                  filesArray.push(subFile);
                }
              }
            } catch (_) {}
          }
        }

        if (filesArray.length === 0) {
          showToast('Thư mục được chọn không chứa tệp tin nào.', 'info');
          return;
        }

        const ok = await restoreDashboardFromFolderFiles(filesArray);
        if (ok) {
          activeDashboardFolderHandle = dirHandle;
          activeDashboardFolderName = dirHandle.name;
          saveActiveWorkspaceHandle(dirHandle, dirHandle.name);
        }
        return;
      } catch (err) {
        if (err.name === 'AbortError') return;
        console.warn('showDirectoryPicker fallback sang thẻ input thư mục:', err);
      }
    }

    // 2. Fallback: kích hoạt thẻ input thư mục
    const folderInput = document.getElementById('folderPickerInput');
    if (folderInput) {
      folderInput.value = '';
      folderInput.click();
    }
  }

  function toggleTheme() {
    document.documentElement.setAttribute('data-theme', 'light');
    DashboardState.present.theme.themeMode = 'light';
    DashboardState.updateUI();
    renderCanvas();
    showToast('Dashboard luôn hoạt động ở chế độ Sáng (Light Mode)', 'info');
  }

  function showToast(message, type = 'info') {
    let container = document.getElementById('dbToastContainer');
    if (!container) {
      container = document.createElement('div');
      container.id = 'dbToastContainer';
      container.className = 'db-toast-container';
      document.body.appendChild(container);
    }
    const toast = document.createElement('div');
    toast.className = `db-toast ${type}`;
    const icon = type === 'success' ? '✓' : type === 'error' ? '✕' : 'ℹ';
    toast.innerHTML = `<span>${icon}</span><span>${message}</span>`;
    container.appendChild(toast);
    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(10px)';
      setTimeout(() => toast.remove(), 250);
    }, 3200);
  }

  // ===========================================================================
  // ===========================================================================
  // 11. EXCEL / CSV DATA SOURCE MANAGER & PREVIEW (THÊM, SỬA, XÓA, XEM TRƯỚC)
  // ===========================================================================
  let activePreviewFile = null;
  let activePreviewSheet = null;

  function handleExcelFilesUpload(filesList) {
    if (!filesList || filesList.length === 0) return;

    const filesArray = Array.from(filesList).filter(f => {
      const lower = f.name.toLowerCase();
      return lower.endsWith('.xlsx') || lower.endsWith('.xls') || lower.endsWith('.csv');
    });

    if (filesArray.length === 0) {
      showToast('Vui lòng chọn file định dạng .xlsx, .xls hoặc .csv', 'error');
      return;
    }

    let processed = 0;
    let newlyAddedFirstKey = null;

    filesArray.forEach(async (file) => {
      const isCsv = file.name.toLowerCase().endsWith('.csv');
      const isLargeFile = file.size >= PARQUET_THRESHOLD_BYTES;
      const sizeStr = file.size > 1048576 
        ? (file.size / 1048576).toFixed(1) + ' MB' 
        : Math.max(1, Math.round(file.size / 1024)) + ' KB';

      try {
        const sheetsData = {};
        let sheetNames = [];

        if (isCsv) {
          // Sử dụng Fast Streaming CSV Parser: Đọc trực tiếp thành Columnar Table, không qua SheetJS, tiết kiệm 95% RAM
          const parquetTable = await FastCsvParser.parse(file);
          const sName = file.name.replace(/\.[^/.]+$/, '');
          sheetsData[sName] = ParquetEngine.createVirtualArray(parquetTable);
          sheetNames = [sName];
          if (isLargeFile) {
            showToast(`📦 File CSV lớn "${file.name}" (${sizeStr}): Đã kích hoạt Streaming Columnar Parser siêu nhẹ.`, 'info', 5000);
          }
        } else {
          // File Excel XLSX / XLS: Tối ưu với dense: true và giải phóng bộ nhớ từng sheet
          const arrayBuffer = await file.arrayBuffer();
          const workbook = XLSX.read(new Uint8Array(arrayBuffer), {
            type: 'array',
            dense: true,
            cellFormula: false,
            cellHTML: false,
            cellText: false
          });
          sheetNames = workbook.SheetNames || [];

          sheetNames.forEach(sName => {
            const worksheet = workbook.Sheets[sName];
            const rows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
            delete workbook.Sheets[sName]; // Giải phóng ngay sheet khỏi bộ nhớ
            if (isLargeFile && rows.length > 500) {
              const parquetTable = ParquetEngine.encodeToParquet(rows, { fileName: file.name, sheetName: sName });
              sheetsData[sName] = ParquetEngine.createVirtualArray(parquetTable);
            } else {
              sheetsData[sName] = rows;
            }
          });

          if (isLargeFile) {
            showToast(`📦 File Excel lớn "${file.name}" (${sizeStr}): Đã tối ưu hóa bộ nhớ và nén Columnar thành công!`, 'info', 5000);
          }
        }

        // Lưu trữ dữ liệu vào DataStore (giải phóng rawBytes nhị phân cho các file >= 5MB)
        DataStore.files[file.name] = {
          name: file.name,
          sizeStr: isLargeFile ? `${sizeStr} (📦 Parquet)` : sizeStr,
          uploadedAt: new Date().toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit' }),
          sheets: sheetsData,
          isParquet: isLargeFile,
          rawBytes: file.size < 5 * 1024 * 1024 && !isCsv ? null : null
        };

        if (!newlyAddedFirstKey && sheetNames.length > 0) {
          newlyAddedFirstKey = `${file.name}::${sheetNames[0]}`;
        }
      } catch (err) {
        console.error('Lỗi phân tích file:', file.name, err);
        showToast(`Lỗi phân tích file "${file.name}": ${err.message}`, 'error');
      }

      processed++;
      if (processed === filesArray.length) {
            // Nếu đây là lần nạp file thật đầu tiên, gỡ bỏ file demo mẫu để ưu tiên dữ liệu thật
            if (DataStore.files['Du_Lieu_Mau_Kinh_Doanh.xlsx'] && Object.keys(DataStore.files).length > 1) {
              delete DataStore.files['Du_Lieu_Mau_Kinh_Doanh.xlsx'];
            }

            // Tự động chuyển các chart đang dùng mock data sang nguồn file mới nạp
            if (newlyAddedFirstKey) {
              DashboardState.present.items.forEach(item => {
                if (!item.dataSourceKey || item.dataSourceKey.startsWith('Du_Lieu_Mau_Kinh_Doanh.xlsx')) {
                  item.dataSourceKey = newlyAddedFirstKey;
                }
              });
            }

            updateDataSourcesSummary();
            updateExcelButtonState();
            renderCanvas();

            // Mở ngay popup quản lý nguồn dữ liệu để người dùng kiểm tra dữ liệu và xem trước các sheet
            openExcelManagerModal();
            showToast(`✓ Đã nạp thành công ${filesArray.length} file nguồn dữ liệu!`, 'success');
          }
    });
  }

  function updateDataSourcesSummary() {
    const container = document.getElementById('dbSourceChips');
    if (!container) return;
    container.innerHTML = '';
    const sources = DataStore.getAvailableSources();
    sources.forEach(s => {
      const chip = document.createElement('span');
      chip.className = 'db-source-chip';
      chip.textContent = `📁 ${s.label}`;
      chip.title = 'Click để xem trước dữ liệu file này';
      chip.style.cursor = 'pointer';
      chip.addEventListener('click', () => {
        openExcelManagerModal();
        renderExcelSheetPreview(s.fileName, s.sheetName);
      });
      container.appendChild(chip);
    });
  }

  function updateExcelButtonState() {
    const btn = document.getElementById('btnUploadExcelFiles');
    if (!btn) return;

    const fileNames = Object.keys(DataStore.files).filter(f => f !== 'Du_Lieu_Mau_Kinh_Doanh.xlsx');
    const hasCustomFiles = fileNames.length > 0;

    if (hasCustomFiles) {
      btn.dataset.mode = 'manage';
      btn.innerHTML = `🗂️ Quản Lý Nguồn Dữ Liệu (${fileNames.length})`;
      btn.title = `Đã nạp ${fileNames.length} file dữ liệu. Bấm để mở popup Quản lý (Thêm, Sửa, Xóa, Xem trước).`;
      btn.style.background = 'linear-gradient(135deg, #0284c7 0%, #0369a1 100%)';
      btn.style.borderColor = '#0284c7';
    } else {
      btn.dataset.mode = 'upload';
      btn.innerHTML = `📤 Nạp File Excel / CSV`;
      btn.title = 'Tải một hoặc nhiều file Excel (.xlsx, .xls) hoặc .csv lên làm nguồn dữ liệu cho Dashboard';
      btn.style.background = 'linear-gradient(135deg, #0284c7 0%, #0369a1 100%)';
      btn.style.borderColor = '#0284c7';
    }
  }

  function openExcelManagerModal() {
    const modal = document.getElementById('excelFilesModal');
    if (!modal) return;
    renderLoadedExcelFilesList();
    modal.classList.add('open');
  }

  function closeExcelManagerModal() {
    const modal = document.getElementById('excelFilesModal');
    if (!modal) return;
    modal.classList.remove('open');
  }

  function renderLoadedExcelFilesList() {
    const container = document.getElementById('loadedExcelFilesList');
    if (!container) return;
    container.innerHTML = '';

    const badge = document.getElementById('excelFilesTotalBadge');
    const fileNames = Object.keys(DataStore.files);
    const customFiles = fileNames.filter(f => f !== 'Du_Lieu_Mau_Kinh_Doanh.xlsx');

    if (badge) {
      badge.textContent = `${fileNames.length} tệp nguồn (${customFiles.length} file do bạn nạp)`;
    }

    if (fileNames.length === 0) {
      container.innerHTML = `
        <div style="color:var(--db-text-sub);text-align:center;padding:24px;background:#f8fafc;border-radius:8px;border:1px dashed #cbd5e1;font-size:13px;">
          Chưa có file nguồn dữ liệu nào. Kéo thả hoặc bấm <strong>+ Nạp Thêm File</strong> phía trên để đưa dữ liệu vào.
        </div>
      `;
      const previewWrapper = document.getElementById('excelPreviewTableWrapper');
      if (previewWrapper) previewWrapper.innerHTML = '<div style="color:var(--db-text-sub);text-align:center;padding:20px;font-size:12px;">Chưa có dữ liệu để xem trước.</div>';
      return;
    }

    fileNames.forEach(fileName => {
      const fileObj = DataStore.files[fileName];
      const sheetNames = Object.keys(fileObj.sheets || {});
      let totalRows = 0;
      sheetNames.forEach(s => {
        const sh = fileObj.sheets[s];
        totalRows += sh ? (sh.length !== undefined ? sh.length : (sh.rowCount || 0)) : 0;
      });

      const isSample = fileName === 'Du_Lieu_Mau_Kinh_Doanh.xlsx';
      const isCsv = fileName.toLowerCase().endsWith('.csv');
      const isPreviewingThisFile = activePreviewFile === fileName;

      const card = document.createElement('div');
      card.className = `db-excel-file-card ${isPreviewingThisFile ? 'active-preview' : ''}`;

      // Tạo danh sách nút chọn Sheet
      const sheetsHtml = sheetNames.map(sName => {
        const sh = fileObj.sheets[sName];
        const rowCount = sh ? (sh.length !== undefined ? sh.length : (sh.rowCount || 0)) : 0;
        const isActive = isPreviewingThisFile && activePreviewSheet === sName;
        return `<button class="db-excel-sheet-pill ${isActive ? 'active' : ''}" data-file="${escapeHtml(fileName)}" data-sheet="${escapeHtml(sName)}" title="Xem trước dữ liệu sheet này">📑 ${escapeHtml(sName)} (${rowCount.toLocaleString('vi-VN')} dòng)</button>`;
      }).join(' ');

      card.innerHTML = `
        <div style="flex:1;min-width:0;">
          <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
            <span style="font-size:16px;">${isCsv ? '📊' : '📗'}</span>
            <strong style="color:#0f172a;font-size:13px;word-break:break-all;">${escapeHtml(fileName)}</strong>
            <span style="font-size:10px;background:${isCsv ? '#e0f2fe' : '#dcfce7'};color:${isCsv ? '#0369a1' : '#15803d'};border:1px solid ${isCsv ? '#bae6fd' : '#bbf7d0'};padding:1px 6px;border-radius:4px;font-weight:700;">
              ${isCsv ? 'CSV' : 'EXCEL'}
            </span>
            ${fileObj.isParquet ? '<span style="font-size:10px;background:#ede9fe;color:#6d28d9;border:1px solid #ddd6fe;padding:1px 6px;border-radius:4px;font-weight:700;" title="Dữ liệu đã được nén Parquet Columnar để tiết kiệm RAM trình duyệt">📦 Nén Parquet (&gt;50MB)</span>' : ''}
            ${isSample ? '<span style="font-size:10px;background:#f1f5f9;color:#64748b;padding:1px 6px;border-radius:4px;">File mẫu mặc định</span>' : ''}
            <span style="font-size:11px;color:var(--db-text-sub);">${fileObj.sizeStr || ''} ${fileObj.uploadedAt ? '• Nạp lúc ' + fileObj.uploadedAt : ''}</span>
          </div>

          <div style="margin-top:6px;display:flex;align-items:center;gap:6px;flex-wrap:wrap;">
            <span style="font-size:11px;color:#475569;font-weight:600;">Sheets:</span>
            ${sheetsHtml}
          </div>
        </div>

        <div style="display:flex;gap:6px;align-items:center;flex-shrink:0;">
          <button class="db-btn-tool" data-preview-file="${escapeHtml(fileName)}" style="padding:5px 10px;font-size:12px;background:#ffffff;" title="Xem trước bảng dữ liệu sheet đầu tiên">
            👁️ Xem
          </button>
          <button class="db-btn-tool" data-rename-file="${escapeHtml(fileName)}" style="padding:5px 10px;font-size:12px;background:#ffffff;" title="Đổi tên file / alias nguồn dữ liệu">
            ✏️ Sửa Tên
          </button>
          <button class="db-btn-tool" data-remove-file="${escapeHtml(fileName)}" style="padding:5px 10px;font-size:12px;color:#ef4444;border-color:rgba(239,68,68,0.3);background:#ffffff;" title="Xóa file này khỏi danh sách nguồn">
            🗑️ Xóa
          </button>
        </div>
      `;
      container.appendChild(card);
    });

    // 1. Sự kiện bấm chọn Sheet để xem trước
    container.querySelectorAll('.db-excel-sheet-pill').forEach(pill => {
      pill.addEventListener('click', (e) => {
        e.stopPropagation();
        const f = pill.dataset.file;
        const s = pill.dataset.sheet;
        renderExcelSheetPreview(f, s);
      });
    });

    // 2. Sự kiện bấm nút Xem
    container.querySelectorAll('[data-preview-file]').forEach(btn => {
      btn.addEventListener('click', () => {
        const f = btn.dataset.previewFile;
        const firstSheet = Object.keys(DataStore.files[f]?.sheets || {})[0];
        if (firstSheet) renderExcelSheetPreview(f, firstSheet);
      });
    });

    // 3. Sự kiện bấm nút Sửa / Đổi Tên (Quyền SỬA)
    container.querySelectorAll('[data-rename-file]').forEach(btn => {
      btn.addEventListener('click', async () => {
        const oldName = btn.dataset.renameFile;
        const newName = await AppDialog.prompt(`Nhập tên mới (alias) cho file "${oldName}":`, oldName, 'Đổi Tên Nguồn Dữ Liệu');
        if (newName && newName.trim() && newName.trim() !== oldName) {
          renameExcelDataSource(oldName, newName.trim());
        }
      });
    });

    // 4. Sự kiện bấm nút Xóa (Quyền XÓA)
    container.querySelectorAll('[data-remove-file]').forEach(btn => {
      btn.addEventListener('click', async () => {
        const fName = btn.dataset.removeFile;
        const ok = await AppDialog.confirm(`Bạn có chắc muốn xóa nguồn dữ liệu "${fName}"? Các biểu đồ đang dùng nguồn này có thể cần được chọn lại.`, 'Xóa Nguồn Dữ Liệu', { isDanger: true, confirmText: 'Xóa Nguồn' });
        if (ok) {
          deleteExcelDataSource(fName);
        }
      });
    });

    // Tự động kích hoạt xem trước sheet đầu tiên nếu chưa có sheet nào được chọn
    if (!activePreviewFile || !DataStore.files[activePreviewFile]) {
      const firstF = fileNames[0];
      const firstS = Object.keys(DataStore.files[firstF]?.sheets || {})[0];
      if (firstF && firstS) renderExcelSheetPreview(firstF, firstS);
    }
  }

  function renderExcelSheetPreview(fileName, sheetName) {
    activePreviewFile = fileName;
    activePreviewSheet = sheetName;

    const fileObj = DataStore.files[fileName];
    if (!fileObj || !fileObj.sheets || !fileObj.sheets[sheetName]) return;

    const rows = fileObj.sheets[sheetName];
    const columns = rows.length > 0 ? Object.keys(rows[0]) : [];

    const titleEl = document.getElementById('excelPreviewTitle');
    const metaEl = document.getElementById('excelPreviewMeta');
    const wrapper = document.getElementById('excelPreviewTableWrapper');

    if (titleEl) titleEl.textContent = `Xem trước: ${fileName} » ${sheetName}`;
    if (metaEl) metaEl.textContent = `Tổng ${rows.length.toLocaleString('vi-VN')} dòng • ${columns.length} cột (Hiển thị 10 dòng đầu)`;

    if (!wrapper) return;

    if (rows.length === 0) {
      wrapper.innerHTML = '<div style="color:var(--db-text-sub);text-align:center;padding:24px;font-size:12px;">Sheet này trống, không có dòng dữ liệu nào.</div>';
      return;
    }

    const previewRows = rows.slice(0, 10);
    let tableHtml = '<table class="db-excel-table"><thead><tr>';
    tableHtml += '<th style="width:40px;text-align:center;">#</th>';
    columns.forEach(col => {
      tableHtml += `<th>${escapeHtml(col)}</th>`;
    });
    tableHtml += '</tr></thead><tbody>';

    previewRows.forEach((row, idx) => {
      tableHtml += `<tr><td style="text-align:center;color:#94a3b8;font-size:11px;">${idx + 1}</td>`;
      columns.forEach(col => {
        const val = row[col] !== undefined && row[col] !== null ? String(row[col]) : '';
        tableHtml += `<td>${escapeHtml(val)}</td>`;
      });
      tableHtml += '</tr>';
    });
    tableHtml += '</tbody></table>';

    wrapper.innerHTML = tableHtml;

    // Cập nhật lại giao diện active của các sheet pills
    document.querySelectorAll('.db-excel-sheet-pill').forEach(pill => {
      if (pill.dataset.file === fileName && pill.dataset.sheet === sheetName) {
        pill.classList.add('active');
      } else {
        pill.classList.remove('active');
      }
    });

    document.querySelectorAll('.db-excel-file-card').forEach(c => {
      c.classList.remove('active-preview');
    });
  }

  function renameExcelDataSource(oldName, newName) {
    if (DataStore.files[newName]) {
      showToast(`Tên "${newName}" đã tồn tại! Vui lòng chọn tên khác.`, 'error');
      return;
    }

    DashboardState.pushSnapshot();
    const fileObj = DataStore.files[oldName];
    fileObj.name = newName;
    DataStore.files[newName] = fileObj;
    delete DataStore.files[oldName];

    // Cập nhật lại các chart đang dùng dataSourceKey trỏ đến oldName
    DashboardState.present.items.forEach(item => {
      if (item.dataSourceKey && item.dataSourceKey.startsWith(oldName + '::')) {
        item.dataSourceKey = item.dataSourceKey.replace(oldName + '::', newName + '::');
      }
    });

    if (activePreviewFile === oldName) activePreviewFile = newName;

    updateDataSourcesSummary();
    renderLoadedExcelFilesList();
    renderCanvas();
    updateExcelButtonState();
    showToast(`✓ Đã đổi tên nguồn dữ liệu thành "${newName}"!`, 'success');
  }

  function deleteExcelDataSource(fileName) {
    DashboardState.pushSnapshot();
    delete DataStore.files[fileName];

    // Nếu không còn file nào thì khởi tạo lại mẫu mặc định
    if (Object.keys(DataStore.files).length === 0) {
      DataStore.initDefault();
      activePreviewFile = 'Du_Lieu_Mau_Kinh_Doanh.xlsx';
      activePreviewSheet = 'BaoCaoDoanhSo';
    } else {
      activePreviewFile = null;
      activePreviewSheet = null;
    }

    updateDataSourcesSummary();
    renderLoadedExcelFilesList();
    renderCanvas();
    updateExcelButtonState();
    showToast(`Đã xóa nguồn dữ liệu "${fileName}"`, 'info');
  }

  // ===========================================================================
  // 12. CANVAS ZOOM, SNAPPING, ALIGNMENT & INTERACTIVE DRAG/RESIZE
  // ===========================================================================
  const SNAP_THRESHOLD = 8;

  function updateCanvasTransform() {
    const canvas = document.getElementById('dashboardCanvas');
    if (!canvas) return;

    const wrapper = document.getElementById('canvasWrapper') || canvas.parentElement;
    const availW = wrapper ? wrapper.clientWidth : window.innerWidth;
    const safePaddingX = 48; // Luôn đảm bảo tối thiểu 24px đệm 2 bên lề cân xứng
    const autoFitScale = Math.min(1.0, Math.max(0.2, (availW - safePaddingX) / 1360));
    const curH = parseInt(canvas.style.minHeight || '800', 10);

    if (DashboardState.mode === 'preview') {
      if (autoFitScale < 1.0) {
        canvas.style.transform = `scale(${autoFitScale})`;
        canvas.style.transformOrigin = 'top center';
        canvas.style.marginBottom = `-${Math.round(curH * (1 - autoFitScale))}px`;
      } else {
        canvas.style.transform = 'none';
        canvas.style.transformOrigin = 'top center';
        canvas.style.marginBottom = '0px';
      }
      return;
    }

    // Chế độ Edit:
    const userZoom = DashboardState.zoom || 1.0;
    const pan = DashboardState.pan || { x: 0, y: 0 };

    let effectiveScale = autoFitScale;
    if (userZoom !== 1.0) {
      effectiveScale = userZoom;
    }

    if (effectiveScale < 1.0 && pan.x === 0 && pan.y === 0) {
      canvas.style.transform = `scale(${effectiveScale})`;
      canvas.style.transformOrigin = 'top center';
      canvas.style.marginBottom = `-${Math.round(curH * (1 - effectiveScale))}px`;
    } else if (pan.x !== 0 || pan.y !== 0 || userZoom !== 1.0) {
      canvas.style.transform = `translate(${pan.x}px, ${pan.y}px) scale(${effectiveScale})`;
      canvas.style.transformOrigin = 'top center';
      canvas.style.marginBottom = '0px';
    } else {
      canvas.style.transform = 'none';
      canvas.style.transformOrigin = 'top center';
      canvas.style.marginBottom = '0px';
    }

    if (typeof updateTopFilterBarStickyPosition === 'function') {
      updateTopFilterBarStickyPosition();
    }
  }

  function setZoom(newZoom) {
    if (DashboardState.mode === 'preview') return;
    const clamped = Math.min(2.5, Math.max(0.3, Math.round(newZoom * 10) / 10));
    DashboardState.zoom = clamped;
    const badge = document.getElementById('zoomLevelBadge');
    if (badge) badge.textContent = `${Math.round(clamped * 100)}%`;
    updateCanvasTransform();
    // Resize ECharts instances
    document.querySelectorAll('.db-chart-canvas').forEach(el => {
      const inst = echarts.getInstanceByDom(el);
      if (inst) inst.resize();
    });
  }

  function resetZoomAndPan() {
    if (DashboardState.mode === 'preview') return;
    DashboardState.zoom = 1.0;
    DashboardState.pan = { x: 0, y: 0 };
    const badge = document.getElementById('zoomLevelBadge');
    if (badge) badge.textContent = '100%';
    updateCanvasTransform();
    document.querySelectorAll('.db-chart-canvas').forEach(el => {
      const inst = echarts.getInstanceByDom(el);
      if (inst) inst.resize();
    });
    showToast('Đã khôi phục Zoom 100% & Căn Giữa Canvas', 'info');
  }

  function initFigmaCanvasPan() {
    const wrapper = document.getElementById('canvasWrapper');
    const viewport = document.getElementById('canvasViewport');
    const canvas = document.getElementById('dashboardCanvas');
    if (!wrapper || !canvas) return;
    if (wrapper.dataset.panInitialized === 'true') return;
    wrapper.dataset.panInitialized = 'true';

    let isSpacePressed = false;
    let isPanning = false;
    let startMouseX = 0;
    let startMouseY = 0;
    let startPanX = 0;
    let startPanY = 0;

    // Phím Space để bật Hand tool như Figma
    window.addEventListener('keydown', (e) => {
      if (DashboardState.mode === 'preview') return;
      if (e.code === 'Space' && !['INPUT', 'TEXTAREA'].includes(document.activeElement?.tagName) && !document.activeElement?.isContentEditable) {
        if (!isSpacePressed) {
          isSpacePressed = true;
          wrapper.style.cursor = 'grab';
        }
      }
    });

    window.addEventListener('keyup', (e) => {
      if (e.code === 'Space') {
        isSpacePressed = false;
        if (!isPanning) {
          wrapper.style.cursor = '';
        }
      }
    });

    // Mousedown để bắt đầu pan
    wrapper.addEventListener('mousedown', (e) => {
      if (DashboardState.mode === 'preview') return;
      const isMiddle = e.button === 1;
      const isLeft = e.button === 0;
      const clickedOnBg = e.target === wrapper || e.target === viewport || e.target === canvas || e.target.classList.contains('db-empty-canvas') || e.target.closest('.db-empty-canvas');

      const shouldPan = isMiddle || (isSpacePressed && isLeft) || (clickedOnBg && isLeft && !e.target.closest('.db-item-box') && !e.target.closest('.figma-floating-toolbar') && !e.target.closest('#canvaTextToolbar'));

      if (shouldPan) {
        e.preventDefault();
        isPanning = true;
        startMouseX = e.clientX;
        startMouseY = e.clientY;
        startPanX = DashboardState.pan.x || 0;
        startPanY = DashboardState.pan.y || 0;
        wrapper.style.cursor = 'grabbing';
        document.body.style.userSelect = 'none';

        function onMouseMove(ev) {
          if (!isPanning) return;
          const dx = ev.clientX - startMouseX;
          const dy = ev.clientY - startMouseY;
          DashboardState.pan.x = startPanX + dx;
          DashboardState.pan.y = startPanY + dy;
          updateCanvasTransform();
        }

        function onMouseUp() {
          isPanning = false;
          wrapper.style.cursor = isSpacePressed ? 'grab' : '';
          document.body.style.userSelect = '';
          window.removeEventListener('mousemove', onMouseMove);
          window.removeEventListener('mouseup', onMouseUp);
        }

        window.addEventListener('mousemove', onMouseMove);
        window.addEventListener('mouseup', onMouseUp);
      }
    });

    // Cuộn chuột: Ctrl/Meta = Zoom, Shift = Cuộn ngang, Cuộn dọc = Native scroll tự nhiên
    wrapper.addEventListener('wheel', (e) => {
      if (DashboardState.mode === 'preview') return;
      if (e.ctrlKey || e.metaKey) {
        e.preventDefault();
        const delta = e.deltaY < 0 ? 0.1 : -0.1;
        setZoom((DashboardState.zoom || 1.0) + delta);
      } else if (e.shiftKey) {
        e.preventDefault();
        wrapper.scrollLeft += e.deltaY;
      }
      // Khi cuộn dọc bình thường: Để trình duyệt cuộn dọc tự nhiên qua wrapper.scrollTop
    }, { passive: false });
  }

  function calculateSnapping(currentPos, excludeId) {
    const guides = [];
    let snappedX = currentPos.x;
    let snappedY = currentPos.y;

    const activeTabId = DashboardState.present.activeTabId;
    const others = DashboardState.present.items
      .filter(i => i.id !== excludeId && (i.tabId || 'default') === activeTabId)
      .map(i => i.pos);

    const currentRight = currentPos.x + currentPos.w;
    const currentBottom = currentPos.y + currentPos.h;
    const currentCenterX = currentPos.x + currentPos.w / 2;
    const currentCenterY = currentPos.y + currentPos.h / 2;

    let closestDeltaX = SNAP_THRESHOLD;
    let closestDeltaY = SNAP_THRESHOLD;

    const xTargets = [];
    const yTargets = [];

    // Điểm bắt dính 2 ranh giới an toàn Dashboard (Lề trái 24px, Tâm 680px, Lề phải 1336px)
    xTargets.push(24, 1336, 680);
    yTargets.push(24);

    others.forEach(other => {
      const otherRight = other.x + other.w;
      const otherCenterX = other.x + other.w / 2;
      xTargets.push(other.x, otherRight, otherCenterX);

      const otherBottom = other.y + other.h;
      const otherCenterY = other.y + other.h / 2;
      yTargets.push(other.y, otherBottom, otherCenterY);
    });

    xTargets.forEach(targetVal => {
      const deltas = [
        { delta: currentPos.x - targetVal, snapVal: targetVal },
        { delta: currentRight - targetVal, snapVal: targetVal - currentPos.w },
        { delta: currentCenterX - targetVal, snapVal: targetVal - currentPos.w / 2 }
      ];
      deltas.forEach(d => {
        if (Math.abs(d.delta) < Math.abs(closestDeltaX)) {
          closestDeltaX = d.delta;
          snappedX = d.snapVal;
        }
      });
    });

    yTargets.forEach(targetVal => {
      const deltas = [
        { delta: currentPos.y - targetVal, snapVal: targetVal },
        { delta: currentBottom - targetVal, snapVal: targetVal - currentPos.h },
        { delta: currentCenterY - targetVal, snapVal: targetVal - currentPos.h / 2 }
      ];
      deltas.forEach(d => {
        if (Math.abs(d.delta) < Math.abs(closestDeltaY)) {
          closestDeltaY = d.delta;
          snappedY = d.snapVal;
        }
      });
    });

    const finalRight = snappedX + currentPos.w;
    const finalCenterX = snappedX + currentPos.w / 2;
    xTargets.forEach(targetVal => {
      if (Math.abs(snappedX - targetVal) < 0.5 || Math.abs(finalRight - targetVal) < 0.5 || Math.abs(finalCenterX - targetVal) < 0.5) {
        if (!guides.some(g => g.x === targetVal && g.type === 'v')) {
          guides.push({ x: targetVal, type: 'v' });
        }
      }
    });

    const finalBottom = snappedY + currentPos.h;
    const finalCenterY = snappedY + currentPos.h / 2;
    yTargets.forEach(targetVal => {
      if (Math.abs(snappedY - targetVal) < 0.5 || Math.abs(finalBottom - targetVal) < 0.5 || Math.abs(finalCenterY - targetVal) < 0.5) {
        if (!guides.some(g => g.y === targetVal && g.type === 'h')) {
          guides.push({ y: targetVal, type: 'h' });
        }
      }
    });

    return { snappedX: Math.round(snappedX), snappedY: Math.round(snappedY), guides };
  }

  function showSnappingGuides(guides) {
    const container = document.getElementById('snappingGuides');
    if (!container) return;
    container.innerHTML = '';
    guides.forEach(g => {
      const line = document.createElement('div');
      if (g.type === 'v') {
        line.className = 'db-snap-guide-v';
        line.style.left = `${g.x}px`;
      } else {
        line.className = 'db-snap-guide-h';
        line.style.top = `${g.y}px`;
      }
      container.appendChild(line);
    });
  }

  function clearSnappingGuides() {
    const container = document.getElementById('snappingGuides');
    if (container) container.innerHTML = '';
  }

  function alignWidgets(direction) {
    const currentTabId = DashboardState.present.activeTabId;
    const items = DashboardState.present.items.filter(i => (i.tabId || 'default') === currentTabId);
    if (items.length < 2) {
      showToast('Cần ít nhất 2 widget để căn chỉnh.', 'info');
      return;
    }

    DashboardState.pushSnapshot();

    if (direction === 'left') {
      const minX = Math.min(...items.map(i => i.pos.x));
      items.forEach(i => { i.pos.x = minX; });
      showToast('Đã căn lề trái', 'success');
    } else if (direction === 'right') {
      const maxRight = Math.max(...items.map(i => i.pos.x + i.pos.w));
      items.forEach(i => { i.pos.x = maxRight - i.pos.w; });
      showToast('Đã căn lề phải', 'success');
    } else if (direction === 'top') {
      const minY = Math.min(...items.map(i => i.pos.y));
      items.forEach(i => { i.pos.y = minY; });
      showToast('Đã căn lề trên', 'success');
    } else if (direction === 'bottom') {
      const maxBottom = Math.max(...items.map(i => i.pos.y + i.pos.h));
      items.forEach(i => { i.pos.y = maxBottom - i.pos.h; });
      showToast('Đã căn lề dưới', 'success');
    } else if (direction === 'center') {
      const avgCenterX = items.reduce((acc, i) => acc + (i.pos.x + i.pos.w / 2), 0) / items.length;
      items.forEach(i => { i.pos.x = Math.round(avgCenterX - i.pos.w / 2); });
      showToast('Đã căn giữa', 'success');
    } else if (direction === 'distributeH') {
      const sorted = [...items].sort((a, b) => a.pos.x - b.pos.x);
      const first = sorted[0];
      const last = sorted[sorted.length - 1];
      const totalWidth = sorted.reduce((sum, item) => sum + item.pos.w, 0);
      const span = (last.pos.x + last.pos.w) - first.pos.x;
      const remainingSpace = span - totalWidth;
      const gap = remainingSpace / (sorted.length - 1);
      let curX = first.pos.x;
      sorted.forEach((item) => {
        item.pos.x = Math.round(curX);
        curX += item.pos.w + gap;
      });
      showToast('Đã giãn đều khoảng cách ngang', 'success');
    }

    renderCanvas();
  }

  // ===========================================================================
  // 11.4 TAB BAR CUSTOMIZATION & 4-WAY LAYOUT ENGINE
  // ===========================================================================
  const TabBarManager = {
    DEFAULT_CONFIG: {
      position: 'top',       // 'top', 'bottom', 'left', 'right'
      collapsed: false,      // true | false
      style: 'pills',        // 'pills', 'cards', 'underline'
      size: 'normal',        // 'compact', 'normal', 'comfortable'
      accentColor: '#0284c7',
      showIndex: false,
      showCloseBtn: true,
      verticalWidth: 230     // 180 - 360 px
    },

    getConfig() {
      if (DashboardState.present && DashboardState.present.tabBarConfig) {
        return { ...this.DEFAULT_CONFIG, ...DashboardState.present.tabBarConfig };
      }
      try {
        const saved = localStorage.getItem('mb_tab_bar_config_v1');
        if (saved) return { ...this.DEFAULT_CONFIG, ...JSON.parse(saved) };
      } catch (e) {}
      return { ...this.DEFAULT_CONFIG };
    },

    saveConfig(cfg) {
      if (!DashboardState.present.tabBarConfig) DashboardState.present.tabBarConfig = {};
      Object.assign(DashboardState.present.tabBarConfig, cfg);
      DashboardState.isDirty = true;
      DashboardState.updateUI();
      try {
        localStorage.setItem('mb_tab_bar_config_v1', JSON.stringify(cfg));
      } catch (e) {}
    },

    apply(cfg = null) {
      const config = cfg || this.getConfig();
      const container = document.getElementById('dbStudioContainer');
      const tabsBar = document.getElementById('dbTabsBar');
      const collapseIcon = document.getElementById('dbCollapseIcon');
      const collapseBtn = document.getElementById('btnCollapseTabsBar');

      if (!container || !tabsBar) return;

      // 1. Áp dụng vị trí 4 hướng
      container.classList.remove('tabs-pos-top', 'tabs-pos-bottom', 'tabs-pos-left', 'tabs-pos-right');
      container.classList.add(`tabs-pos-${config.position || 'top'}`);

      // 2. Áp dụng thu gọn / mở rộng
      container.classList.toggle('tabs-collapsed', Boolean(config.collapsed));

      // 3. Cập nhật biểu tượng nút thu gọn nhanh theo hướng (chỉ hiển thị mũi tên lên/xuống/trái/phải, không hiển thị chữ)
      if (collapseBtn) {
        let icon = '▲';
        let tooltip = '';
        if (config.position === 'bottom') {
          icon = config.collapsed ? '▲' : '▼';
          tooltip = config.collapsed ? 'Mở rộng thanh Tab lên trên' : 'Thu gọn thanh Tab xuống dưới';
        } else if (config.position === 'left') {
          icon = config.collapsed ? '▶' : '◀';
          tooltip = config.collapsed ? 'Mở rộng thanh Tab sang phải' : 'Thu gọn thanh Tab sang trái';
        } else if (config.position === 'right') {
          icon = config.collapsed ? '◀' : '▶';
          tooltip = config.collapsed ? 'Mở rộng thanh Tab sang trái' : 'Thu gọn thanh Tab sang phải';
        } else {
          // top
          icon = config.collapsed ? '▼' : '▲';
          tooltip = config.collapsed ? 'Mở rộng thanh Tab xuống dưới' : 'Thu gọn thanh Tab lên trên';
        }
        collapseBtn.innerHTML = `<span class="db-collapse-icon" id="dbCollapseIcon">${icon}</span>`;
        collapseBtn.title = tooltip;
      }

      // 4. Áp dụng kiểu dáng & kích cỡ
      tabsBar.classList.remove('tab-style-pills', 'tab-style-cards', 'tab-style-underline');
      tabsBar.classList.add(`tab-style-${config.style || 'pills'}`);

      tabsBar.classList.remove('tab-size-compact', 'tab-size-normal', 'tab-size-comfortable');
      tabsBar.classList.add(`tab-size-${config.size || 'normal'}`);

      // 5. Áp dụng màu Accent & Chiều rộng dọc
      const color = config.accentColor || '#0284c7';
      tabsBar.style.setProperty('--tab-accent', color);

      if (config.verticalWidth) {
        tabsBar.style.setProperty('--tab-vertical-width', `${config.verticalWidth}px`);
        container.style.setProperty('--tab-vertical-width', `${config.verticalWidth}px`);
      }

      // Render lại tabs
      renderAllTabs();
    },

    toggleCollapse() {
      const cfg = this.getConfig();
      cfg.collapsed = !cfg.collapsed;
      this.saveConfig(cfg);
      this.apply(cfg);
      if (typeof TabConfigSidebar !== 'undefined' && TabConfigSidebar.syncTabUI) {
        TabConfigSidebar.syncTabUI();
      } else if (typeof FigmaPropertiesSidebar !== 'undefined' && FigmaPropertiesSidebar.syncTabUI) {
        FigmaPropertiesSidebar.syncTabUI();
      }
      showToast(cfg.collapsed ? 'Đã thu gọn thanh Tab' : 'Đã mở rộng thanh Tab', 'info');
    },

    openModal() {
      if (typeof TabConfigSidebar !== 'undefined' && TabConfigSidebar.open) {
        TabConfigSidebar.open();
        return;
      }
      if (typeof FigmaPropertiesSidebar !== 'undefined' && FigmaPropertiesSidebar.openTabConfig) {
        FigmaPropertiesSidebar.openTabConfig();
        return;
      }
      const modal = document.getElementById('tabBarConfigModal');
      if (!modal) return;
      const cfg = this.getConfig();

      // Vị trí
      const posRadios = document.querySelectorAll('input[name="tabPositionRadio"]');
      posRadios.forEach(r => { r.checked = (r.value === cfg.position); });

      // Trạng thái thu gọn
      const stateRadios = document.querySelectorAll('input[name="tabCollapsedRadio"]');
      stateRadios.forEach(r => { r.checked = (r.value === (cfg.collapsed ? 'collapsed' : 'expanded')); });

      // Style & Size
      const styleSelect = document.getElementById('tcTabStyleSelect');
      if (styleSelect) styleSelect.value = cfg.style || 'pills';

      const sizeSelect = document.getElementById('tcTabSizeSelect');
      if (sizeSelect) sizeSelect.value = cfg.size || 'normal';

      // Màu sắc
      const colorDots = document.querySelectorAll('#tcColorPalette .tc-color-dot');
      colorDots.forEach(dot => {
        dot.classList.toggle('active', dot.dataset.color === cfg.accentColor);
      });

      // Checkboxes
      const showIndexCb = document.getElementById('tcShowIndexCheckbox');
      if (showIndexCb) showIndexCb.checked = Boolean(cfg.showIndex);

      const showCloseCb = document.getElementById('tcShowCloseBtnCheckbox');
      if (showCloseCb) showCloseCb.checked = cfg.showCloseBtn !== false;

      modal.style.display = 'flex';
      modal.classList.add('open');
    },

    closeModal() {
      const modal = document.getElementById('tabBarConfigModal');
      if (!modal) return;
      modal.style.display = 'none';
      modal.classList.remove('open');
    },

    saveFromModal() {
      const posRadio = document.querySelector('input[name="tabPositionRadio"]:checked');
      const stateRadio = document.querySelector('input[name="tabCollapsedRadio"]:checked');
      const styleSelect = document.getElementById('tcTabStyleSelect');
      const sizeSelect = document.getElementById('tcTabSizeSelect');
      const activeColorDot = document.querySelector('#tcColorPalette .tc-color-dot.active');
      const showIndexCb = document.getElementById('tcShowIndexCheckbox');
      const showCloseCb = document.getElementById('tcShowCloseBtnCheckbox');

      DashboardState.pushSnapshot();

      const newCfg = {
        position: posRadio ? posRadio.value : 'top',
        collapsed: stateRadio ? stateRadio.value === 'collapsed' : false,
        style: styleSelect ? styleSelect.value : 'pills',
        size: sizeSelect ? sizeSelect.value : 'normal',
        accentColor: activeColorDot ? activeColorDot.dataset.color : '#0284c7',
        showIndex: showIndexCb ? showIndexCb.checked : false,
        showCloseBtn: showCloseCb ? showCloseCb.checked : true
      };

      this.saveConfig(newCfg);
      this.apply(newCfg);
      this.closeModal();
      showToast('Đã áp dụng cấu hình thanh Tab thành công!', 'success');
    },

    resetDefaults() {
      const cfg = { ...this.DEFAULT_CONFIG };
      this.saveConfig(cfg);
      this.apply(cfg);
      if (typeof TabConfigSidebar !== 'undefined' && TabConfigSidebar.syncTabUI) {
        TabConfigSidebar.syncTabUI();
      } else if (typeof FigmaPropertiesSidebar !== 'undefined' && FigmaPropertiesSidebar.syncTabUI) {
        FigmaPropertiesSidebar.syncTabUI();
      }
      showToast('Đã khôi phục cấu hình thanh Tab về mặc định!', 'info');
    }
  };

  function switchDashboardTab(targetTabId) {
    if (!targetTabId) return;
    const currentTab = DashboardState.present.activeTabId;
    if (currentTab === targetTabId) return;

    syncItemPositionsFromDOM();

    // 1. Lưu bộ lọc của Tab hiện tại
    if (!DashboardState.tabFilters) DashboardState.tabFilters = {};
    if (currentTab) {
      DashboardState.tabFilters[currentTab] = JSON.parse(JSON.stringify(DashboardState.activeFilters || {}));
    }

    // 2. Chuyển sang Tab đích
    DashboardState.present.activeTabId = targetTabId;

    // 3. Khôi phục bộ lọc độc lập của Tab đích
    DashboardState.activeFilters = JSON.parse(JSON.stringify((DashboardState.tabFilters && DashboardState.tabFilters[targetTabId]) || {}));

    // 4. Đồng bộ vào present để undo/redo và xuất JSON trọn vẹn
    DashboardState.present.tabFilters = JSON.parse(JSON.stringify(DashboardState.tabFilters));
    DashboardState.present.activeFilters = JSON.parse(JSON.stringify(DashboardState.activeFilters));

    // 5. Cập nhật UI
    renderAllTabs();
    renderCanvas();
    triggerDebouncedAutoSave();
  }

  function renderAllTabs() {
    const tabsList = document.getElementById('dbTabsList');
    if (!tabsList) return;
    tabsList.innerHTML = '';

    const cfg = TabBarManager.getConfig();

    DashboardState.present.tabs.forEach((tab, index) => {
      const pill = document.createElement('button');
      pill.className = `db-tab-pill ${tab.id === DashboardState.present.activeTabId ? 'active' : ''}`;
      pill.title = `${tab.name} (Tab ${index + 1}) - Click để xem, nhấp đúp để đổi tên`;

      let inner = '';
      if (cfg.showIndex) {
        inner += `<span class="tab-index-badge">${index + 1}</span>`;
      }
      inner += `<span class="tab-name-label">${tab.name}</span>`;
      if (cfg.showCloseBtn && DashboardState.present.tabs.length > 1) {
        inner += `<span class="tab-close-btn" data-del-tab="${tab.id}" title="Xóa tab">✕</span>`;
      }

      pill.innerHTML = inner;
      tabsList.appendChild(pill);

      // Tab click
      pill.addEventListener('click', async (e) => {
        if (e.target.dataset.delTab) {
          e.stopPropagation();
          const ok = await AppDialog.confirm(`Xóa tab "${tab.name}" và toàn bộ biểu đồ bên trong?`, 'Xóa Tab Dashboard', { isDanger: true, confirmText: 'Xóa Tab' });
          if (ok) {
            syncItemPositionsFromDOM();
            DashboardState.pushSnapshot();
            DashboardState.present.tabs = DashboardState.present.tabs.filter(t => t.id !== tab.id);
            DashboardState.present.items = DashboardState.present.items.filter(i => i.tabId !== tab.id);
            if (DashboardState.tabFilters) delete DashboardState.tabFilters[tab.id];
            const nextTabId = DashboardState.present.tabs[0]?.id;
            switchDashboardTab(nextTabId);
            DashboardState.isDirty = true;
          }
          return;
        }
        if (DashboardState.present.activeTabId !== tab.id) {
          switchDashboardTab(tab.id);
        }
      });

      // Tab double click to rename
      pill.addEventListener('dblclick', async (e) => {
        e.stopPropagation();
        const newName = await AppDialog.prompt('Đổi tên Tab:', tab.name, 'Đổi Tên Tab');
        if (newName && newName.trim()) {
          syncItemPositionsFromDOM();
          DashboardState.pushSnapshot();
          tab.name = newName.trim();
          DashboardState.isDirty = true;
          renderAllTabs();
          triggerDebouncedAutoSave();
          showToast('Đã đổi tên tab thành công!', 'success');
        }
      });
    });
  }

  // ===========================================================================
  // 11.45 FIGMA-STYLE RIGHT PROPERTIES SIDEBAR CONTROLLER
  // ===========================================================================
  const FigmaPropertiesSidebar = {
    sidebarEl: null,
    currentItemId: null,

    init() {
      this.sidebarEl = document.getElementById('figmaPropertiesSidebar');
      if (!this.sidebarEl) return;

      // Close button
      document.getElementById('fpsBtnClose')?.addEventListener('click', () => {
        this.close();
      });

      // 1. Position & Dimensions Inputs (X, Y, W, H)
      const propX = document.getElementById('fpsPropX');
      const propY = document.getElementById('fpsPropY');
      const propW = document.getElementById('fpsPropW');
      const propH = document.getElementById('fpsPropH');

      const onPosDimChange = () => {
        const item = this.getItem();
        if (!item) return;
        DashboardState.pushSnapshot();
        const rawX = parseInt(propX.value, 10) || 0;
        const rawY = parseInt(propY.value, 10) || 0;
        const rawW = parseInt(propW.value, 10) || 10;
        const rawH = parseInt(propH.value, 10) || 2;

        const minX = 24;
        const minY = 24;
        const w = Math.max(10, Math.min(1336 - minX, rawW));
        const maxX = Math.max(minX, 1360 - w - 24);
        const x = Math.max(minX, Math.min(maxX, rawX));
        const y = Math.max(minY, rawY);
        const h = Math.max(2, rawH);

        item.pos.x = x;
        item.pos.y = y;
        item.pos.w = w;
        item.pos.h = h;
        if (propX) propX.value = x;
        if (propY) propY.value = y;
        if (propW) propW.value = w;
        if (propH) propH.value = h;
        const card = document.getElementById(`item_${item.id}`);
        if (card) {
          card.style.left = `${x}px`;
          card.style.top = `${y}px`;
          card.style.width = `${w}px`;
          card.style.height = `${h}px`;
        }
      };
      [propX, propY, propW, propH].forEach(input => {
        input?.addEventListener('change', onPosDimChange);
      });

      // Presets 1:1, 16:9, 4:3
      document.getElementById('fpsBtnPresetSquare')?.addEventListener('click', () => {
        const item = this.getItem();
        if (!item) return;
        DashboardState.pushSnapshot();
        const size = Math.max(item.pos.w, item.pos.h);
        item.pos.w = size;
        item.pos.h = size;
        if (propW) propW.value = size;
        if (propH) propH.value = size;
        renderCanvas();
        this.syncUI(item);
      });
      document.getElementById('fpsBtnPreset169')?.addEventListener('click', () => {
        const item = this.getItem();
        if (!item) return;
        DashboardState.pushSnapshot();
        const newH = Math.round((item.pos.w * 9) / 16);
        item.pos.h = Math.max(20, newH);
        if (propH) propH.value = item.pos.h;
        renderCanvas();
        this.syncUI(item);
      });
      document.getElementById('fpsBtnPreset43')?.addEventListener('click', () => {
        const item = this.getItem();
        if (!item) return;
        DashboardState.pushSnapshot();
        const newH = Math.round((item.pos.w * 3) / 4);
        item.pos.h = Math.max(20, newH);
        if (propH) propH.value = item.pos.h;
        renderCanvas();
        this.syncUI(item);
      });

      // Z-Order & Lock
      document.getElementById('fpsBtnZUp')?.addEventListener('click', () => {
        const item = this.getItem();
        if (!item) return;
        DashboardState.pushSnapshot();
        item.pos.zIndex = (item.pos.zIndex || 10) + 2;
        renderCanvas();
      });
      document.getElementById('fpsBtnZDown')?.addEventListener('click', () => {
        const item = this.getItem();
        if (!item) return;
        DashboardState.pushSnapshot();
        item.pos.zIndex = Math.max(1, (item.pos.zIndex || 10) - 2);
        renderCanvas();
      });
      document.getElementById('fpsBtnToggleLock')?.addEventListener('click', () => {
        const item = this.getItem();
        if (!item) return;
        DashboardState.pushSnapshot();
        item.isLocked = !item.isLocked;
        const lockBtn = document.getElementById('fpsBtnToggleLock');
        if (lockBtn) lockBtn.textContent = item.isLocked ? '🔒' : '🔓';
        renderCanvas();
      });

      // 2. Fill & Color & Opacity
      const fillTypeSelect = document.getElementById('fpsFillType');
      const fillColorPicker = document.getElementById('fpsFillColorPicker');
      const fillColorHex = document.getElementById('fpsFillColorHex');
      const fillColorSwatch = document.getElementById('fpsFillColorSwatch');
      const opacitySlider = document.getElementById('fpsPropOpacity');
      const opacityVal = document.getElementById('fpsOpacityVal');

      const applyFill = (color) => {
        const item = this.getItem();
        if (!item) return;
        DashboardState.pushSnapshot();
        item.bgColor = color;
        if (fillColorHex) fillColorHex.value = color.toUpperCase();
        if (fillColorPicker) fillColorPicker.value = color.startsWith('#') && color.length === 7 ? color : '#ffffff';
        if (fillColorSwatch) fillColorSwatch.style.background = color;
        renderItemContent(item);
      };

      fillColorPicker?.addEventListener('input', (e) => applyFill(e.target.value));
      fillColorHex?.addEventListener('change', (e) => {
        let val = e.target.value.trim();
        if (!val.startsWith('#') && !val.startsWith('rgb')) val = '#' + val;
        applyFill(val);
      });

      fillTypeSelect?.addEventListener('change', (e) => {
        const mode = e.target.value;
        if (mode === 'transparent') {
          applyFill('transparent');
        } else if (mode === 'gradient') {
          applyFill('linear-gradient(135deg, #0284c7 0%, #38bdf8 100%)');
        } else {
          applyFill('#ffffff');
        }
      });

      document.querySelectorAll('#fpsSectionFill .fps-palette-dot').forEach(dot => {
        dot.addEventListener('click', () => {
          const fill = dot.dataset.fill;
          if (fill) applyFill(fill);
        });
      });

      opacitySlider?.addEventListener('input', (e) => {
        const val = parseInt(e.target.value, 10) || 100;
        if (opacityVal) opacityVal.textContent = `${val}%`;
        const item = this.getItem();
        if (!item) return;
        item.opacity = val;
        renderItemContent(item);
      });
      opacitySlider?.addEventListener('change', () => {
        DashboardState.pushSnapshot();
      });

      // 3. Stroke / Border
      const strokeColorPicker = document.getElementById('fpsStrokeColorPicker');
      const strokeColorHex = document.getElementById('fpsStrokeColorHex');
      const strokeColorSwatch = document.getElementById('fpsStrokeColorSwatch');
      const strokeWidth = document.getElementById('fpsStrokeWidth');
      const strokeStyle = document.getElementById('fpsStrokeStyle');

      const applyStroke = () => {
        const item = this.getItem();
        if (!item) return;
        DashboardState.pushSnapshot();
        item.borderColor = strokeColorPicker.value;
        item.borderWidth = parseInt(strokeWidth.value, 10) || 0;
        item.borderStyle = strokeStyle.value;
        if (strokeColorHex) strokeColorHex.value = item.borderColor.toUpperCase();
        if (strokeColorSwatch) strokeColorSwatch.style.background = item.borderColor;
        renderItemContent(item);
      };

      strokeColorPicker?.addEventListener('input', () => applyStroke());
      strokeColorHex?.addEventListener('change', (e) => {
        let val = e.target.value.trim();
        if (!val.startsWith('#') && !val.startsWith('rgb')) val = '#' + val;
        if (strokeColorPicker) strokeColorPicker.value = val;
        applyStroke();
      });
      strokeWidth?.addEventListener('input', () => applyStroke());
      strokeStyle?.addEventListener('change', () => applyStroke());

      // 4. Corner Radius
      const radiusSlider = document.getElementById('fpsPropRadius');
      const radiusVal = document.getElementById('fpsRadiusVal');
      const btnToggleIndividualRadius = document.getElementById('fpsBtnToggleIndividualRadius');
      const singleRadiusRow = document.getElementById('fpsSingleRadiusRow');
      const individualRadiusRow = document.getElementById('fpsIndividualRadiusRow');
      const radiusTL = document.getElementById('fpsRadiusTL');
      const radiusTR = document.getElementById('fpsRadiusTR');
      const radiusBR = document.getElementById('fpsRadiusBR');
      const radiusBL = document.getElementById('fpsRadiusBL');

      radiusSlider?.addEventListener('input', (e) => {
        const val = parseInt(e.target.value, 10) || 0;
        if (radiusVal) radiusVal.textContent = `${val}px`;
        const item = this.getItem();
        if (!item) return;
        item.borderRadius = val;
        delete item.radiusTL; delete item.radiusTR; delete item.radiusBR; delete item.radiusBL;
        renderItemContent(item);
      });
      radiusSlider?.addEventListener('change', () => DashboardState.pushSnapshot());

      btnToggleIndividualRadius?.addEventListener('click', () => {
        const isIndividual = individualRadiusRow.style.display !== 'none';
        if (isIndividual) {
          individualRadiusRow.style.display = 'none';
          singleRadiusRow.style.display = 'block';
          btnToggleIndividualRadius.classList.remove('active');
        } else {
          individualRadiusRow.style.display = 'block';
          singleRadiusRow.style.display = 'none';
          btnToggleIndividualRadius.classList.add('active');
        }
      });

      const applyIndividualRadius = () => {
        const item = this.getItem();
        if (!item) return;
        DashboardState.pushSnapshot();
        item.radiusTL = parseInt(radiusTL.value, 10) || 0;
        item.radiusTR = parseInt(radiusTR.value, 10) || 0;
        item.radiusBR = parseInt(radiusBR.value, 10) || 0;
        item.radiusBL = parseInt(radiusBL.value, 10) || 0;
        renderItemContent(item);
      };
      [radiusTL, radiusTR, radiusBR, radiusBL].forEach(inp => {
        inp?.addEventListener('input', applyIndividualRadius);
      });

      // 5. Line Special
      const lineThickness = document.getElementById('fpsLineThickness');
      const lineCap = document.getElementById('fpsLineCap');
      const lineOrientH = document.getElementById('fpsLineOrientH');
      const lineOrientV = document.getElementById('fpsLineOrientV');

      lineThickness?.addEventListener('input', (e) => {
        const item = this.getItem();
        if (!item) return;
        DashboardState.pushSnapshot();
        item.borderWidth = parseInt(e.target.value, 10) || 2;
        renderItemContent(item);
      });

      lineCap?.addEventListener('change', (e) => {
        const item = this.getItem();
        if (!item) return;
        DashboardState.pushSnapshot();
        item.lineCap = e.target.value;
        renderItemContent(item);
      });

      lineOrientH?.addEventListener('click', () => {
        const item = this.getItem();
        if (!item) return;
        DashboardState.pushSnapshot();
        item.lineOrientation = 'horizontal';
        lineOrientH.classList.add('active');
        lineOrientV.classList.remove('active');
        if (item.pos.w < item.pos.h) {
          const tmp = item.pos.w; item.pos.w = item.pos.h; item.pos.h = 24;
        }
        renderCanvas();
        this.syncUI(item);
      });

      lineOrientV?.addEventListener('click', () => {
        const item = this.getItem();
        if (!item) return;
        DashboardState.pushSnapshot();
        item.lineOrientation = 'vertical';
        lineOrientV.classList.add('active');
        lineOrientH.classList.remove('active');
        if (item.pos.h < item.pos.w) {
          const tmp = item.pos.h; item.pos.h = item.pos.w; item.pos.w = 24;
        }
        renderCanvas();
        this.syncUI(item);
      });

      // 6. Shadow
      const shadowToggle = document.getElementById('fpsShadowToggle');
      const shadowDetails = document.getElementById('fpsShadowDetails');
      const shadowBlur = document.getElementById('fpsShadowBlur');
      const shadowSpread = document.getElementById('fpsShadowSpread');

      shadowToggle?.addEventListener('change', (e) => {
        const item = this.getItem();
        if (!item) return;
        DashboardState.pushSnapshot();
        if (e.target.checked) {
          shadowDetails.style.display = 'block';
          item.boxShadow = '0 8px 20px rgba(15, 23, 42, 0.12)';
        } else {
          shadowDetails.style.display = 'none';
          item.boxShadow = 'none';
        }
        renderItemContent(item);
      });

      document.querySelectorAll('#fpsShadowDetails [data-shadow-preset]').forEach(btn => {
        btn.addEventListener('click', () => {
          const item = this.getItem();
          if (!item) return;
          DashboardState.pushSnapshot();
          const p = btn.dataset.shadowPreset;
          if (p === 'soft') item.boxShadow = '0 2px 8px rgba(15, 23, 42, 0.08)';
          else if (p === 'medium') item.boxShadow = '0 8px 20px rgba(15, 23, 42, 0.14)';
          else if (p === 'float') item.boxShadow = '0 16px 36px rgba(15, 23, 42, 0.22)';
          else if (p === 'glow') item.boxShadow = '0 0 16px rgba(2, 132, 199, 0.45)';
          renderItemContent(item);
        });
      });

      const applyCustomShadow = () => {
        const item = this.getItem();
        if (!item) return;
        const b = parseInt(shadowBlur.value, 10) || 12;
        const s = parseInt(shadowSpread.value, 10) || 0;
        item.boxShadow = `0 6px ${b}px ${s}px rgba(15, 23, 42, 0.15)`;
        renderItemContent(item);
      };
      shadowBlur?.addEventListener('input', applyCustomShadow);
      shadowSpread?.addEventListener('input', applyCustomShadow);

      // 7. Name & Header Toggle
      const propName = document.getElementById('fpsPropName');
      propName?.addEventListener('change', (e) => {
        const item = this.getItem();
        if (!item) return;
        DashboardState.pushSnapshot();
        item.name = e.target.value.trim() || item.type;
        const cardTitle = document.querySelector(`#item_${item.id} .db-item-title`);
        if (cardTitle) cardTitle.textContent = item.name;
      });

      document.getElementById('fpsHideHeaderToggle')?.addEventListener('change', (e) => {
        const item = this.getItem();
        if (!item) return;
        DashboardState.pushSnapshot();
        item.hideHeader = e.target.checked;
        renderCanvas();
      });

      // 8. Dedicated Component Section Listeners
      // --- Chart Controls ---
      document.getElementById('fpsChartSourceSelect')?.addEventListener('change', (e) => {
        const item = this.getItem();
        if (!item || item.type !== 'chart') return;
        DashboardState.pushSnapshot();
        item.dataSourceKey = e.target.value;
        renderCanvas();
        showToast(`Đã đổi nguồn dữ liệu sang "${e.target.value}"`, 'info');
      });

      document.querySelectorAll('#fpsSectionChart [data-chart-preset]').forEach(btn => {
        btn.addEventListener('click', () => {
          const item = this.getItem();
          if (!item || item.type !== 'chart') return;
          const preset = btn.dataset.chartPreset;
          if (ChartTemplates[preset]) {
            DashboardState.pushSnapshot();
            item.code = ChartTemplates[preset](item.dataSourceKey);
            renderItemContent(item);
            showToast(`Đã áp dụng mẫu ${btn.textContent.trim()}`, 'success');
          }
        });
      });

      document.getElementById('fpsBtnOpenChartStudio')?.addEventListener('click', () => {
        const item = this.getItem();
        if (item && item.type === 'chart') {
          ChartStudioModal.open(item.id);
        }
      });

      // --- KPI Controls ---
      document.getElementById('fpsKpiMetricSelect')?.addEventListener('change', (e) => {
        const item = this.getItem();
        if (!item || item.type !== 'kpi') return;
        DashboardState.pushSnapshot();
        item.metricColumn = e.target.value;
        renderItemContent(item);
      });

      document.getElementById('fpsKpiAggregation')?.addEventListener('change', (e) => {
        const item = this.getItem();
        if (!item || item.type !== 'kpi') return;
        DashboardState.pushSnapshot();
        item.aggregation = e.target.value;
        renderItemContent(item);
      });

      document.getElementById('fpsKpiPrefix')?.addEventListener('input', (e) => {
        const item = this.getItem();
        if (!item || item.type !== 'kpi') return;
        item.prefix = e.target.value;
        renderItemContent(item);
      });
      document.getElementById('fpsKpiPrefix')?.addEventListener('change', () => DashboardState.pushSnapshot());

      document.getElementById('fpsKpiSuffix')?.addEventListener('input', (e) => {
        const item = this.getItem();
        if (!item || item.type !== 'kpi') return;
        item.suffix = e.target.value;
        renderItemContent(item);
      });
      document.getElementById('fpsKpiSuffix')?.addEventListener('change', () => DashboardState.pushSnapshot());

      document.getElementById('fpsKpiFontSize')?.addEventListener('change', (e) => {
        const item = this.getItem();
        if (!item || item.type !== 'kpi') return;
        DashboardState.pushSnapshot();
        let val = e.target.value.trim();
        if (val && !val.endsWith('px')) val += 'px';
        item.fontSize = val;
        renderItemContent(item);
      });

      const applyKpiColor = (c) => {
        const item = this.getItem();
        if (!item || item.type !== 'kpi') return;
        DashboardState.pushSnapshot();
        item.color = c;
        const picker = document.getElementById('fpsKpiColorPicker');
        const hex = document.getElementById('fpsKpiColorHex');
        const swatch = document.getElementById('fpsKpiColorSwatch');
        if (picker && c.startsWith('#') && c.length === 7) picker.value = c;
        if (hex) hex.value = c.toUpperCase();
        if (swatch) swatch.style.background = c;
        renderItemContent(item);
      };
      document.getElementById('fpsKpiColorPicker')?.addEventListener('input', (e) => applyKpiColor(e.target.value));
      document.getElementById('fpsKpiColorHex')?.addEventListener('change', (e) => {
        let val = e.target.value.trim();
        if (!val.startsWith('#') && !val.startsWith('rgb')) val = '#' + val;
        applyKpiColor(val);
      });

      // --- Filter Controls ---
      document.getElementById('fpsFilterSourceSelect')?.addEventListener('change', (e) => {
        const item = this.getItem();
        if (!item || item.type !== 'filter') return;
        DashboardState.pushSnapshot();
        item.dataSourceKey = e.target.value;
        const cols = DataStore.getColumns(item.dataSourceKey);
        if (cols.length > 0 && !cols.includes(item.filterColumn)) {
          item.filterColumn = cols[0];
          if (!item.name || item.name.startsWith('Lọc: ')) {
            item.name = `Lọc: ${cols[0]}`;
          }
        }
        this.syncPropertiesPanel(item);
        renderCanvas();
      });

      document.getElementById('fpsFilterColumnSelect')?.addEventListener('change', (e) => {
        const item = this.getItem();
        if (!item || item.type !== 'filter') return;
        DashboardState.pushSnapshot();
        item.filterColumn = e.target.value;
        const rows = DataStore.getData(item.dataSourceKey, item.filterColumn);
        if (isProbableDateColumn(item.filterColumn, rows)) {
          item.filterType = 'date';
          item.operator = 'date_range';
        }
        if (!item.name || item.name === 'Bộ Lọc' || item.name.startsWith('Lọc: ') || item.name === item.filterColumn) {
          item.name = `Lọc: ${e.target.value}`;
          const cardTitle = document.querySelector(`#item_${item.id} .db-item-title`);
          if (cardTitle) cardTitle.textContent = item.name;
        }
        this.syncPropertiesPanel(item);
        renderCanvas();
      });

      document.getElementById('fpsFilterTypeSelect')?.addEventListener('change', (e) => {
        const item = this.getItem();
        if (!item || item.type !== 'filter') return;
        DashboardState.pushSnapshot();
        item.filterType = e.target.value;
        if (item.filterType === 'range') {
          item.operator = 'between';
        } else if (item.filterType === 'multiselect') {
          item.operator = 'in';
        } else if (item.filterType === 'date') {
          item.operator = 'date_range';
        } else if (item.filterType === 'text') {
          item.operator = 'contains';
        } else if (item.filterType === 'comparison') {
          item.operator = '>';
        } else {
          item.operator = '=';
        }
        this.syncPropertiesPanel(item);
        renderCanvas();
      });

      document.getElementById('fpsFilterParentSelect')?.addEventListener('change', (e) => {
        const item = this.getItem();
        if (!item || item.type !== 'filter') return;
        DashboardState.pushSnapshot();
        item.parentFilterId = e.target.value;
        renderCanvas();
      });

      document.getElementById('fpsFilterOpSelect')?.addEventListener('change', (e) => {
        const item = this.getItem();
        if (!item || item.type !== 'filter') return;
        DashboardState.pushSnapshot();
        item.operator = e.target.value;
        renderCanvas();
      });

      document.getElementById('fpsBtnOpenFilterModal')?.addEventListener('click', () => {
        const item = this.getItem();
        if (item && item.type === 'filter') {
          WidgetConfigModal.open('filter', item.id);
        }
      });

      // --- Text & Heading Controls ---
      document.getElementById('fpsTextContent')?.addEventListener('input', (e) => {
        const item = this.getItem();
        if (!item || (item.type !== 'text' && item.type !== 'heading')) return;
        item.text = e.target.value;
        const editableEl = document.querySelector(`#item_${item.id} .db-text-content-editable`);
        if (editableEl) editableEl.textContent = item.text;
      });
      document.getElementById('fpsTextContent')?.addEventListener('change', () => DashboardState.pushSnapshot());

      document.getElementById('fpsTextFontSize')?.addEventListener('change', (e) => {
        const item = this.getItem();
        if (!item || (item.type !== 'text' && item.type !== 'heading')) return;
        DashboardState.pushSnapshot();
        let val = e.target.value.trim();
        if (val && !val.endsWith('px')) val += 'px';
        item.fontSize = val;
        renderItemContent(item);
      });

      document.getElementById('fpsTextFontWeight')?.addEventListener('change', (e) => {
        const item = this.getItem();
        if (!item || (item.type !== 'text' && item.type !== 'heading')) return;
        DashboardState.pushSnapshot();
        item.fontWeight = e.target.value;
        renderItemContent(item);
      });

      const applyTextColor = (c) => {
        const item = this.getItem();
        if (!item || (item.type !== 'text' && item.type !== 'heading')) return;
        DashboardState.pushSnapshot();
        item.color = c;
        const picker = document.getElementById('fpsTextColorPicker');
        const hex = document.getElementById('fpsTextColorHex');
        const swatch = document.getElementById('fpsTextColorSwatch');
        if (picker && c.startsWith('#') && c.length === 7) picker.value = c;
        if (hex) hex.value = c.toUpperCase();
        if (swatch) swatch.style.background = c;
        renderItemContent(item);
      };
      document.getElementById('fpsTextColorPicker')?.addEventListener('input', (e) => applyTextColor(e.target.value));
      document.getElementById('fpsTextColorHex')?.addEventListener('change', (e) => {
        let val = e.target.value.trim();
        if (!val.startsWith('#') && !val.startsWith('rgb')) val = '#' + val;
        applyTextColor(val);
      });

      document.getElementById('fpsTextAlign')?.addEventListener('change', (e) => {
        const item = this.getItem();
        if (!item || (item.type !== 'text' && item.type !== 'heading')) return;
        DashboardState.pushSnapshot();
        item.textAlign = e.target.value;
        renderItemContent(item);
      });

      // --- Button Controls ---
      document.getElementById('fpsButtonLabel')?.addEventListener('input', (e) => {
        const item = this.getItem();
        if (!item || item.type !== 'button') return;
        item.name = e.target.value;
        renderItemContent(item);
      });
      document.getElementById('fpsButtonLabel')?.addEventListener('change', () => DashboardState.pushSnapshot());

      const applyBtnBg = (c) => {
        const item = this.getItem();
        if (!item || item.type !== 'button') return;
        DashboardState.pushSnapshot();
        item.btnBg = c;
        const picker = document.getElementById('fpsButtonBgPicker');
        const hex = document.getElementById('fpsButtonBgHex');
        const swatch = document.getElementById('fpsButtonBgSwatch');
        if (picker && c.startsWith('#') && c.length === 7) picker.value = c;
        if (hex) hex.value = c.toUpperCase();
        if (swatch) swatch.style.background = c;
        renderItemContent(item);
      };
      document.getElementById('fpsButtonBgPicker')?.addEventListener('input', (e) => applyBtnBg(e.target.value));
      document.getElementById('fpsButtonBgHex')?.addEventListener('change', (e) => {
        let val = e.target.value.trim();
        if (!val.startsWith('#') && !val.startsWith('rgb')) val = '#' + val;
        applyBtnBg(val);
      });

      const applyBtnColor = (c) => {
        const item = this.getItem();
        if (!item || item.type !== 'button') return;
        DashboardState.pushSnapshot();
        item.btnColor = c;
        const picker = document.getElementById('fpsButtonColorPicker');
        const hex = document.getElementById('fpsButtonColorHex');
        const swatch = document.getElementById('fpsButtonColorSwatch');
        if (picker && c.startsWith('#') && c.length === 7) picker.value = c;
        if (hex) hex.value = c.toUpperCase();
        if (swatch) swatch.style.background = c;
        renderItemContent(item);
      };
      document.getElementById('fpsButtonColorPicker')?.addEventListener('input', (e) => applyBtnColor(e.target.value));
      document.getElementById('fpsButtonColorHex')?.addEventListener('change', (e) => {
        let val = e.target.value.trim();
        if (!val.startsWith('#') && !val.startsWith('rgb')) val = '#' + val;
        applyBtnColor(val);
      });

      document.getElementById('fpsButtonRadius')?.addEventListener('change', (e) => {
        const item = this.getItem();
        if (!item || item.type !== 'button') return;
        DashboardState.pushSnapshot();
        item.borderRadius = parseInt(e.target.value, 10) || 0;
        renderItemContent(item);
      });

      // --- Image Controls ---
      document.getElementById('fpsImageUrl')?.addEventListener('change', (e) => {
        const item = this.getItem();
        if (!item || item.type !== 'image') return;
        DashboardState.pushSnapshot();
        item.imageUrl = e.target.value.trim();
        renderItemContent(item);
      });

      document.getElementById('fpsBtnUploadImage')?.addEventListener('click', () => {
        const item = this.getItem();
        if (!item || item.type !== 'image') return;
        const fileInput = document.createElement('input');
        fileInput.type = 'file';
        fileInput.accept = 'image/*';
        fileInput.onchange = (e) => {
          const file = e.target.files[0];
          if (!file) return;
          const reader = new FileReader();
          reader.onload = (evt) => {
            DashboardState.pushSnapshot();
            item.imageUrl = evt.target.result;
            const urlInput = document.getElementById('fpsImageUrl');
            if (urlInput) urlInput.value = item.imageUrl.slice(0, 100) + '...';
            renderItemContent(item);
            showToast('✓ Đã tải và nhúng hình ảnh vào Dashboard!', 'success');
          };
          reader.readAsDataURL(file);
        };
        fileInput.click();
      });

      document.getElementById('fpsImageFit')?.addEventListener('change', (e) => {
        const item = this.getItem();
        if (!item || item.type !== 'image') return;
        DashboardState.pushSnapshot();
        item.fit = e.target.value;
        renderItemContent(item);
      });

      document.getElementById('fpsImageRadius')?.addEventListener('change', (e) => {
        const item = this.getItem();
        if (!item || item.type !== 'image') return;
        DashboardState.pushSnapshot();
        item.borderRadius = parseInt(e.target.value, 10) || 0;
        renderItemContent(item);
      });

      document.getElementById('fpsBtnDuplicate')?.addEventListener('click', () => {
        const orig = this.getItem();
        if (!orig) return;
        DashboardState.pushSnapshot();
        const dupId = orig.type + '_' + Date.now();
        const copy = {
          ...JSON.parse(JSON.stringify(orig)),
          id: dupId,
          name: `${orig.name} (Bản sao)`,
          pos: { ...orig.pos, x: orig.pos.x + 30, y: orig.pos.y + 30 }
        };
        DashboardState.present.items.push(copy);
        renderCanvas();
        this.open(dupId);
        showToast('Đã nhân bản thành phần!', 'success');
      });

      document.getElementById('fpsBtnDelete')?.addEventListener('click', async () => {
        const item = this.getItem();
        if (!item) return;
        const ok = await AppDialog.confirm(`Xóa thành phần "${item.name}"?`, 'Xác Nhận Xóa', { isDanger: true, confirmText: 'Xóa' });
        if (ok) {
          DashboardState.pushSnapshot();
          DashboardState.present.items = DashboardState.present.items.filter(i => i.id !== item.id);
          DashboardState.selectedItemId = null;
          this.close();
          renderCanvas();
          showToast('Đã xóa thành phần', 'info');
        }
      });
    },

    getItem() {
      if (!this.currentItemId) return null;
      return DashboardState.present.items.find(i => i.id === this.currentItemId) || null;
    },

    open(itemId) {
      this.currentItemId = itemId;
      const item = this.getItem();
      if (!item || !this.sidebarEl) return;

      // Đóng TabConfigSidebar nếu đang mở để tránh chồng lấn
      if (typeof TabConfigSidebar !== 'undefined' && TabConfigSidebar.close) {
        TabConfigSidebar.close();
      }

      this.syncUI(item);
      this.sidebarEl.style.display = 'flex';
      this.sidebarEl.classList.add('open');
      document.body.classList.add('has-right-sidebar');
      updateCanvasTransform();
    },

    openTabConfig() {
      if (typeof TabConfigSidebar !== 'undefined') {
        TabConfigSidebar.open();
      }
    },

    close() {
      if (!this.sidebarEl) return;
      this.sidebarEl.style.display = 'none';
      this.sidebarEl.classList.remove('open');
      this.currentItemId = null;
      if (typeof TabConfigSidebar === 'undefined' || !TabConfigSidebar.isOpen()) {
        document.body.classList.remove('has-right-sidebar');
      }
      updateCanvasTransform();
    },

    isOpen() {
      return Boolean(this.sidebarEl && this.sidebarEl.classList.contains('open'));
    },

    syncTabUI() {
      if (typeof TabConfigSidebar !== 'undefined') {
        TabConfigSidebar.syncTabUI();
      }
    },

    syncPropertiesPanel(item) {
      this.syncUI(item);
    },

    syncUI(item) {
      if (!item) return;

      const isChart = item.type === 'chart';
      const isKpi = item.type === 'kpi';
      const isFilter = item.type === 'filter';
      const isHeading = item.type === 'heading';
      const isText = item.type === 'text';
      const isButton = item.type === 'button';
      const isImage = item.type === 'image';
      const isShape = item.type === 'shape';
      const isFrame = item.type === 'frame';

      const isLine = isShape && item.shapeType === 'line';
      const isRect = isShape && (item.shapeType === 'rectangle' || !item.shapeType);
      const isCircle = isShape && item.shapeType === 'circle';

      // Header info
      const typeIcon = document.getElementById('fpsTypeIcon');
      const title = document.getElementById('fpsTitle');
      const subtitle = document.getElementById('fpsSubtitle');

      if (typeIcon) {
        if (isChart) typeIcon.textContent = '📊';
        else if (isKpi) typeIcon.textContent = '📈';
        else if (isFilter) typeIcon.textContent = '🔍';
        else if (isHeading) typeIcon.textContent = '🏷️';
        else if (isText) typeIcon.textContent = '📝';
        else if (isButton) typeIcon.textContent = '🔘';
        else if (isImage) typeIcon.textContent = '🖼️';
        else if (isFrame) typeIcon.textContent = '🔲';
        else if (isLine) typeIcon.textContent = '➖';
        else if (isCircle) typeIcon.textContent = '⚪';
        else typeIcon.textContent = '⬜';
      }

      if (title) {
        title.textContent = item.name || (
          isChart ? 'Biểu Đồ ECharts' :
          isKpi ? 'Chỉ Số KPI' :
          isFilter ? 'Bộ Lọc' :
          isHeading ? 'Tiêu Đề' :
          isText ? 'Văn Bản' :
          isButton ? 'Nút Bấm' :
          isImage ? 'Hình Ảnh' :
          isFrame ? 'Khung Phân Vùng' : 'Hình Khối'
        );
      }

      if (subtitle) {
        if (isChart) subtitle.textContent = 'ECharts Code & Nguồn Dữ Liệu';
        else if (isKpi) subtitle.textContent = 'Chỉ Số Tổng Hợp Tự Động';
        else if (isFilter) subtitle.textContent = 'Bộ Lọc Tương Tác Dữ Liệu';
        else if (isHeading) subtitle.textContent = 'Tiêu Đề Báo Cáo / Khu Vực';
        else if (isText) subtitle.textContent = 'Đoạn Ghi Chú & Thuyết Minh';
        else if (isButton) subtitle.textContent = 'Nút Bấm & Hành Động';
        else if (isImage) subtitle.textContent = 'Ảnh Minh Họa & Logo';
        else if (isFrame) subtitle.textContent = 'Khung Phân Vùng Giao Diện';
        else if (isLine) subtitle.textContent = 'Đường Kẻ Phân Cách (Line)';
        else if (isCircle) subtitle.textContent = 'Hình Tròn (Circle)';
        else subtitle.textContent = 'Hình Chữ Nhật / Vuông';
      }

      // 0. General: Name & Hide Header Toggle
      const propName = document.getElementById('fpsPropName');
      if (propName) propName.value = item.name || '';
      const hideHeaderToggle = document.getElementById('fpsHideHeaderToggle');
      if (hideHeaderToggle) hideHeaderToggle.checked = Boolean(item.hideHeader);

      // 1. Layout inputs (X, Y, W, H)
      const propX = document.getElementById('fpsPropX');
      const propY = document.getElementById('fpsPropY');
      const propW = document.getElementById('fpsPropW');
      const propH = document.getElementById('fpsPropH');
      if (propX) propX.value = item.pos.x;
      if (propY) propY.value = item.pos.y;
      if (propW) propW.value = item.pos.w;
      if (propH) propH.value = item.pos.h;

      const lockBtn = document.getElementById('fpsBtnToggleLock');
      if (lockBtn) lockBtn.textContent = item.isLocked ? '🔒' : '🔓';

      // Section visibility toggling
      const secChart = document.getElementById('fpsSectionChart');
      if (secChart) secChart.style.display = isChart ? 'block' : 'none';

      const secKpi = document.getElementById('fpsSectionKpi');
      if (secKpi) secKpi.style.display = isKpi ? 'block' : 'none';

      const secFilter = document.getElementById('fpsSectionFilter');
      if (secFilter) secFilter.style.display = isFilter ? 'block' : 'none';

      const secText = document.getElementById('fpsSectionText');
      if (secText) secText.style.display = (isText || isHeading) ? 'block' : 'none';

      const secBtn = document.getElementById('fpsSectionButton');
      if (secBtn) secBtn.style.display = isButton ? 'block' : 'none';

      const secImg = document.getElementById('fpsSectionImage');
      if (secImg) secImg.style.display = isImage ? 'block' : 'none';

      const secFill = document.getElementById('fpsSectionFill');
      if (secFill) secFill.style.display = ((isShape && !isLine) || isFrame) ? 'flex' : 'none';

      const secStroke = document.getElementById('fpsSectionStroke');
      if (secStroke) secStroke.style.display = (isShape || isFrame) ? 'block' : 'none';

      const secRadius = document.getElementById('fpsSectionRadius');
      if (secRadius) secRadius.style.display = ((isShape && !isLine && !isCircle) || isFrame) ? 'flex' : 'none';

      const secLine = document.getElementById('fpsSectionLine');
      if (secLine) secLine.style.display = isLine ? 'flex' : 'none';

      const secShadow = document.getElementById('fpsSectionShadow');
      if (secShadow) secShadow.style.display = (isShape || isFrame || isChart || isKpi) ? 'block' : 'none';

      // Synchronize Chart section
      if (isChart) {
        const srcSelect = document.getElementById('fpsChartSourceSelect');
        if (srcSelect) {
          const sources = DataStore.getAvailableSources();
          srcSelect.innerHTML = sources.map(s => `<option value="${escapeHtml(s.key)}">${escapeHtml(s.label)}</option>`).join('');
          if (item.dataSourceKey) srcSelect.value = item.dataSourceKey;
        }
      }

      // Synchronize KPI section
      if (isKpi) {
        const kpiSelect = document.getElementById('fpsKpiMetricSelect');
        if (kpiSelect) {
          const cols = DataStore.getColumns(item.dataSourceKey || Object.keys(DataStore.files)[0]);
          kpiSelect.innerHTML = cols.map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');
          if (item.metricColumn) kpiSelect.value = item.metricColumn;
        }
        const kpiAgg = document.getElementById('fpsKpiAggregation');
        if (kpiAgg) kpiAgg.value = item.aggregation || 'sum';
        const kpiPre = document.getElementById('fpsKpiPrefix');
        if (kpiPre) kpiPre.value = item.prefix || '';
        const kpiSuf = document.getElementById('fpsKpiSuffix');
        if (kpiSuf) kpiSuf.value = item.suffix || '';
        const kpiFs = document.getElementById('fpsKpiFontSize');
        if (kpiFs) kpiFs.value = item.fontSize || '30px';
        const kpiC = item.color || '#0f172a';
        const kpiPicker = document.getElementById('fpsKpiColorPicker');
        const kpiHex = document.getElementById('fpsKpiColorHex');
        const kpiSwatch = document.getElementById('fpsKpiColorSwatch');
        if (kpiPicker && kpiC.startsWith('#') && kpiC.length === 7) kpiPicker.value = kpiC;
        if (kpiHex) kpiHex.value = kpiC.toUpperCase();
        if (kpiSwatch) kpiSwatch.style.background = kpiC;
      }

      // Synchronize Filter section
      if (isFilter) {
        const sourceSelect = document.getElementById('fpsFilterSourceSelect');
        if (sourceSelect) {
          const sources = DataStore.getAvailableSources();
          sourceSelect.innerHTML = sources.map(s => `<option value="${escapeHtml(s.key)}" ${s.key === item.dataSourceKey ? 'selected' : ''}>📄 ${escapeHtml(s.label)}</option>`).join('');
          if (item.dataSourceKey) {
            sourceSelect.value = item.dataSourceKey;
          } else if (sources.length > 0) {
            const matched = sources.find(s => DataStore.getColumns(s.key).includes(item.filterColumn));
            item.dataSourceKey = matched ? matched.key : sources[0].key;
            sourceSelect.value = item.dataSourceKey;
          }
        }

        const filterSelect = document.getElementById('fpsFilterColumnSelect');
        if (filterSelect) {
          const cols = DataStore.getColumns(item.dataSourceKey);
          filterSelect.innerHTML = cols.map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');
          if (item.filterColumn && cols.includes(item.filterColumn)) {
            filterSelect.value = item.filterColumn;
          } else if (cols.length > 0) {
            item.filterColumn = cols[0];
            filterSelect.value = cols[0];
          }
        }

        const typeSelect = document.getElementById('fpsFilterTypeSelect');
        if (typeSelect) {
          typeSelect.value = item.filterType || 'select';
        }

        const parentSelect = document.getElementById('fpsFilterParentSelect');
        if (parentSelect) {
          const otherFilters = DashboardState.present.items.filter(i => i.type === 'filter' && i.id !== item.id);
          parentSelect.innerHTML = '<option value="">-- Không có (Bộ lọc độc lập) --</option>' +
            otherFilters.map(f => `<option value="${f.id}" ${f.id === item.parentFilterId ? 'selected' : ''}>🔗 ${escapeHtml(f.name || f.filterColumn)} (Cột: ${escapeHtml(f.filterColumn)})</option>`).join('');
        }

        const opSelect = document.getElementById('fpsFilterOpSelect');
        if (opSelect) {
          const fType = item.filterType || 'range';
          const curOp = item.operator || (fType === 'range' ? 'between' : (fType === 'multiselect' ? 'in' : (fType === 'date' ? 'date_range' : (fType === 'text' ? 'contains' : '='))));
          if (fType === 'range') {
            opSelect.innerHTML = `
              <option value="between" ${curOp === 'between' ? 'selected' : ''}>↔️ Trong Khoảng [Từ ... Đến ...]</option>
            `;
          } else if (fType === 'comparison') {
            opSelect.innerHTML = `
              <option value="=" ${curOp === '=' ? 'selected' : ''}>= Bằng</option>
              <option value=">" ${curOp === '>' ? 'selected' : ''}>&gt; Lớn hơn</option>
              <option value=">=" ${curOp === '>=' ? 'selected' : ''}>&ge; Lớn hơn hoặc bằng</option>
              <option value="<" ${curOp === '<' ? 'selected' : ''}>&lt; Bé hơn</option>
              <option value="<=" ${curOp === '<=' ? 'selected' : ''}>&le; Bé hơn hoặc bằng</option>
              <option value="!=" ${curOp === '!=' ? 'selected' : ''}>≠ Khác</option>
            `;
          } else if (fType === 'multiselect') {
            opSelect.innerHTML = `
              <option value="in" ${curOp === 'in' ? 'selected' : ''}>☑️ Thuộc danh sách (Chọn nhiều mục)</option>
            `;
          } else if (fType === 'date') {
            opSelect.innerHTML = `
              <option value="date_range" ${curOp === 'date_range' ? 'selected' : ''}>📅 Khoảng ngày (Date Range)</option>
              <option value="date_eq" ${curOp === 'date_eq' ? 'selected' : ''}>🎯 Đúng ngày (=)</option>
              <option value="date_gte" ${curOp === 'date_gte' ? 'selected' : ''}>⏩ Từ ngày trở đi (&ge;)</option>
              <option value="date_lte" ${curOp === 'date_lte' ? 'selected' : ''}>⏪ Đến ngày trở về trước (&le;)</option>
            `;
          } else if (fType === 'text') {
            opSelect.innerHTML = `
              <option value="contains" ${curOp === 'contains' ? 'selected' : ''}>Chứa từ khóa</option>
              <option value="starts_with" ${curOp === 'starts_with' ? 'selected' : ''}>Bắt đầu bằng</option>
              <option value="ends_with" ${curOp === 'ends_with' ? 'selected' : ''}>Kết thúc bằng</option>
              <option value="not_contains" ${curOp === 'not_contains' ? 'selected' : ''}>Không chứa từ khóa</option>
              <option value="=" ${curOp === '=' ? 'selected' : ''}>= Chính xác</option>
            `;
          } else {
            opSelect.innerHTML = `
              <option value="=" ${curOp === '=' ? 'selected' : ''}>= Bằng (Chọn 1 từ danh sách)</option>
            `;
          }
        }
      }

      // Synchronize Text / Heading section
      if (isText || isHeading) {
        const txtContent = document.getElementById('fpsTextContent');
        if (txtContent) txtContent.value = item.text !== undefined ? item.text : (item.name || '');
        const txtFs = document.getElementById('fpsTextFontSize');
        if (txtFs) txtFs.value = item.fontSize || (isHeading ? '24px' : '14px');
        const txtFw = document.getElementById('fpsTextFontWeight');
        if (txtFw) txtFw.value = item.fontWeight || (isHeading ? '700' : 'normal');
        const txtAlign = document.getElementById('fpsTextAlign');
        if (txtAlign) txtAlign.value = item.textAlign || 'left';
        const txtC = item.color || '#0f172a';
        const txtPicker = document.getElementById('fpsTextColorPicker');
        const txtHex = document.getElementById('fpsTextColorHex');
        const txtSwatch = document.getElementById('fpsTextColorSwatch');
        if (txtPicker && txtC.startsWith('#') && txtC.length === 7) txtPicker.value = txtC;
        if (txtHex) txtHex.value = txtC.toUpperCase();
        if (txtSwatch) txtSwatch.style.background = txtC;
      }

      // Synchronize Button section
      if (isButton) {
        const btnLbl = document.getElementById('fpsButtonLabel');
        if (btnLbl) btnLbl.value = item.name || 'Nút Bấm';
        const btnBg = item.btnBg || '#004ea8';
        const btnBgPicker = document.getElementById('fpsButtonBgPicker');
        const btnBgHex = document.getElementById('fpsButtonBgHex');
        const btnBgSwatch = document.getElementById('fpsButtonBgSwatch');
        if (btnBgPicker && btnBg.startsWith('#') && btnBg.length === 7) btnBgPicker.value = btnBg;
        if (btnBgHex) btnBgHex.value = btnBg.toUpperCase();
        if (btnBgSwatch) btnBgSwatch.style.background = btnBg;

        const btnColor = item.btnColor || '#ffffff';
        const btnColorPicker = document.getElementById('fpsButtonColorPicker');
        const btnColorHex = document.getElementById('fpsButtonColorHex');
        const btnColorSwatch = document.getElementById('fpsButtonColorSwatch');
        if (btnColorPicker && btnColor.startsWith('#') && btnColor.length === 7) btnColorPicker.value = btnColor;
        if (btnColorHex) btnColorHex.value = btnColor.toUpperCase();
        if (btnColorSwatch) btnColorSwatch.style.background = btnColor;

        const btnRad = document.getElementById('fpsButtonRadius');
        if (btnRad) btnRad.value = item.borderRadius !== undefined ? item.borderRadius : 8;
      }

      // Synchronize Image section
      if (isImage) {
        const imgUrl = document.getElementById('fpsImageUrl');
        if (imgUrl) imgUrl.value = item.imageUrl || '';
        const imgFit = document.getElementById('fpsImageFit');
        if (imgFit) imgFit.value = item.fit || 'contain';
        const imgRad = document.getElementById('fpsImageRadius');
        if (imgRad) imgRad.value = item.borderRadius !== undefined ? item.borderRadius : 6;
      }

      // Synchronize Fill & Opacity
      const bg = item.bgColor || '#ffffff';
      const fillColorPicker = document.getElementById('fpsFillColorPicker');
      const fillColorHex = document.getElementById('fpsFillColorHex');
      const fillColorSwatch = document.getElementById('fpsFillColorSwatch');
      const fillTypeSelect = document.getElementById('fpsFillType');

      if (fillTypeSelect) {
        if (bg === 'transparent') fillTypeSelect.value = 'transparent';
        else if (bg.includes('gradient')) fillTypeSelect.value = 'gradient';
        else fillTypeSelect.value = 'solid';
      }
      if (fillColorHex) fillColorHex.value = bg;
      if (fillColorPicker) fillColorPicker.value = bg.startsWith('#') && bg.length === 7 ? bg : '#ffffff';
      if (fillColorSwatch) fillColorSwatch.style.background = bg;

      const opacity = item.opacity !== undefined ? item.opacity : 100;
      const opacitySlider = document.getElementById('fpsPropOpacity');
      const opacityVal = document.getElementById('fpsOpacityVal');
      if (opacitySlider) opacitySlider.value = opacity;
      if (opacityVal) opacityVal.textContent = `${opacity}%`;

      // Synchronize Stroke / Border
      const strokeColorPicker = document.getElementById('fpsStrokeColorPicker');
      const strokeColorHex = document.getElementById('fpsStrokeColorHex');
      const strokeColorSwatch = document.getElementById('fpsStrokeColorSwatch');
      const strokeWidth = document.getElementById('fpsStrokeWidth');
      const strokeStyle = document.getElementById('fpsStrokeStyle');

      const strokeC = item.borderColor || '#cbd5e1';
      if (strokeColorHex) strokeColorHex.value = strokeC;
      if (strokeColorPicker) strokeColorPicker.value = strokeC.startsWith('#') && strokeC.length === 7 ? strokeC : '#cbd5e1';
      if (strokeColorSwatch) strokeColorSwatch.style.background = strokeC;
      if (strokeWidth) strokeWidth.value = item.borderWidth !== undefined ? item.borderWidth : (isLine ? 2 : 1);
      if (strokeStyle) strokeStyle.value = item.borderStyle || 'solid';

      // Synchronize Corner Radius
      const radius = item.borderRadius !== undefined ? item.borderRadius : 8;
      const radiusSlider = document.getElementById('fpsPropRadius');
      const radiusVal = document.getElementById('fpsRadiusVal');
      if (radiusSlider) radiusSlider.value = radius;
      if (radiusVal) radiusVal.textContent = `${radius}px`;

      const singleRadiusRow = document.getElementById('fpsSingleRadiusRow');
      const individualRadiusRow = document.getElementById('fpsIndividualRadiusRow');
      const btnToggleIndividualRadius = document.getElementById('fpsBtnToggleIndividualRadius');
      const isInd = item.radiusTL !== undefined;
      if (singleRadiusRow) singleRadiusRow.style.display = isInd ? 'none' : 'block';
      if (individualRadiusRow) individualRadiusRow.style.display = isInd ? 'block' : 'none';
      if (btnToggleIndividualRadius) btnToggleIndividualRadius.classList.toggle('active', isInd);

      const radiusTL = document.getElementById('fpsRadiusTL');
      const radiusTR = document.getElementById('fpsRadiusTR');
      const radiusBR = document.getElementById('fpsRadiusBR');
      const radiusBL = document.getElementById('fpsRadiusBL');
      if (radiusTL) radiusTL.value = item.radiusTL || radius;
      if (radiusTR) radiusTR.value = item.radiusTR || radius;
      if (radiusBR) radiusBR.value = item.radiusBR || radius;
      if (radiusBL) radiusBL.value = item.radiusBL || radius;

      // Synchronize Line Special
      const lineThickness = document.getElementById('fpsLineThickness');
      const lineCap = document.getElementById('fpsLineCap');
      const lineOrientH = document.getElementById('fpsLineOrientH');
      const lineOrientV = document.getElementById('fpsLineOrientV');
      if (lineThickness) lineThickness.value = item.borderWidth || 2;
      if (lineCap) lineCap.value = item.lineCap || 'round';
      const isVert = item.lineOrientation === 'vertical';
      if (lineOrientH) lineOrientH.classList.toggle('active', !isVert);
      if (lineOrientV) lineOrientV.classList.toggle('active', isVert);

      // Synchronize Shadow
      const shadowToggle = document.getElementById('fpsShadowToggle');
      const shadowDetails = document.getElementById('fpsShadowDetails');
      const hasShadow = item.boxShadow && item.boxShadow !== 'none';
      if (shadowToggle) shadowToggle.checked = Boolean(hasShadow);
      if (shadowDetails) shadowDetails.style.display = hasShadow ? 'block' : 'none';
    },

  };

  // ===========================================================================
  // 11.46 DEDICATED TAB CONFIGURATION SIDEBAR CONTROLLER (ĐỘC LẬP VỚI THUỘC TÍNH THÀNH PHẦN)
  // ===========================================================================
  const TabConfigSidebar = {
    sidebarEl: null,

    init() {
      this.sidebarEl = document.getElementById('tabConfigSidebar');
      if (!this.sidebarEl) return;

      // Close button
      document.getElementById('tabSidebarBtnClose')?.addEventListener('click', () => {
        this.close();
      });

      // 1. Position 4-Direction Grid Cards
      document.querySelectorAll('#fpsTabPosGrid .fps-pos-card').forEach(card => {
        card.addEventListener('click', () => {
          const pos = card.dataset.pos;
          if (!pos) return;
          DashboardState.pushSnapshot();
          const cfg = TabBarManager.getConfig();
          cfg.position = pos;
          TabBarManager.saveConfig(cfg);
          TabBarManager.apply(cfg);
          this.syncTabUI();
          showToast(`Đã đổi thanh Tab sang: ${pos === 'top' ? 'Trên Cùng' : (pos === 'bottom' ? 'Dưới Cùng' : (pos === 'left' ? 'Cạnh Trái' : 'Cạnh Phải'))}`, 'info');
        });
      });

      // 2. Collapse Segment Toggles
      document.querySelectorAll('#fpsTabCollapseSegment button').forEach(btn => {
        btn.addEventListener('click', () => {
          const state = btn.dataset.state;
          const isCollapsed = (state === 'collapsed');
          DashboardState.pushSnapshot();
          const cfg = TabBarManager.getConfig();
          cfg.collapsed = isCollapsed;
          TabBarManager.saveConfig(cfg);
          TabBarManager.apply(cfg);
          this.syncTabUI();
        });
      });

      // Vertical width slider
      const verticalWidthSlider = document.getElementById('fpsTabVerticalWidth');
      const verticalWidthVal = document.getElementById('fpsTabWidthVal');
      verticalWidthSlider?.addEventListener('input', (e) => {
        const val = parseInt(e.target.value, 10) || 230;
        if (verticalWidthVal) verticalWidthVal.textContent = `${val}px`;
        const cfg = TabBarManager.getConfig();
        cfg.verticalWidth = val;
        TabBarManager.apply(cfg);
      });
      verticalWidthSlider?.addEventListener('change', (e) => {
        const val = parseInt(e.target.value, 10) || 230;
        DashboardState.pushSnapshot();
        const cfg = TabBarManager.getConfig();
        cfg.verticalWidth = val;
        TabBarManager.saveConfig(cfg);
        TabBarManager.apply(cfg);
      });

      // 3. Style & Size
      document.getElementById('fpsTabStyleSelect')?.addEventListener('change', (e) => {
        DashboardState.pushSnapshot();
        const cfg = TabBarManager.getConfig();
        cfg.style = e.target.value;
        TabBarManager.saveConfig(cfg);
        TabBarManager.apply(cfg);
      });
      document.getElementById('fpsTabSizeSelect')?.addEventListener('change', (e) => {
        DashboardState.pushSnapshot();
        const cfg = TabBarManager.getConfig();
        cfg.size = e.target.value;
        TabBarManager.saveConfig(cfg);
        TabBarManager.apply(cfg);
      });

      // 4. Accent Color
      const tabAccentPicker = document.getElementById('fpsTabAccentColorPicker');
      const tabAccentHex = document.getElementById('fpsTabAccentColorHex');
      const tabAccentSwatch = document.getElementById('fpsTabAccentColorSwatch');

      const applyTabAccent = (color) => {
        DashboardState.pushSnapshot();
        const cfg = TabBarManager.getConfig();
        cfg.accentColor = color;
        TabBarManager.saveConfig(cfg);
        TabBarManager.apply(cfg);
        if (tabAccentSwatch) tabAccentSwatch.style.background = color;
        if (tabAccentHex) tabAccentHex.value = color.toUpperCase();
        if (tabAccentPicker && color.startsWith('#') && color.length === 7) tabAccentPicker.value = color;
        document.querySelectorAll('#fpsTabPaletteRow .fps-palette-dot').forEach(d => {
          d.classList.toggle('active', d.dataset.color.toLowerCase() === color.toLowerCase());
        });
      };

      tabAccentPicker?.addEventListener('input', (e) => applyTabAccent(e.target.value));
      tabAccentHex?.addEventListener('change', (e) => {
        let val = e.target.value.trim();
        if (!val.startsWith('#')) val = '#' + val;
        applyTabAccent(val);
      });
      document.querySelectorAll('#fpsTabPaletteRow .fps-palette-dot').forEach(dot => {
        dot.addEventListener('click', () => {
          applyTabAccent(dot.dataset.color);
        });
      });

      // 5. Toggles (Index & Close)
      document.getElementById('fpsTabShowIndexToggle')?.addEventListener('change', (e) => {
        DashboardState.pushSnapshot();
        const cfg = TabBarManager.getConfig();
        cfg.showIndex = e.target.checked;
        TabBarManager.saveConfig(cfg);
        TabBarManager.apply(cfg);
      });
      document.getElementById('fpsTabShowCloseToggle')?.addEventListener('change', (e) => {
        DashboardState.pushSnapshot();
        const cfg = TabBarManager.getConfig();
        cfg.showCloseBtn = e.target.checked;
        TabBarManager.saveConfig(cfg);
        TabBarManager.apply(cfg);
      });

      // 6. Add new tab in Sidebar
      document.getElementById('fpsBtnAddNewTab')?.addEventListener('click', () => {
        syncItemPositionsFromDOM();
        DashboardState.pushSnapshot();
        const newId = 'tab_' + Date.now();
        const newName = `Trang ${DashboardState.present.tabs.length + 1}`;
        DashboardState.present.tabs.push({ id: newId, name: newName });
        switchDashboardTab(newId);
        DashboardState.isDirty = true;
        this.syncTabUI();
        showToast(`Đã tạo "${newName}"`, 'success');
      });

      // 7. Reset Tab Config
      document.getElementById('fpsBtnResetTabConfig')?.addEventListener('click', () => {
        TabBarManager.resetDefaults();
        this.syncTabUI();
      });
    },

    open() {
      if (!this.sidebarEl) this.init();
      if (!this.sidebarEl) return;
      // Đóng bảng thuộc tính thành phần để tránh đè lên nhau
      if (typeof FigmaPropertiesSidebar !== 'undefined' && FigmaPropertiesSidebar.close) {
        FigmaPropertiesSidebar.close();
      }
      this.syncTabUI();
      this.sidebarEl.style.display = 'flex';
      this.sidebarEl.classList.add('open');
      document.body.classList.add('has-right-sidebar');
      updateCanvasTransform();
    },

    close() {
      if (!this.sidebarEl) return;
      this.sidebarEl.style.display = 'none';
      this.sidebarEl.classList.remove('open');
      if (typeof FigmaPropertiesSidebar === 'undefined' || !FigmaPropertiesSidebar.isOpen()) {
        document.body.classList.remove('has-right-sidebar');
      }
      updateCanvasTransform();
    },

    isOpen() {
      return Boolean(this.sidebarEl && this.sidebarEl.classList.contains('open'));
    },

    syncTabUI() {
      const cfg = TabBarManager.getConfig();

      // 1. Position Cards
      const pos = cfg.position || 'top';
      document.querySelectorAll('#fpsTabPosGrid .fps-pos-card').forEach(c => {
        c.classList.toggle('active', c.dataset.pos === pos);
      });
      const posBadge = document.getElementById('fpsCurrentPosBadge');
      if (posBadge) {
        posBadge.textContent = pos === 'top' ? 'Trên Cùng' : (pos === 'bottom' ? 'Dưới Cùng' : (pos === 'left' ? 'Cạnh Trái' : 'Cạnh Phải'));
      }

      // 2. Collapse & Vertical Width
      const isCollapsed = Boolean(cfg.collapsed);
      document.querySelectorAll('#fpsTabCollapseSegment button').forEach(b => {
        b.classList.toggle('active', (b.dataset.state === 'collapsed') === isCollapsed);
      });

      const vertRow = document.getElementById('fpsTabVerticalWidthRow');
      if (vertRow) {
        vertRow.style.display = (pos === 'left' || pos === 'right') ? 'flex' : 'none';
      }
      const widthSlider = document.getElementById('fpsTabVerticalWidth');
      const widthVal = document.getElementById('fpsTabWidthVal');
      const currentWidth = cfg.verticalWidth || 230;
      if (widthSlider) widthSlider.value = currentWidth;
      if (widthVal) widthVal.textContent = `${currentWidth}px`;

      // 3. Style & Size
      const styleSelect = document.getElementById('fpsTabStyleSelect');
      if (styleSelect) styleSelect.value = cfg.style || 'pills';

      const sizeSelect = document.getElementById('fpsTabSizeSelect');
      if (sizeSelect) sizeSelect.value = cfg.size || 'normal';

      // 4. Accent Color
      const accent = cfg.accentColor || '#0284c7';
      const picker = document.getElementById('fpsTabAccentColorPicker');
      const hex = document.getElementById('fpsTabAccentColorHex');
      const swatch = document.getElementById('fpsTabAccentColorSwatch');
      if (picker && accent.startsWith('#') && accent.length === 7) picker.value = accent;
      if (hex) hex.value = accent.toUpperCase();
      if (swatch) swatch.style.background = accent;

      document.querySelectorAll('#fpsTabPaletteRow .fps-palette-dot').forEach(dot => {
        dot.classList.toggle('active', dot.dataset.color.toLowerCase() === accent.toLowerCase());
      });

      // 5. Toggles
      const showIndexToggle = document.getElementById('fpsTabShowIndexToggle');
      if (showIndexToggle) showIndexToggle.checked = Boolean(cfg.showIndex);

      const showCloseToggle = document.getElementById('fpsTabShowCloseToggle');
      if (showCloseToggle) showCloseToggle.checked = (cfg.showCloseBtn !== false);

      // 6. Tab Items List
      this.renderTabItemsList();
    },

    renderTabItemsList() {
      const container = document.getElementById('fpsTabItemsList');
      if (!container) return;
      container.innerHTML = '';

      const tabs = DashboardState.present.tabs || [];
      const activeId = DashboardState.present.activeTabId;

      tabs.forEach((tab, index) => {
        const itemCard = document.createElement('div');
        itemCard.className = `fps-tab-item-card ${tab.id === activeId ? 'is-active' : ''}`;

        itemCard.innerHTML = `
          <span class="fps-tab-item-idx" title="Vị trí trang">${index + 1}</span>
          <input type="text" class="fps-tab-item-input" value="${escapeHtml(tab.name)}" title="Nhấp để đổi tên tab trực tiếp" />
          <div class="fps-tab-item-actions">
            <button type="button" class="fps-tab-action-icon-btn btn-tab-activate" title="Chuyển đến tab này">👁️</button>
            ${index > 0 ? `<button type="button" class="fps-tab-action-icon-btn btn-tab-up" title="Di chuyển lên trên">▲</button>` : ''}
            ${index < tabs.length - 1 ? `<button type="button" class="fps-tab-action-icon-btn btn-tab-down" title="Di chuyển xuống dưới">▼</button>` : ''}
            <button type="button" class="fps-tab-action-icon-btn btn-tab-duplicate" title="Nhân bản tab">📋</button>
            ${tabs.length > 1 ? `<button type="button" class="fps-tab-action-icon-btn is-del btn-tab-delete" title="Xóa tab này">🗑️</button>` : ''}
          </div>
        `;

        // Direct rename
        const nameInput = itemCard.querySelector('.fps-tab-item-input');
        nameInput?.addEventListener('change', (e) => {
          const newName = e.target.value.trim();
          if (newName) {
            syncItemPositionsFromDOM();
            DashboardState.pushSnapshot();
            tab.name = newName;
            DashboardState.isDirty = true;
            renderAllTabs();
            triggerDebouncedAutoSave();
            showToast('Đã đổi tên tab!', 'success');
          }
        });

        // Activate
        itemCard.querySelector('.btn-tab-activate')?.addEventListener('click', () => {
          if (DashboardState.present.activeTabId !== tab.id) {
            switchDashboardTab(tab.id);
            this.renderTabItemsList();
          }
        });

        // Move Up
        itemCard.querySelector('.btn-tab-up')?.addEventListener('click', () => {
          if (index > 0) {
            syncItemPositionsFromDOM();
            DashboardState.pushSnapshot();
            const temp = tabs[index];
            tabs[index] = tabs[index - 1];
            tabs[index - 1] = temp;
            DashboardState.isDirty = true;
            renderAllTabs();
            this.renderTabItemsList();
            triggerDebouncedAutoSave();
          }
        });

        // Move Down
        itemCard.querySelector('.btn-tab-down')?.addEventListener('click', () => {
          if (index < tabs.length - 1) {
            syncItemPositionsFromDOM();
            DashboardState.pushSnapshot();
            const temp = tabs[index];
            tabs[index] = tabs[index + 1];
            tabs[index + 1] = temp;
            DashboardState.isDirty = true;
            renderAllTabs();
            this.renderTabItemsList();
            triggerDebouncedAutoSave();
          }
        });

        // Duplicate
        itemCard.querySelector('.btn-tab-duplicate')?.addEventListener('click', () => {
          syncItemPositionsFromDOM();
          DashboardState.pushSnapshot();
          const dupTabId = 'tab_' + Date.now();
          const dupName = `${tab.name} (Bản sao)`;
          tabs.push({ id: dupTabId, name: dupName });

          // Duplicate items in tab
          const tabItems = DashboardState.present.items.filter(i => i.tabId === tab.id);
          tabItems.forEach(orig => {
            const copy = JSON.parse(JSON.stringify(orig));
            copy.id = orig.type + '_' + Date.now() + Math.random().toString(36).substr(2, 4);
            copy.tabId = dupTabId;
            DashboardState.present.items.push(copy);
          });

          // Duplicate tab filters nếu có
          if (DashboardState.tabFilters && DashboardState.tabFilters[tab.id]) {
            DashboardState.tabFilters[dupTabId] = JSON.parse(JSON.stringify(DashboardState.tabFilters[tab.id]));
          }

          switchDashboardTab(dupTabId);
          DashboardState.isDirty = true;
          this.renderTabItemsList();
          showToast(`Đã nhân bản tab "${dupName}"!`, 'success');
        });

        // Delete
        itemCard.querySelector('.btn-tab-delete')?.addEventListener('click', async () => {
          const ok = await AppDialog.confirm(`Xóa tab "${tab.name}" và toàn bộ biểu đồ bên trong?`, 'Xóa Tab', { isDanger: true, confirmText: 'Xóa Tab' });
          if (ok) {
            syncItemPositionsFromDOM();
            DashboardState.pushSnapshot();
            DashboardState.present.tabs = DashboardState.present.tabs.filter(t => t.id !== tab.id);
            DashboardState.present.items = DashboardState.present.items.filter(i => i.tabId !== tab.id);
            if (DashboardState.tabFilters) delete DashboardState.tabFilters[tab.id];
            const nextTabId = DashboardState.present.tabs[0]?.id;
            switchDashboardTab(nextTabId);
            DashboardState.isDirty = true;
            this.renderTabItemsList();
            showToast('Đã xóa tab', 'info');
          }
        });

        container.appendChild(itemCard);
      });
    }
  };

  // ===========================================================================
  // 11.5. CANVA-STYLE TEXT FORMATTING & INLINE EDITING ENGINE
  // ===========================================================================
  const CanvaTextManager = {
    toolbarEl: null,

    init() {
      if (this._initialized) return;
      this._initialized = true;
      this.toolbarEl = document.getElementById('canvaTextToolbar');
      if (!this.toolbarEl) return;

      // Font Family
      document.getElementById('cttFontFamily')?.addEventListener('change', (e) => {
        this.applyStyleToCurrent('fontFamily', e.target.value);
      });

      // Font Size Controls
      document.getElementById('cttSizeDown')?.addEventListener('click', () => {
        this.stepFontSize(-2);
      });
      document.getElementById('cttSizeUp')?.addEventListener('click', () => {
        this.stepFontSize(2);
      });
      document.getElementById('cttFontSizeInput')?.addEventListener('change', (e) => {
        const val = Math.max(8, Math.min(140, parseInt(e.target.value, 10) || 20));
        this.applyStyleToCurrent('fontSize', `${val}px`);
      });

      // Color Picker
      const colorInput = document.getElementById('cttColorPicker');
      colorInput?.addEventListener('input', (e) => {
        const c = e.target.value;
        const ind = document.getElementById('cttColorIndicator');
        if (ind) ind.style.background = c;
        this.applyStyleToCurrent('color', c);
      });

      // Bold Toggle
      document.getElementById('cttBoldBtn')?.addEventListener('click', () => {
        this.toggleBold();
      });

      // Italic Toggle
      document.getElementById('cttItalicBtn')?.addEventListener('click', () => {
        this.toggleItalic();
      });

      // Underline Toggle
      document.getElementById('cttUnderlineBtn')?.addEventListener('click', () => {
        this.toggleUnderline();
      });

      // Alignment Buttons
      document.querySelectorAll('#canvaTextToolbar [data-align]').forEach(btn => {
        btn.addEventListener('click', () => {
          this.applyStyleToCurrent('textAlign', btn.dataset.align);
          this.updateToolbarActiveStates();
        });
      });

      // Edit text button -> trigger inline edit
      document.getElementById('cttEditTextBtn')?.addEventListener('click', () => {
        const itemId = this.toolbarEl.dataset.currentItemId;
        const item = DashboardState.present.items.find(i => i.id === itemId);
        const card = document.getElementById(`item_${itemId}`);
        const textEl = card?.querySelector('.db-text-content-editable');
        if (item && textEl) {
          this.startInlineEdit(item, textEl);
        }
      });
    },

    getCurrentItem() {
      const itemId = this.toolbarEl?.dataset.currentItemId;
      if (!itemId) return null;
      return DashboardState.present.items.find(i => i.id === itemId) || null;
    },

    show(item, cardEl) {
      if (!this.toolbarEl || DashboardState.mode === 'preview') return;
      if (item.type !== 'heading' && item.type !== 'text') {
        this.hide();
        return;
      }

      this.toolbarEl.dataset.currentItemId = item.id;
      this.syncToolbarValues(item);

      this.toolbarEl.style.display = 'flex';
      // Calculate top / left relative to canvas container
      const x = Math.max(10, item.pos.x);
      const y = Math.max(10, item.pos.y - 44);
      this.toolbarEl.style.left = `${x}px`;
      this.toolbarEl.style.top = `${y}px`;
    },

    hide() {
      if (!this.toolbarEl) return;
      this.toolbarEl.style.display = 'none';
      this.toolbarEl.dataset.currentItemId = '';
    },

    syncToolbarValues(item) {
      const fontSelect = document.getElementById('cttFontFamily');
      if (fontSelect) fontSelect.value = item.fontFamily || 'Inter, sans-serif';

      const sizeInput = document.getElementById('cttFontSizeInput');
      const curSize = parseInt(item.fontSize || '20', 10) || 20;
      if (sizeInput) sizeInput.value = curSize;

      const color = item.color || '#0f172a';
      const colorInput = document.getElementById('cttColorPicker');
      const colorInd = document.getElementById('cttColorIndicator');
      if (colorInput) colorInput.value = color.startsWith('#') ? color : '#0f172a';
      if (colorInd) colorInd.style.background = color;

      this.updateToolbarActiveStates(item);
    },

    updateToolbarActiveStates(item) {
      if (!item) item = this.getCurrentItem();
      if (!item) return;

      const isBold = item.fontWeight === 'bold' || item.fontWeight === '700' || item.fontWeight === 800;
      document.getElementById('cttBoldBtn')?.classList.toggle('active', Boolean(isBold));

      const isItalic = item.fontStyle === 'italic';
      document.getElementById('cttItalicBtn')?.classList.toggle('active', Boolean(isItalic));

      const isUnderline = item.textDecoration === 'underline';
      document.getElementById('cttUnderlineBtn')?.classList.toggle('active', Boolean(isUnderline));

      const align = item.textAlign || 'left';
      document.querySelectorAll('#canvaTextToolbar [data-align]').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.align === align);
      });
    },

    applyStyleToCurrent(prop, val) {
      const item = this.getCurrentItem();
      if (!item) return;
      DashboardState.pushSnapshot();
      item[prop] = val;
      renderItemContent(item);
      this.syncToolbarValues(item);
    },

    stepFontSize(delta) {
      const item = this.getCurrentItem();
      if (!item) return;
      const cur = parseInt(item.fontSize || '20', 10) || 20;
      const next = Math.max(8, Math.min(140, cur + delta));
      this.applyStyleToCurrent('fontSize', `${next}px`);
    },

    toggleBold() {
      const item = this.getCurrentItem();
      if (!item) return;
      const isBold = item.fontWeight === 'bold' || item.fontWeight === '700' || item.fontWeight === 800;
      this.applyStyleToCurrent('fontWeight', isBold ? 'normal' : 'bold');
    },

    toggleItalic() {
      const item = this.getCurrentItem();
      if (!item) return;
      const isItalic = item.fontStyle === 'italic';
      this.applyStyleToCurrent('fontStyle', isItalic ? 'normal' : 'italic');
    },

    toggleUnderline() {
      const item = this.getCurrentItem();
      if (!item) return;
      const isUnderline = item.textDecoration === 'underline';
      this.applyStyleToCurrent('textDecoration', isUnderline ? 'none' : 'underline');
    },

    startInlineEdit(item, targetEl) {
      if (DashboardState.mode === 'preview') return;
      targetEl.setAttribute('contenteditable', 'true');
      targetEl.classList.add('is-editing');
      targetEl.focus();

      // Select all text
      const range = document.createRange();
      range.selectNodeContents(targetEl);
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);

      const commit = () => {
        targetEl.removeAttribute('contenteditable');
        targetEl.classList.remove('is-editing');
        const newText = targetEl.innerText.trim();
        if (newText !== item.text) {
          DashboardState.pushSnapshot();
          item.text = newText;
          item.name = newText.substring(0, 30) || item.name;
          const cardTitle = document.querySelector(`#item_${item.id} .db-item-title`);
          if (cardTitle) cardTitle.textContent = item.name;
        }
      };

      const keyHandler = (e) => {
        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'b') {
          e.preventDefault();
          this.toggleBold();
        } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'i') {
          e.preventDefault();
          this.toggleItalic();
        } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'u') {
          e.preventDefault();
          this.toggleUnderline();
        } else if (e.key === 'Escape') {
          e.preventDefault();
          targetEl.blur();
        }
      };

      targetEl.addEventListener('keydown', keyHandler);
      targetEl.addEventListener('blur', () => {
        targetEl.removeEventListener('keydown', keyHandler);
        commit();
      }, { once: true });
    }
  };

  // Note: getFilteredData is defined above in Section 3 (DATA FILTERING ENGINE & ECHARTS EVALUATION)

  // ===========================================================================
  // TÍNH TOÁN VỊ TRÍ CHO THÀNH PHẦN MỚI (LUÔN XẾP XUỐNG DƯỚI CÁC THÀNH PHẦN ĐÃ CÓ)
  // ===========================================================================
  function getNextAvailablePosition(defaultW = 320, defaultH = 180) {
    syncItemPositionsFromDOM();
    const activeTab = DashboardState.present.activeTabId;
    const tabItems = DashboardState.present.items.filter(i => i.tabId === activeTab);

    if (tabItems.length === 0) {
      return { x: 30, y: 30, w: defaultW, h: defaultH };
    }

    // Tìm điểm đáy thấp nhất của tất cả thành phần hiện có trong tab hiện tại
    let maxBottom = 0;
    tabItems.forEach(item => {
      const b = (item.pos?.y || 0) + (item.pos?.h || 0);
      if (b > maxBottom) {
        maxBottom = b;
      }
    });

    // Thêm vào ngay phía dưới các thành phần hiện tại với khoảng cách 24px
    const nextY = maxBottom + 24;
    const nextX = 30;

    return { x: nextX, y: nextY, w: defaultW, h: defaultH };
  }

  const TopFilterBarManager = {
    getConfig(tabId = null) {
      const curTab = tabId || DashboardState.present.activeTabId || 'default';
      if (!DashboardState.present.topFilterBarConfig) DashboardState.present.topFilterBarConfig = {};
      if (!DashboardState.present.topFilterBarConfig[curTab]) {
        DashboardState.present.topFilterBarConfig[curTab] = {
          enabled: false,
          isSticky: true,
          isCollapsed: false,
          collapseStyle: 'circle',
          title: 'BỘ LỌC TỔNG HỢP',
          pos: { x: 24, y: 24, w: 1312 }
        };
      }
      return DashboardState.present.topFilterBarConfig[curTab];
    },

    saveConfig(cfg, tabId = null) {
      const curTab = tabId || DashboardState.present.activeTabId || 'default';
      if (!DashboardState.present.topFilterBarConfig) DashboardState.present.topFilterBarConfig = {};
      DashboardState.present.topFilterBarConfig[curTab] = {
        ...this.getConfig(curTab),
        ...cfg
      };
      DashboardState.isDirty = true;
    }
  };

  function attachTopFilterBarDrag(topBarEl, barConfig, activeTab) {
    const dragTargets = [
      topBarEl.querySelector('.db-top-filter-header'),
      topBarEl.querySelector('.db-top-filter-drag-grip'),
      topBarEl.querySelector('.db-top-filter-title-group'),
      topBarEl.querySelector('.db-top-filter-circle-btn')
    ].filter(Boolean);

    dragTargets.forEach(target => {
      target.addEventListener('mousedown', (e) => {
        if (DashboardState.mode === 'preview') return;
        if (e.target.closest('button') || e.target.closest('input') || e.target.closest('select')) return;
        if (e._dragHandled) return;
        e._dragHandled = true;

        e.preventDefault();
        e.stopPropagation();

        const zoom = DashboardState.zoom || 1.0;
        const startMouseX = e.clientX;
        const startMouseY = e.clientY;
        const origX = barConfig.pos?.x !== undefined ? barConfig.pos.x : 24;
        const origY = barConfig.pos?.y !== undefined ? barConfig.pos.y : 24;
        let hasMoved = false;

        topBarEl.classList.add('is-dragging');

        function onMouseMove(ev) {
          const dx = (ev.clientX - startMouseX) / zoom;
          const dy = (ev.clientY - startMouseY) / zoom;
          if (Math.abs(dx) > 3 || Math.abs(dy) > 3) {
            hasMoved = true;
          }
          const newX = Math.max(24, Math.min(1360 - 200, Math.round(origX + dx)));
          const newY = Math.max(10, Math.round(origY + dy));
          topBarEl.style.left = `${newX}px`;
          topBarEl.style.top = `${newY}px`;
          barConfig.pos = barConfig.pos || {};
          barConfig.pos.x = newX;
          barConfig.pos.y = newY;
          topBarEl.dataset.origTop = newY;

          // Mở rộng chiều cao canvas nếu kéo thanh xuống phía dưới
          const canvas = document.getElementById('dashboardCanvas');
          if (canvas) {
            const barBottom = newY + topBarEl.offsetHeight + 80;
            const currentMinH = parseInt(canvas.style.minHeight || '800', 10);
            if (barBottom > currentMinH) {
              canvas.style.minHeight = `${barBottom}px`;
            }
          }
        }

        function onMouseUp() {
          window.removeEventListener('mousemove', onMouseMove);
          window.removeEventListener('mouseup', onMouseUp);
          topBarEl.classList.remove('is-dragging');
          if (hasMoved) {
            topBarEl.dataset.origTop = barConfig.pos?.y || 24;
            TopFilterBarManager.saveConfig({ pos: { ...barConfig.pos } }, activeTab);
            triggerDebouncedAutoSave();
            topBarEl._wasDragged = true;
            setTimeout(() => { delete topBarEl._wasDragged; }, 120);
          }
        }

        window.addEventListener('mousemove', onMouseMove);
        window.addEventListener('mouseup', onMouseUp);
      });
    });
  }

  function unpinFilterToCanvas(item) {
    DashboardState.pushSnapshot();
    item.isPinnedTop = false;
    item.container = null;
    const nextPos = getNextAvailablePosition(320, 95);
    item.pos = { x: nextPos.x, y: Math.max(140, nextPos.y), w: 320, h: 95, zIndex: 10 };
    DashboardState.isDirty = true;
    renderCanvas();
    showToast(`✓ Đã đưa bộ lọc "${item.name || item.filterColumn}" về canvas tự do.`, 'info');
  }

  function pinFilterToTopBar(item) {
    DashboardState.pushSnapshot();
    const curTab = item.tabId || DashboardState.present.activeTabId || 'default';
    const cfg = TopFilterBarManager.getConfig(curTab);
    if (!cfg.enabled) {
      cfg.enabled = true;
      TopFilterBarManager.saveConfig({ enabled: true }, curTab);
    }
    item.isPinnedTop = true;
    item.container = 'sticky_top';
    DashboardState.isDirty = true;
    renderCanvas();
    showToast(`✓ Đã ghim bộ lọc "${item.name || item.filterColumn}" lên thanh bộ lọc tổng hợp!`, 'success');
  }

  function updateTopFilterBarStickyPosition() {
    const topBar = document.getElementById('dbTopFilterBar');
    if (!topBar) return;
    const canvas = document.getElementById('dashboardCanvas');
    if (!canvas) return;

    // Chỉ ghim khi ở chế độ Xem Trước (Preview) hoặc Bản Xuất Độc Lập
    // Trong chế độ Edit (thiết kế), KHÔNG ghim để người dùng thao tác kéo thả và quan sát canvas tự do
    if (typeof DashboardState !== 'undefined' && DashboardState.mode !== 'preview') {
      topBar.style.transform = '';
      topBar.classList.remove('is-pinned-floating');
      return;
    }

    const canvasRect = canvas.getBoundingClientRect();
    const effectiveScale = canvas.offsetWidth > 0 ? (canvasRect.width / canvas.offsetWidth) : 1.0;

    let initialLocalY = 24;
    if (topBar.dataset.origTop !== undefined && topBar.dataset.origTop !== '') {
      initialLocalY = parseFloat(topBar.dataset.origTop) || 24;
    } else {
      initialLocalY = topBar.offsetTop || 24;
      topBar.dataset.origTop = initialLocalY;
    }

    // Xác định toạ độ đỉnh mục tiêu (targetScreenY) trên màn hình viewport:
    // Nếu header đang sticky ở đỉnh viewport (top <= 2 && bottom > 0): ghim sát ngay mép dưới header.
    // Nếu không có header hoặc header đã cuộn đi: ghim sát mép đỉnh màn hình (targetScreenY = 0px, không khoảng hở)!
    const headerEl = document.querySelector('.db-header');
    let headerBottom = 0;
    if (headerEl && headerEl.offsetHeight > 0) {
      const headerRect = headerEl.getBoundingClientRect();
      if (headerRect.top <= 2 && headerRect.bottom > 0) {
        headerBottom = headerRect.bottom;
      }
    }

    const tabsBar = document.getElementById('dbTabsBar');
    let tabsBottom = 0;
    if (tabsBar && tabsBar.offsetHeight > 0) {
      const tabsRect = tabsBar.getBoundingClientRect();
      if (tabsRect.top <= headerBottom + 2 && tabsRect.bottom > headerBottom) {
        tabsBottom = tabsRect.bottom;
      }
    }

    // Ghim sát lên đầu hẳn: 0px khoảng trống thừa
    const targetScreenY = Math.max(headerBottom, tabsBottom);

    // Toạ độ tự nhiên của filter bar trên màn hình viewport (nếu không có translateY)
    const naturalScreenY = canvasRect.top + (initialLocalY * effectiveScale);

    if (naturalScreenY <= targetScreenY) {
      // Cần dịch chuyển topBar trong Canvas sao cho:
      // canvasRect.top + (initialLocalY + neededTranslateY) * effectiveScale === targetScreenY
      const neededTranslateY = (targetScreenY - canvasRect.top) / effectiveScale - initialLocalY;
      const maxTranslateY = Math.max(0, canvas.offsetHeight - topBar.offsetHeight - 20);
      const finalY = Math.min(maxTranslateY, Math.max(0, neededTranslateY));

      topBar.style.transform = `translate3d(0, ${finalY}px, 0)`;
      topBar.classList.add('is-pinned-floating');
    } else {
      topBar.style.transform = '';
      topBar.classList.remove('is-pinned-floating');
    }
  }

  function renderCanvas() {
    syncItemPositionsFromDOM();
    const canvas = document.getElementById('dashboardCanvas');
    if (!canvas) return;
    canvas.innerHTML = '';

    const activeTab = DashboardState.present.activeTabId || 'default';
    const barConfig = TopFilterBarManager.getConfig(activeTab);
    const currentTabItems = DashboardState.present.items.filter(i => i.tabId === activeTab);

    const pinnedFilterItems = currentTabItems.filter(i => i.type === 'filter' && (i.isPinnedTop || i.container === 'sticky_top'));
    const freeCanvasItems = currentTabItems.filter(i => !(i.type === 'filter' && (i.isPinnedTop || i.container === 'sticky_top')));

    // Giới hạn chiều cao vừa vặn các thành phần tự do trên canvas
    let maxContentBottom = 800;
    freeCanvasItems.forEach(item => {
      const bottom = (item.pos?.y || 0) + (item.pos?.h || 0);
      if (bottom > maxContentBottom) maxContentBottom = bottom;
    });
    if (barConfig.enabled || pinnedFilterItems.length > 0) {
      const barBottom = (barConfig.pos?.y || 24) + 120;
      if (barBottom > maxContentBottom) maxContentBottom = barBottom;
    }
    const baseMinH = DashboardState.mode === 'preview' ? (window.innerHeight - 56) : 800;
    const computedMinHeight = Math.max(baseMinH, maxContentBottom + 60);
    canvas.style.minHeight = `${computedMinHeight}px`;
    updateCanvasTransform();

    // 2 vạch ranh giới an toàn 2 mép Trái - Phải của Dashboard (Trái: 24px, Phải: 1336px)
    const leftGuide = document.createElement('div');
    leftGuide.className = 'db-canvas-boundary-guide is-left';
    canvas.appendChild(leftGuide);

    const rightGuide = document.createElement('div');
    rightGuide.className = 'db-canvas-boundary-guide is-right';
    canvas.appendChild(rightGuide);

    // Chỉ giữ lại container snapping guides (các đường gióng đỏ/xanh tạm thời khi kéo thả)
    if (DashboardState.mode !== 'preview') {
      const snappingGuidesEl = document.createElement('div');
      snappingGuidesEl.className = 'db-snapping-guides';
      snappingGuidesEl.id = 'snappingGuides';
      canvas.appendChild(snappingGuidesEl);
    }

    // 1. RENDER THANH BỘ LỌC TỔNG HỢP NẾU ĐƯỢC BẬT HOẶC CÓ BỘ LỌC ĐÃ GHIM
    const shouldRenderTopBar = barConfig.enabled || pinnedFilterItems.length > 0;
    if (shouldRenderTopBar) {
      const topBarEl = document.createElement('div');
      const isCollapsed = Boolean(barConfig.isCollapsed);
      const collapseStyle = barConfig.collapseStyle || 'circle';
      const isCircle = collapseStyle === 'circle';
      topBarEl.id = 'dbTopFilterBar';

      topBarEl.className = `db-top-filter-bar is-movable ${isCollapsed ? 'is-collapsed' : ''} ${isCircle ? 'collapse-style-circle' : 'collapse-style-bar'}`;
      topBarEl.style.position = 'absolute';
      topBarEl.style.left = `${barConfig.pos?.x || 24}px`;
      topBarEl.style.top = `${barConfig.pos?.y || 24}px`;
      topBarEl.style.width = isCollapsed && isCircle ? '48px' : `${barConfig.pos?.w || 1312}px`;
      topBarEl.style.zIndex = '900';

      const countLabel = pinnedFilterItems.length === 0 ? 'Chưa có bộ lọc' : `${pinnedFilterItems.length} bộ lọc`;
      topBarEl.innerHTML = `
        <div class="db-top-filter-circle-btn" id="btnTopFilterCircleOpen" title="🔍 Mở Bộ Lọc Tổng Hợp (${countLabel})">
          <span class="db-filter-circle-icon">🔍</span>
          ${pinnedFilterItems.length > 0 ? `<span class="db-filter-circle-badge">${pinnedFilterItems.length}</span>` : ''}
        </div>
        <div class="db-top-filter-expanded-wrap">
          <div class="db-top-filter-header" title="${DashboardState.mode !== 'preview' ? 'Kéo để di chuyển vị trí thanh bộ lọc trên Canvas' : ''}">
            <div class="db-top-filter-title-group">
              ${DashboardState.mode !== 'preview' ? `<span class="db-top-filter-drag-grip" title="Kéo để di chuyển vị trí thanh bộ lọc trên Canvas">⠿</span>` : ''}
              <span class="db-top-filter-icon">🔍</span>
              <span class="db-top-filter-title">${escapeHtml(barConfig.title || 'BỘ LỌC TỔNG HỢP')}</span>
              <span class="db-top-filter-count-badge">📌 ${countLabel}</span>
            </div>
            <div class="db-top-filter-actions">
              ${DashboardState.mode !== 'preview' ? `
                <button type="button" class="db-top-filter-btn" id="btnToggleCollapseStyle" title="Chuyển đổi kiểu khi đóng: Nút Tròn gọn gàng hoặc Thanh Ngang">${isCircle ? '🔘 Kiểu đóng: Nút Tròn' : '▬ Kiểu đóng: Thanh Ngang'}</button>
                <button type="button" class="db-top-filter-btn" id="btnResetPosTopBar" title="Căn nhanh thanh bộ lọc về đầu Canvas (X: 24px, Y: 24px)">⬆️ Về đầu Canvas</button>
              ` : ''}
              <button type="button" class="db-top-filter-btn primary" id="btnToggleCollapseTopBar" title="Thu gọn bộ lọc">
                ${isCircle ? '▲ Thu Gọn (Nút Tròn)' : '▲ Thu Gọn'}
              </button>
              ${DashboardState.mode !== 'preview' ? `
                <button type="button" class="db-top-filter-btn is-danger" id="btnCloseTopBar" title="Tắt thanh bộ lọc tổng hợp (đưa các bộ lọc về canvas)">✕ Tắt</button>
              ` : ''}
            </div>
          </div>
          <div class="db-top-filter-body" id="dbTopFilterBody"></div>
        </div>
      `;

      const topBodyEl = topBarEl.querySelector('#dbTopFilterBody');

      if (pinnedFilterItems.length === 0) {
        topBodyEl.innerHTML = `
          <div class="db-top-filter-empty-dropzone" id="dbTopFilterDropzone">
            <span style="font-size:16px;">📥</span>
            <span>Kéo thả các Bộ Lọc vào đây để gán vào thanh tổng hợp, hoặc bấm nút "📌 Ghim Top" trên bộ lọc.</span>
          </div>
        `;
      } else {
        pinnedFilterItems.forEach(item => {
          const wrapper = document.createElement('div');
          wrapper.className = 'db-top-filter-item-wrapper';
          wrapper.id = `pinned_item_${item.id}`;

          wrapper.innerHTML = `
            <div class="db-top-filter-item-header">
              <span class="db-top-filter-item-title">${escapeHtml(item.name || item.filterColumn)}</span>
              ${DashboardState.mode !== 'preview' ? `
                <button type="button" class="db-top-filter-unpin-btn" data-unpin-id="${item.id}" title="Tháo khỏi thanh trên cùng, chuyển về canvas tự do">📌 Bỏ ghim</button>
              ` : ''}
            </div>
            <div class="db-item-body" id="item_body_${item.id}"></div>
          `;
          topBodyEl.appendChild(wrapper);

          wrapper.querySelector('[data-unpin-id]')?.addEventListener('click', (e) => {
            e.stopPropagation();
            unpinFilterToCanvas(item);
          });
        });
      }

      canvas.appendChild(topBarEl);

      // Lưu lại toạ độ gốc ban đầu khi chưa bị translate
      topBarEl.dataset.origTop = topBarEl.offsetTop;

      // Render nội dung bộ lọc SAU KHI topBarEl và các container item_body_${item.id} đã gắn vào DOM
      pinnedFilterItems.forEach(item => {
        renderItemContent(item);
      });

      // Mở thanh bộ lọc khi nhấp vào nút tròn
      topBarEl.querySelector('#btnTopFilterCircleOpen')?.addEventListener('click', (e) => {
        if (topBarEl._wasDragged) return;
        e.stopPropagation();
        barConfig.isCollapsed = false;
        TopFilterBarManager.saveConfig({ isCollapsed: false }, activeTab);
        renderCanvas();
        triggerDebouncedAutoSave();
      });

      // Kéo di chuyển thanh khi ở chế độ thiết kế (Edit)
      if (DashboardState.mode !== 'preview') {
        attachTopFilterBarDrag(topBarEl, barConfig, activeTab);

        topBarEl.querySelector('#btnToggleCollapseStyle')?.addEventListener('click', (e) => {
          e.stopPropagation();
          const nextStyle = barConfig.collapseStyle === 'circle' ? 'bar' : 'circle';
          barConfig.collapseStyle = nextStyle;
          TopFilterBarManager.saveConfig({ collapseStyle: nextStyle }, activeTab);
          renderCanvas();
          triggerDebouncedAutoSave();
          showToast(nextStyle === 'circle' ? '✓ Đã chọn: Thu gọn thành Nút Tròn gọn gàng.' : '✓ Đã chọn: Thu gọn thành Thanh Ngang tiêu đề.', 'info');
        });

        topBarEl.querySelector('#btnResetPosTopBar')?.addEventListener('click', (e) => {
          e.stopPropagation();
          barConfig.pos = { x: 24, y: 24, w: 1312 };
          topBarEl.dataset.origTop = 24;
          TopFilterBarManager.saveConfig({ pos: barConfig.pos }, activeTab);
          renderCanvas();
          triggerDebouncedAutoSave();
          showToast('✓ Đã căn thanh bộ lọc về vị trí đầu Canvas (X: 24, Y: 24).', 'info');
        });
      }

      topBarEl.querySelector('#btnToggleCollapseTopBar')?.addEventListener('click', (e) => {
        e.stopPropagation();
        const nextCollapsed = !barConfig.isCollapsed;
        barConfig.isCollapsed = nextCollapsed;
        TopFilterBarManager.saveConfig({ isCollapsed: nextCollapsed }, activeTab);
        renderCanvas();
        triggerDebouncedAutoSave();
      });

      topBarEl.querySelector('#btnCloseTopBar')?.addEventListener('click', async (e) => {
        e.stopPropagation();
        const ok = await AppDialog.confirm('Bạn có chắc muốn tắt Thanh Bộ Lọc Tổng Hợp? Các bộ lọc bên trong sẽ được đưa về Canvas tự do.', 'Tắt Thanh Bộ Lọc', { confirmText: 'Tắt Thanh', isDanger: true });
        if (ok) {
          DashboardState.pushSnapshot();
          pinnedFilterItems.forEach(item => {
            item.isPinnedTop = false;
            item.container = null;
            const nextPos = getNextAvailablePosition(320, 95);
            item.pos = { x: nextPos.x, y: Math.max(140, nextPos.y), w: 320, h: 95, zIndex: 10 };
          });
          TopFilterBarManager.saveConfig({ enabled: false }, activeTab);
          DashboardState.isDirty = true;
          renderCanvas();
          triggerDebouncedAutoSave();
          showToast('✓ Đã tắt Thanh Bộ Lọc Tổng Hợp.', 'info');
        }
      });
    }

    // Đảm bảo toạ độ các widget tự do luôn nằm an toàn bên trong 2 ranh giới (Trái: 24px, Phải: 1336px)
    freeCanvasItems.forEach(item => {
      if (!item.pos) item.pos = { x: 24, y: 24, w: 400, h: 300 };
      if (item.pos.y < 24) item.pos.y = 24;
      if (item.pos.x < 24) item.pos.x = 24;
      if (item.pos.w > 1312) item.pos.w = 1312;
      if (item.pos.x + item.pos.w > 1336) {
        item.pos.x = Math.max(24, 1336 - item.pos.w);
      }
    });

    if (currentTabItems.length === 0) {
      const empty = document.createElement('div');
      empty.className = 'db-empty-canvas';
      empty.innerHTML = `
        <div class="db-empty-icon">📊</div>
        <div class="db-empty-title">Tab Này Chưa Có Biểu Đồ Hoặc Widget Nào</div>
        <div class="db-empty-desc">Nhấn vào nút "+ Thêm Biểu Đồ Mới" để thiết kế mã biểu đồ ECharts, hoặc "+ Thêm Bộ Lọc Mới" để lọc dữ liệu Excel theo các điều kiện.</div>
        <div style="display:flex;gap:12px;">
          <button class="db-btn-tool primary" id="btnEmptyAddChart">📊 + Thêm Biểu Đồ Mới</button>
          <button class="db-btn-tool" id="btnEmptyAddFilter">🔍 + Thêm Bộ Lọc Mới</button>
        </div>
      `;
      canvas.appendChild(empty);
      document.getElementById('btnEmptyAddChart')?.addEventListener('click', () => ChartStudioModal.open());
      document.getElementById('btnEmptyAddFilter')?.addEventListener('click', () => WidgetConfigModal.open('filter'));
      CanvaTextManager.hide();
      return;
    }

    freeCanvasItems.forEach(item => {
      const el = document.createElement('div');
      const isLocked = Boolean(item.isLocked);
      const isSelected = DashboardState.mode !== 'preview' && DashboardState.selectedItemId === item.id;
      const isHideHeader = Boolean(item.hideHeader);
      const isFramelessType = ['shape', 'heading', 'text', 'image'].includes(item.type);
      const shapeSubclass = item.type === 'shape' ? (item.shapeType === 'line' ? 'is-shape-line' : 'is-shape') : '';
      el.className = `db-item-box ${item.type === 'frame' ? 'is-frame' : ''} ${shapeSubclass} ${isFramelessType ? 'is-frameless-element' : ''} is-type-${item.type} ${isLocked ? 'is-locked' : ''} ${isSelected ? 'selected' : ''} ${isHideHeader ? 'hide-header' : ''}`;
      el.id = `item_${item.id}`;
      el.style.left = `${item.pos.x}px`;
      el.style.top = `${item.pos.y}px`;
      el.style.width = `${item.pos.w}px`;
      el.style.height = `${item.pos.h}px`;
      if (item.pos.zIndex) {
        el.style.zIndex = item.pos.zIndex;
      }

      if (DashboardState.mode === 'preview') {
        if (isFramelessType || isHideHeader) {
          el.innerHTML = `<div class="db-item-body" id="item_body_${item.id}" style="${isFramelessType ? 'padding:0;' : ''}height:100%;"></div>`;
        } else {
          el.innerHTML = `
            <div class="db-item-header">
              <div class="db-item-header-left">
                <span class="db-item-title">${escapeHtml(item.name || item.type)}</span>
              </div>
            </div>
            <div class="db-item-body" id="item_body_${item.id}"></div>
          `;
        }
      } else {
        el.innerHTML = `
          <div class="db-item-header">
            <div class="db-item-header-left">
              <span class="db-item-drag-grip" title="${isLocked ? 'Widget đang bị khóa vị trí' : 'Kéo để di chuyển (Snapping tự động)'}">⠿</span>
              <span class="db-item-title">${escapeHtml(item.name || item.type)}</span>
            </div>
            <div class="db-item-actions">
              ${item.type === 'filter' ? `<button class="db-item-action-btn pin-btn" data-action="pin-top" data-id="${item.id}" title="Ghim lên thanh cố định ở trên cùng">📌 Ghim Top</button>` : ''}
              <button class="db-item-action-btn" data-action="toggle-header" data-id="${item.id}" title="Ẩn thanh tiêu đề (Chỉ hiện chart &amp; nội dung)">🕶️</button>
              <button class="db-item-action-btn" data-action="toggle-lock" data-id="${item.id}" title="${isLocked ? 'Mở khóa di chuyển' : 'Khóa vị trí'}">${isLocked ? '🔒' : '🔓'}</button>
              <button class="db-item-action-btn" data-action="z-up" data-id="${item.id}" title="Đưa lên trên (Bring to front)">⬆️</button>
              <button class="db-item-action-btn" data-action="z-down" data-id="${item.id}" title="Đưa xuống dưới (Send to back)">⬇️</button>
              <button class="db-item-action-btn props-btn" data-action="open-sidebar" data-id="${item.id}" title="Mở bảng Thuộc tính / Sidebar (Figma Properties)">⚙️ Thuộc tính</button>
              ${item.type === 'chart' ? `<button class="db-item-action-btn edit" data-action="edit" data-id="${item.id}" title="Sửa code ECharts">✏️ Code</button>` : ''}
              ${item.type !== 'chart' ? `<button class="db-item-action-btn edit" data-action="edit-widget" data-id="${item.id}" title="Chỉnh sửa widget">✏️ Sửa</button>` : ''}
              <button class="db-item-action-btn" data-action="duplicate" data-id="${item.id}" title="Nhân bản">📋</button>
              <button class="db-item-action-btn delete" data-action="delete" data-id="${item.id}" title="Xóa">🗑️</button>
            </div>
          </div>
          <!-- Floating mini actions when header is hidden -->
          <div class="db-item-mini-actions" data-id="${item.id}">
            <span class="db-item-mini-grip" title="Kéo để di chuyển">⠿</span>
            ${item.type === 'filter' ? `<button class="db-item-mini-btn" data-action="pin-top" data-id="${item.id}" title="Ghim lên thanh trên cùng">📌 Ghim</button>` : ''}
            <button class="db-item-mini-btn" data-action="toggle-header" data-id="${item.id}" title="Hiện lại thanh tiêu đề (Header)">👁️ Hiện Header</button>
            <button class="db-item-mini-btn props-btn" data-action="open-sidebar" data-id="${item.id}" title="Mở bảng Thuộc tính / Sidebar">⚙️ Thuộc tính</button>
            ${item.type === 'chart' ? `<button class="db-item-mini-btn" data-action="edit" data-id="${item.id}" title="Sửa code ECharts">✏️ Code</button>` : `<button class="db-item-mini-btn" data-action="edit-widget" data-id="${item.id}" title="Sửa widget">✏️ Sửa</button>`}
            <button class="db-item-mini-btn is-danger" data-action="delete" data-id="${item.id}" title="Xóa">🗑️</button>
          </div>
          <div class="db-item-body" id="item_body_${item.id}"></div>
          <div class="db-item-resize-handle" title="${isLocked ? 'Đang khóa' : 'Kéo để thay đổi kích thước'}">◢</div>
        `;
      }

      // Select item on click
      el.addEventListener('click', (e) => {
        if (DashboardState.mode === 'preview') return;
        if (e.target.closest('.db-item-action-btn') || e.target.closest('.db-item-mini-btn')) return;
        DashboardState.selectedItemId = item.id;
        document.querySelectorAll('.db-item-box').forEach(b => b.classList.remove('selected'));
        el.classList.add('selected');

        if (item.type === 'heading' || item.type === 'text') {
          CanvaTextManager.show(item, el);
        } else {
          CanvaTextManager.hide();
        }

        // KHÔNG tự động mở sidebar nếu đang đóng. Chỉ sync dữ liệu nếu người dùng đã mở sẵn sidebar trước đó!
        if (typeof FigmaPropertiesSidebar !== 'undefined' && FigmaPropertiesSidebar.isOpen && FigmaPropertiesSidebar.isOpen()) {
          FigmaPropertiesSidebar.open(item.id);
        }
      });

      canvas.appendChild(el);
      renderItemContent(item);
      if (DashboardState.mode !== 'preview') {
        attachItemDragAndResize(el, item);
      }
    });

    // Deselect and hide toolbar when clicking on canvas background
    canvas.addEventListener('click', (e) => {
      if (DashboardState.mode === 'preview') return;
      if (!e.target.closest('.db-item-box') && !e.target.closest('#canvaTextToolbar') && !e.target.closest('#figmaPropertiesSidebar')) {
        DashboardState.selectedItemId = null;
        document.querySelectorAll('.db-item-box').forEach(b => b.classList.remove('selected'));
        CanvaTextManager.hide();
        if (typeof FigmaPropertiesSidebar !== 'undefined') {
          FigmaPropertiesSidebar.close();
        }
      }
    });

    updateTopFilterBarStickyPosition();
  }

  function renderItemContent(item) {
    const body = document.getElementById(`item_body_${item.id}`);
    if (!body) return;
    body.innerHTML = '';

    if (item.type === 'chart') {
      const chartDiv = document.createElement('div');
      chartDiv.className = 'db-chart-canvas';
      chartDiv.id = `chart_el_${item.id}`;
      body.appendChild(chartDiv);

      // Nút Edit trực tiếp trên biểu đồ để mở Studio Popup (chỉ hiện trong Chế độ Thiết Kế, không hiện khi Xem Trước)
      if (DashboardState.mode !== 'preview') {
        const editBadge = document.createElement('div');
        editBadge.className = 'db-chart-edit-badge';
        editBadge.title = 'Nhấn để mở Studio Popup: chỉnh sửa mã ECharts JS và xem Live Demo';
        editBadge.innerHTML = '<span>✏️</span><span>Sửa Biểu Đồ (Popup)</span>';
        editBadge.addEventListener('click', (e) => {
          e.stopPropagation();
          ChartStudioModal.open(item.id);
        });
        body.appendChild(editBadge);
      }

      setTimeout(() => {
        let chart = echarts.getInstanceByDom(chartDiv);
        if (chart) chart.dispose();
        // Luôn sử dụng Light Theme cho ECharts
        chart = echarts.init(chartDiv, null);

        const data = getFilteredData(item.dataSourceKey, [], item.tabId);
        const { option, error } = evaluateEChartsCode(item.code, data, { tabId: item.tabId }, chart);

        if (error) {
          chartDiv.innerHTML = `<div class="db-chart-error-banner" style="color:#ef4444;padding:12px;text-align:center;font-size:12px;">⚠️ Lỗi render ECharts:<br/>${escapeHtml(error)}<br/><button class="db-btn-tool primary" style="margin-top:8px;" onclick="window.ChartStudioModal && window.ChartStudioModal.open('${item.id}')">Mở Studio Sửa Lỗi</button></div>`;
        } else if (option && Object.keys(option).length > 0) {
          chart.setOption(option, true);
        }
      }, 50);
    } else if (item.type === 'shape') {
      const sType = item.shapeType || 'rectangle';
      const bg = item.bgColor || '#ffffff';
      const borderW = item.borderWidth !== undefined ? item.borderWidth : 1;
      const borderS = item.borderStyle || 'solid';
      const borderC = item.borderColor || '#cbd5e1';
      const border = borderW > 0 && borderS !== 'none' ? `${borderW}px ${borderS} ${borderC}` : 'none';
      const opacity = item.opacity !== undefined ? (item.opacity / 100) : 1;
      const shadow = item.boxShadow || 'none';

      let radiusCss = `${item.borderRadius !== undefined ? item.borderRadius : 8}px`;
      if (item.radiusTL !== undefined) {
        radiusCss = `${item.radiusTL || 0}px ${item.radiusTR || 0}px ${item.radiusBR || 0}px ${item.radiusBL || 0}px`;
      }
      if (sType === 'circle') radiusCss = '50%';

      if (sType === 'line') {
        const thickness = item.borderWidth || 2;
        const color = item.borderColor || '#94a3b8';
        const style = item.borderStyle || 'solid';
        const cap = item.lineCap === 'round' ? '999px' : '0px';
        const isVert = item.lineOrientation === 'vertical';

        if (isVert) {
          body.innerHTML = `
            <div style="width:100%;height:100%;display:flex;justify-content:center;align-items:center;padding:4px 0;">
              <div style="width:${thickness}px;height:100%;background:${style === 'solid' ? color : 'transparent'};border-left:${style !== 'solid' ? `${thickness}px ${style} ${color}` : 'none'};border-radius:${cap};box-shadow:${shadow};opacity:${opacity};"></div>
            </div>
          `;
        } else {
          body.innerHTML = `
            <div style="width:100%;height:100%;display:flex;align-items:center;padding:0 4px;">
              <div style="width:100%;height:${thickness}px;background:${style === 'solid' ? color : 'transparent'};border-top:${style !== 'solid' ? `${thickness}px ${style} ${color}` : 'none'};border-radius:${cap};box-shadow:${shadow};opacity:${opacity};"></div>
            </div>
          `;
        }
      } else {
        body.innerHTML = `
          <div class="db-shape-container" style="background:${bg};border:${border};border-radius:${radiusCss};opacity:${opacity};box-shadow:${shadow};width:100%;height:100%;box-sizing:border-box;"></div>
        `;
      }
    } else if (item.type === 'text') {
      const fFamily = item.fontFamily || 'Inter, sans-serif';
      const fSize = item.fontSize || '14px';
      const fWeight = item.fontWeight || 'normal';
      const fStyle = item.fontStyle || 'normal';
      const tDeco = item.textDecoration || 'none';
      const tColor = item.color || '#0f172a';
      const tAlign = item.textAlign || 'left';
      const textVal = item.text !== undefined ? item.text : (item.name || 'Nhập nội dung ghi chú...');

      body.innerHTML = `
        <div style="height:100%;width:100%;padding:4px;overflow:hidden;display:flex;flex-direction:column;justify-content:flex-start;">
          <div class="db-text-container db-text-content-editable" data-item-id="${item.id}" style="font-family:${fFamily};font-size:${fSize};font-weight:${fWeight};font-style:${fStyle};text-decoration:${tDeco};color:${tColor};text-align:${tAlign};white-space:pre-wrap;word-break:break-word;line-height:1.5;width:100%;min-height:100%;" title="Nhấp đúp để sửa chữ trực tiếp">${escapeHtml(textVal)}</div>
        </div>
      `;
      const textEl = body.querySelector('.db-text-content-editable');
      textEl?.addEventListener('dblclick', (e) => {
        e.stopPropagation();
        CanvaTextManager.startInlineEdit(item, textEl);
      });
    } else if (item.type === 'button') {
      const btnEl = document.createElement('div');
      btnEl.className = 'db-action-btn-widget';
      btnEl.style.cssText = `background:${item.btnBg || '#004ea8'};color:${item.btnColor || '#ffffff'};border-radius:${item.borderRadius || 8}px;font-size:${item.fontSize || '14px'};font-weight:700;box-shadow:0 2px 6px rgba(0,0,0,0.08);`;
      btnEl.innerHTML = `<span>${escapeHtml(item.name || 'Nút Bấm')}</span>`;
      btnEl.addEventListener('click', () => {
        if (DashboardState.mode === 'preview') {
          if (item.btnAction === 'url' && item.btnTarget) {
            window.open(item.btnTarget, '_blank');
          } else if (item.btnAction === 'tab' && item.btnTarget) {
            switchDashboardTab(item.btnTarget);
          }
        } else {
          showToast(`Nhấp nút: ${item.btnAction === 'url' ? 'Mở link ' + item.btnTarget : 'Chuyển sang Tab'} (Bật "👁️ Xem Trước" để chạy thử)`, 'info');
        }
      });
      body.appendChild(btnEl);
    } else if (item.type === 'image') {
      body.innerHTML = `
        <div class="db-image-widget">
          <img src="${item.imageUrl || ''}" alt="${escapeHtml(item.name || 'Ảnh')}" style="border-radius:${item.borderRadius || 6}px;object-fit:${item.fit || 'contain'};" />
        </div>
      `;
    } else if (item.type === 'filter') {
      const col = item.filterColumn;
      if (!item.dataSourceKey && col) {
        const sources = DataStore.getAvailableSources();
        const matched = sources.find(s => DataStore.getColumns(s.key).includes(col));
        if (matched) item.dataSourceKey = matched.key;
      }
      const allRows = DataStore.getData(item.dataSourceKey, col);
      const isDateCol = isProbableDateColumn(col, allRows);

      // 1. Phân định Loại Bộ Lọc (Chuẩn BI: range, comparison, select, multiselect, date, text)
      let fType = item.filterType;
      if (!fType) {
        if (item.operator === 'between' || item.operator === 'range') fType = 'range';
        else if (item.operator === 'in') fType = 'multiselect';
        else if (item.operator === 'date_range' || String(item.operator).startsWith('date_') || isDateCol) fType = 'date';
        else if (['contains', 'starts_with', 'ends_with', 'not_contains'].includes(item.operator)) fType = 'text';
        else if (['>', '>=', '<', '<=', '!='].includes(item.operator)) fType = 'comparison';
        else fType = isDateCol ? 'date' : 'range';
      }

      let parentItem = null;
      if (item.parentFilterId) {
        parentItem = DashboardState.present.items.find(i => i.id === item.parentFilterId);
      }

      const itemTabId = item.tabId || DashboardState.present?.activeTabId || 'default';
      const tabFiltersMap = DashboardState.tabFilters?.[itemTabId] || {};
      const filterState = tabFiltersMap[col] !== undefined
        ? tabFiltersMap[col]
        : (itemTabId === (DashboardState.present?.activeTabId || 'default') ? DashboardState.activeFilters[col] : undefined);

      const setTabFilterState = (valObj, keepDropdownOpen = false) => {
        if (!DashboardState.tabFilters) DashboardState.tabFilters = {};
        if (!DashboardState.tabFilters[itemTabId]) DashboardState.tabFilters[itemTabId] = {};
        const isCurrentActiveTab = itemTabId === (DashboardState.present?.activeTabId || 'default');

        if (valObj === null || valObj === undefined) {
          delete DashboardState.tabFilters[itemTabId][col];
          if (isCurrentActiveTab) {
            delete DashboardState.activeFilters[col];
          }
        } else {
          DashboardState.tabFilters[itemTabId][col] = valObj;
          if (isCurrentActiveTab) {
            DashboardState.activeFilters[col] = valObj;
          }
        }
        if (DashboardState.present) {
          DashboardState.present.tabFilters = JSON.parse(JSON.stringify(DashboardState.tabFilters));
          DashboardState.present.activeFilters = JSON.parse(JSON.stringify(DashboardState.activeFilters));
        }
        if (keepDropdownOpen) {
          DashboardState._openDropdownFilterId = item.id;
        } else {
          DashboardState._openDropdownFilterId = null;
        }
        renderCanvas();
      };

      let currentOp = item.operator;
      if (!currentOp) {
        if (fType === 'range') currentOp = 'between';
        else if (fType === 'multiselect') currentOp = 'in';
        else if (fType === 'date') currentOp = 'date_range';
        else if (fType === 'text') currentOp = 'contains';
        else if (fType === 'comparison') currentOp = '>';
        else currentOp = '=';
      }

      let currentVal = 'all';
      let minVal = '';
      let maxVal = '';
      let startDate = '';
      let endDate = '';
      let selectedValues = [];

      if (filterState) {
        if (typeof filterState === 'object' && filterState !== null) {
          currentOp = filterState.op || currentOp;
          currentVal = filterState.val !== undefined ? filterState.val : 'all';
          minVal = filterState.min !== undefined ? filterState.min : (filterState.start !== undefined ? filterState.start : '');
          maxVal = filterState.max !== undefined ? filterState.max : (filterState.end !== undefined ? filterState.end : '');
          startDate = filterState.start || filterState.min || '';
          endDate = filterState.end || filterState.max || '';
          selectedValues = Array.isArray(filterState.values) ? filterState.values : (Array.isArray(filterState.val) ? filterState.val : (filterState.val && filterState.val !== 'all' ? [filterState.val] : []));
        } else {
          currentVal = filterState;
        }
      }

      // Danh sách tùy chọn của bộ lọc: độc lập theo toàn bộ bảng, chỉ lọc phân cấp khi có cấu hình bộ lọc cha (parentFilterId)
      let filterSourceRows = allRows;
      if (parentItem && parentItem.filterColumn) {
        const parentCol = parentItem.filterColumn;
        const parentVal = (DashboardState.tabFilters?.[itemTabId] || {})[parentCol] || DashboardState.activeFilters[parentCol];
        if (parentVal) {
          filterSourceRows = getFilteredData(item.dataSourceKey, Object.keys(DashboardState.activeFilters).filter(k => k !== parentCol), itemTabId);
        }
      }
      const uniqueVals = Array.from(new Set((filterSourceRows || allRows).map(r => r[col]).filter(v => v !== null && v !== undefined && v !== ''))).sort((a, b) => {
        const na = Number(a), nb = Number(b);
        if (!isNaN(na) && !isNaN(nb)) return na - nb;
        return String(a).localeCompare(String(b));
      });

      // Luôn đảm bảo các mục người dùng đã chọn xuất hiện trong danh sách lựa chọn
      if (Array.isArray(selectedValues)) {
        selectedValues.forEach(sv => {
          if (sv !== null && sv !== undefined && sv !== '' && !uniqueVals.map(String).includes(String(sv))) {
            uniqueVals.push(sv);
          }
        });
      }

      // Tự động reset nếu giá trị select đơn không còn trong tập giá trị của bộ lọc cha
      if (fType === 'select' && currentVal !== 'all' && currentVal !== '' && typeof currentVal === 'string') {
        if (uniqueVals.length > 0 && !uniqueVals.some(v => String(v) === String(currentVal))) {
          if (DashboardState.tabFilters?.[itemTabId]) {
            delete DashboardState.tabFilters[itemTabId][col];
          }
          if (itemTabId === (DashboardState.present?.activeTabId || 'default')) {
            delete DashboardState.activeFilters[col];
          }
          currentVal = 'all';
        }
      }

      const isDate = fType === 'date' || (fType === 'range' && isDateCol);
      const isNumberCol = !isDate && uniqueVals.some(v => !isNaN(Number(v)) && String(v).trim() !== '');

      // 2. Xác định trạng thái đang lọc & tạo Badge tóm tắt
      let isFiltering = false;
      let activeBadgeHtml = '';

      if (fType === 'range') {
        if (isDate) {
          isFiltering = Boolean(startDate || endDate);
          let bStart = startDate;
          let bEnd = endDate;
          if (startDate && endDate) {
            const sTs = parseDateToTimestamp(startDate);
            const eTs = parseDateToTimestamp(endDate);
            if (sTs !== null && eTs !== null && sTs > eTs) {
              bStart = endDate;
              bEnd = startDate;
            }
          }
          if (isFiltering) activeBadgeHtml = `<span class="db-filter-active-badge">📅 [${bStart || '...'} ➔ ${bEnd || '...'}]</span>`;
        } else {
          isFiltering = Boolean(minVal !== '' || maxVal !== '');
          let bMin = minVal;
          let bMax = maxVal;
          if (minVal !== '' && maxVal !== '' && !isNaN(Number(minVal)) && !isNaN(Number(maxVal)) && Number(minVal) > Number(maxVal)) {
            bMin = maxVal;
            bMax = minVal;
          }
          if (isFiltering) activeBadgeHtml = `<span class="db-filter-active-badge">↔️ [${bMin !== '' ? bMin : '-∞'} ➔ ${bMax !== '' ? bMax : '+∞'}]</span>`;
        }
      } else if (fType === 'multiselect') {
        isFiltering = Boolean(selectedValues.length > 0 && selectedValues.length < uniqueVals.length);
        if (isFiltering) activeBadgeHtml = `<span class="db-filter-active-badge">☑️ (${selectedValues.length} mục)</span>`;
      } else if (fType === 'date') {
        if (currentOp === 'date_range') {
          isFiltering = Boolean(startDate || endDate);
          let bStart = startDate;
          let bEnd = endDate;
          if (startDate && endDate) {
            const sTs = parseDateToTimestamp(startDate);
            const eTs = parseDateToTimestamp(endDate);
            if (sTs !== null && eTs !== null && sTs > eTs) {
              bStart = endDate;
              bEnd = startDate;
            }
          }
          if (isFiltering) activeBadgeHtml = `<span class="db-filter-active-badge">📅 [${bStart || '...'} ➔ ${bEnd || '...'}]</span>`;
        } else {
          isFiltering = Boolean(currentVal !== 'all' && currentVal !== '');
          if (isFiltering) activeBadgeHtml = `<span class="db-filter-active-badge">📅 ${currentOp} ${escapeHtml(currentVal)}</span>`;
        }
      } else {
        isFiltering = Boolean(currentVal !== 'all' && currentVal !== '' && currentVal !== null && currentVal !== undefined);
        if (isFiltering) activeBadgeHtml = `<span class="db-filter-active-badge">${escapeHtml(currentOp)} ${escapeHtml(currentVal)}</span>`;
      }

      let parentBadgeHtml = '';
      if (parentItem) {
        parentBadgeHtml = `<span class="db-filter-parent-badge" title="Bộ lọc phụ thuộc vào: ${escapeHtml(parentItem.name || parentItem.filterColumn)}">↳ Con: ${escapeHtml(parentItem.name || parentItem.filterColumn)}</span>`;
      }

      // 3. Render giao diện các ô điều khiển theo từng Loại Bộ Lọc (BI Platform)
      let controlsHtml = '';

      if (fType === 'range') {
        // [LOẠI 1]: TRONG KHOẢNG (BETWEEN: Từ ... Đến ...)
        if (isDate) {
          controlsHtml = `
            <div class="db-filter-date-range-wrap">
              <input type="date" class="db-filter-date-input" id="filter_date_start_${item.id}" value="${formatDateToYMD(startDate)}" title="Từ ngày" />
              <span class="db-filter-date-sep">➔</span>
              <input type="date" class="db-filter-date-input" id="filter_date_end_${item.id}" value="${formatDateToYMD(endDate)}" title="Đến ngày" />
            </div>
          `;
        } else {
          controlsHtml = `
            <div class="db-filter-range-wrap">
              <input type="${isNumberCol ? 'number' : 'text'}" 
                class="db-filter-range-input" 
                id="filter_range_min_${item.id}" 
                placeholder="Từ (Min)..." 
                value="${escapeHtml(minVal)}" 
                title="Giá trị nhỏ nhất" />
              <span class="db-filter-range-sep">➔</span>
              <input type="${isNumberCol ? 'number' : 'text'}" 
                class="db-filter-range-input" 
                id="filter_range_max_${item.id}" 
                placeholder="Đến (Max)..." 
                value="${escapeHtml(maxVal)}" 
                title="Giá trị lớn nhất" />
            </div>
          `;
        }
      } else if (fType === 'comparison') {
        // [LOẠI 2]: SO SÁNH 1 GIÁ TRỊ (SINGLE VALUE: >, >=, <, <=, =, !=)
        controlsHtml = `
          <select class="db-filter-op-select" id="filter_op_${item.id}" title="Chọn phép so sánh" style="width:84px;">
            <option value=">" ${currentOp === '>' ? 'selected' : ''}>&gt; Lớn hơn</option>
            <option value=">=" ${currentOp === '>=' ? 'selected' : ''}>&ge; Lớn/bằng</option>
            <option value="<" ${currentOp === '<' ? 'selected' : ''}>&lt; Bé hơn</option>
            <option value="<=" ${currentOp === '<=' ? 'selected' : ''}>&le; Bé/bằng</option>
            <option value="=" ${currentOp === '=' ? 'selected' : ''}>= Bằng</option>
            <option value="!=" ${currentOp === '!=' ? 'selected' : ''}>≠ Khác</option>
          </select>
          <div style="flex:1;display:flex;gap:4px;min-width:0;">
            <input type="${isNumberCol ? 'number' : 'text'}" 
              class="db-filter-val-input" 
              id="filter_val_${item.id}" 
              placeholder="${isNumberCol ? 'Nhập số mốc...' : 'Giá trị...'}" 
              value="${currentVal !== 'all' ? escapeHtml(currentVal) : ''}" 
              list="filter_datalist_${item.id}" />
            <datalist id="filter_datalist_${item.id}">
              ${uniqueVals.slice(0, 30).map(v => `<option value="${escapeHtml(v)}"></option>`).join('')}
            </datalist>
          </div>
        `;
      } else if (fType === 'multiselect') {
        // [LOẠI 3]: CHỌN NHIỀU MỤC (MULTI-SELECT CHECKBOXES)
        const countLabel = selectedValues.length === 0 ? `Tất cả (${uniqueVals.length})` : `Đã chọn (${selectedValues.length})`;
        controlsHtml = `
          <div class="db-filter-multiselect-container" id="filter_ms_container_${item.id}">
            <button type="button" class="db-filter-multiselect-btn" id="filter_ms_btn_${item.id}" title="Nhấp để chọn một hoặc nhiều mục">
              <span id="filter_ms_label_${item.id}">☑️ ${escapeHtml(countLabel)}</span>
              <span>▾</span>
            </button>
            <div class="db-filter-multiselect-dropdown" id="filter_ms_drop_${item.id}">
              <input type="text" class="db-filter-multiselect-search" id="filter_ms_search_${item.id}" placeholder="🔍 Tìm giá trị..." />
              <div class="db-filter-multiselect-actions">
                <button type="button" class="db-filter-ms-action-btn" id="filter_ms_all_${item.id}">Chọn tất cả</button>
                <button type="button" class="db-filter-ms-action-btn" id="filter_ms_none_${item.id}">Bỏ chọn</button>
              </div>
              <div class="db-filter-multiselect-list" id="filter_ms_list_${item.id}">
                ${uniqueVals.map(v => `
                  <label class="db-filter-multiselect-item">
                    <input type="checkbox" value="${escapeHtml(v)}" ${selectedValues.map(String).includes(String(v)) ? 'checked' : ''} class="filter_ms_chk_${item.id}" />
                    <span>${escapeHtml(v)}</span>
                  </label>
                `).join('')}
              </div>
            </div>
          </div>
        `;
      } else if (fType === 'text') {
        // [LOẠI 4]: TÌM KIẾM VĂN BẢN (TEXT SEARCH)
        controlsHtml = `
          <select class="db-filter-op-select" id="filter_op_${item.id}" title="Chọn điều kiện tìm kiếm" style="width:84px;">
            <option value="contains" ${currentOp === 'contains' ? 'selected' : ''}>Chứa</option>
            <option value="starts_with" ${currentOp === 'starts_with' ? 'selected' : ''}>Bắt đầu</option>
            <option value="ends_with" ${currentOp === 'ends_with' ? 'selected' : ''}>Kết thúc</option>
            <option value="not_contains" ${currentOp === 'not_contains' ? 'selected' : ''}>K.chứa</option>
            <option value="=" ${currentOp === '=' ? 'selected' : ''}>= Chuẩn</option>
          </select>
          <div style="flex:1;display:flex;gap:4px;min-width:0;">
            <input type="text" 
              class="db-filter-val-input" 
              id="filter_val_${item.id}" 
              placeholder="Nhập từ khóa..." 
              value="${currentVal !== 'all' ? escapeHtml(currentVal) : ''}" />
          </div>
        `;
      } else if (fType === 'date') {
        // [LOẠI 5]: BỘ LỌC NGÀY THÁNG (DATE PICKER)
        controlsHtml = `
          <select class="db-filter-op-select" id="filter_op_${item.id}" title="Chọn kiểu lọc ngày" style="width:105px;">
            <option value="date_range" ${currentOp === 'date_range' ? 'selected' : ''}>📅 Khoảng</option>
            <option value="date_eq" ${currentOp === 'date_eq' ? 'selected' : ''}>🎯 Đúng ngày</option>
            <option value="date_gte" ${currentOp === 'date_gte' ? 'selected' : ''}>⏩ Từ ngày</option>
            <option value="date_lte" ${currentOp === 'date_lte' ? 'selected' : ''}>⏪ Đến ngày</option>
          </select>
          ${currentOp === 'date_range' ? `
            <div class="db-filter-date-range-wrap">
              <input type="date" class="db-filter-date-input" id="filter_date_start_${item.id}" value="${formatDateToYMD(startDate)}" title="Từ ngày" />
              <span class="db-filter-date-sep">➔</span>
              <input type="date" class="db-filter-date-input" id="filter_date_end_${item.id}" value="${formatDateToYMD(endDate)}" title="Đến ngày" />
            </div>
          ` : `
            <div class="db-filter-date-range-wrap">
              <input type="date" class="db-filter-date-input" id="filter_date_val_${item.id}" value="${formatDateToYMD(currentVal !== 'all' ? currentVal : '')}" title="Chọn ngày" />
            </div>
          `}
        `;
      } else {
        // [LOẠI 6]: DANH SÁCH CHỌN 1 (SINGLE SELECT DROPDOWN)
        controlsHtml = `
          <select class="db-filter-val-select" id="filter_val_${item.id}" title="Chọn giá trị từ Excel">
            <option value="all">-- Tất cả (${uniqueVals.length}) --</option>
            ${uniqueVals.map(v => `<option value="${escapeHtml(v)}" ${String(v) === String(currentVal) ? 'selected' : ''}>${escapeHtml(v)}</option>`).join('')}
          </select>
        `;
      }

      body.innerHTML = `
        <div class="db-filter-container">
          <div class="db-filter-header-row">
            <span class="db-filter-label" title="${escapeHtml(item.name || col)}">${escapeHtml(item.name || col)}</span>
            <div style="display:flex;align-items:center;gap:4px;flex-shrink:0;">
              ${parentBadgeHtml}
              ${activeBadgeHtml}
              ${isFiltering ? `<button class="db-filter-reset-btn" id="filter_reset_${item.id}" title="Xóa bỏ điều kiện lọc này">✕</button>` : ''}
            </div>
          </div>
          <div class="db-filter-controls-row">
            ${controlsHtml}
          </div>
        </div>
      `;

      setTimeout(() => {
        const resetBtn = document.getElementById(`filter_reset_${item.id}`);
        resetBtn?.addEventListener('click', (e) => {
          e.stopPropagation();
          setTabFilterState(null);
        });

        if (fType === 'range') {
          if (isDate) {
            const startInp = document.getElementById(`filter_date_start_${item.id}`);
            const endInp = document.getElementById(`filter_date_end_${item.id}`);
            const handleDateRange = () => {
              const sVal = startInp ? startInp.value : '';
              const eVal = endInp ? endInp.value : '';
              if (!sVal && !eVal) {
                setTabFilterState(null);
              } else {
                setTabFilterState({ op: 'date_range', start: sVal, end: eVal, isDate: true });
              }
            };
            startInp?.addEventListener('change', handleDateRange);
            endInp?.addEventListener('change', handleDateRange);
          } else {
            const minInp = document.getElementById(`filter_range_min_${item.id}`);
            const maxInp = document.getElementById(`filter_range_max_${item.id}`);
            const handleRange = () => {
              const minV = minInp ? minInp.value.trim() : '';
              const maxV = maxInp ? maxInp.value.trim() : '';
              if (minV === '' && maxV === '') {
                setTabFilterState(null);
              } else {
                setTabFilterState({ op: 'between', min: minV, max: maxV, start: minV, end: maxV });
              }
            };
            minInp?.addEventListener('change', handleRange);
            maxInp?.addEventListener('change', handleRange);
            minInp?.addEventListener('keydown', (e) => { if (e.key === 'Enter') handleRange(); });
            maxInp?.addEventListener('keydown', (e) => { if (e.key === 'Enter') handleRange(); });
          }
        } else if (fType === 'comparison') {
          const opSel = document.getElementById(`filter_op_${item.id}`);
          const valEl = document.getElementById(`filter_val_${item.id}`);
          const applyComp = () => {
            const op = opSel ? opSel.value : '=';
            item.operator = op;
            const val = valEl ? valEl.value.trim() : '';
            if (val === '' || val === 'all') {
              setTabFilterState(null);
            } else {
              setTabFilterState({ op, val });
            }
          };
          opSel?.addEventListener('change', applyComp);
          valEl?.addEventListener('change', applyComp);
          valEl?.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
              e.preventDefault();
              valEl.blur();
              applyComp();
            }
          });
        } else if (fType === 'multiselect') {
          const btn = document.getElementById(`filter_ms_btn_${item.id}`);
          const drop = document.getElementById(`filter_ms_drop_${item.id}`);
          const searchInp = document.getElementById(`filter_ms_search_${item.id}`);
          const chks = document.querySelectorAll(`.filter_ms_chk_${item.id}`);
          const btnAll = document.getElementById(`filter_ms_all_${item.id}`);
          const btnNone = document.getElementById(`filter_ms_none_${item.id}`);

          // Khôi phục trạng thái mở nếu trước đó dropdown này vừa được tương tác
          if (DashboardState._openDropdownFilterId === item.id && drop) {
            drop.classList.add('open');
            const parentBox = drop.closest('.db-item-box') || drop.closest('.db-top-filter-item-wrapper');
            if (parentBox) {
              if (parentBox.dataset.originalZIndex === undefined) {
                parentBox.dataset.originalZIndex = parentBox.style.zIndex || '';
              }
              parentBox.classList.add('dropdown-open');
              parentBox.style.zIndex = '1000000';
            }
            const topBar = drop.closest('.db-top-filter-bar');
            if (topBar) {
              if (topBar.dataset.originalZIndex === undefined) {
                topBar.dataset.originalZIndex = topBar.style.zIndex || '';
              }
              topBar.classList.add('dropdown-open');
              topBar.style.zIndex = '1000000';
            }
          }

          btn?.addEventListener('click', (e) => {
            e.stopPropagation();
            document.querySelectorAll('.db-filter-multiselect-dropdown.open').forEach(d => {
              if (d !== drop) {
                d.classList.remove('open');
                const pBox = d.closest('.db-item-box') || d.closest('.db-top-filter-item-wrapper');
                if (pBox) {
                  pBox.classList.remove('dropdown-open');
                  if (pBox.dataset.originalZIndex !== undefined) {
                    pBox.style.zIndex = pBox.dataset.originalZIndex;
                  } else {
                    pBox.style.zIndex = '';
                  }
                }
                const tb = d.closest('.db-top-filter-bar');
                if (tb) {
                  tb.classList.remove('dropdown-open');
                  if (tb.dataset.originalZIndex !== undefined) {
                    tb.style.zIndex = tb.dataset.originalZIndex;
                  } else {
                    tb.style.zIndex = '';
                  }
                }
              }
            });

            const isOpen = drop?.classList.toggle('open');
            const parentBox = drop?.closest('.db-item-box') || drop?.closest('.db-top-filter-item-wrapper');
            const topBar = drop?.closest('.db-top-filter-bar');
            if (parentBox) {
              if (isOpen) {
                if (parentBox.dataset.originalZIndex === undefined) {
                  parentBox.dataset.originalZIndex = parentBox.style.zIndex || '';
                }
                parentBox.classList.add('dropdown-open');
                parentBox.style.zIndex = '1000000';
                DashboardState._openDropdownFilterId = item.id;
              } else {
                parentBox.classList.remove('dropdown-open');
                if (parentBox.dataset.originalZIndex !== undefined) {
                  parentBox.style.zIndex = parentBox.dataset.originalZIndex;
                } else {
                  parentBox.style.zIndex = '';
                }
                DashboardState._openDropdownFilterId = null;
              }
            }
            if (topBar) {
              if (isOpen) {
                if (topBar.dataset.originalZIndex === undefined) {
                  topBar.dataset.originalZIndex = topBar.style.zIndex || '';
                }
                topBar.classList.add('dropdown-open');
                topBar.style.zIndex = '1000000';
              } else {
                topBar.classList.remove('dropdown-open');
                if (topBar.dataset.originalZIndex !== undefined) {
                  topBar.style.zIndex = topBar.dataset.originalZIndex;
                } else {
                  topBar.style.zIndex = '';
                }
              }
            }

            if (drop?.classList.contains('open')) {
              searchInp?.focus();
            }
          });

          drop?.addEventListener('click', (e) => e.stopPropagation());

          searchInp?.addEventListener('input', (e) => {
            const term = e.target.value.toLowerCase().trim();
            document.querySelectorAll(`#filter_ms_list_${item.id} .db-filter-multiselect-item`).forEach(label => {
              const text = label.textContent.toLowerCase();
              label.style.display = text.includes(term) ? 'flex' : 'none';
            });
          });

          const applyMultiSelect = () => {
            const checked = Array.from(document.querySelectorAll(`.filter_ms_chk_${item.id}:checked`)).map(c => c.value);
            if (checked.length === 0 || checked.length === uniqueVals.length) {
              setTabFilterState(null, true);
            } else {
              setTabFilterState({ op: 'in', values: checked, val: checked }, true);
            }
          };

          chks.forEach(c => c.addEventListener('change', applyMultiSelect));

          btnAll?.addEventListener('click', () => {
            setTabFilterState(null, true);
          });

          btnNone?.addEventListener('click', () => {
            setTabFilterState({ op: 'in', values: [], val: [] }, true);
          });
        } else if (fType === 'text') {
          const opSel = document.getElementById(`filter_op_${item.id}`);
          const valEl = document.getElementById(`filter_val_${item.id}`);
          const applyText = () => {
            const op = opSel ? opSel.value : 'contains';
            item.operator = op;
            const val = valEl ? valEl.value.trim() : '';
            if (val === '' || val === 'all') {
              setTabFilterState(null);
            } else {
              setTabFilterState({ op, val });
            }
          };
          opSel?.addEventListener('change', applyText);
          valEl?.addEventListener('change', applyText);
          valEl?.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
              e.preventDefault();
              valEl.blur();
              applyText();
            }
          });
        } else if (fType === 'date') {
          const opSel = document.getElementById(`filter_op_${item.id}`);
          const startInp = document.getElementById(`filter_date_start_${item.id}`);
          const endInp = document.getElementById(`filter_date_end_${item.id}`);
          const singleInp = document.getElementById(`filter_date_val_${item.id}`);

          const handleDateChange = () => {
            const op = opSel ? opSel.value : 'date_range';
            if (op === 'date_range') {
              const sVal = startInp ? startInp.value : '';
              const eVal = endInp ? endInp.value : '';
              if (!sVal && !eVal) {
                setTabFilterState(null);
              } else {
                setTabFilterState({ op: 'date_range', start: sVal, end: eVal, isDate: true });
              }
            } else {
              const dVal = singleInp ? singleInp.value : '';
              if (!dVal) {
                setTabFilterState(null);
              } else {
                setTabFilterState({ op, val: dVal, isDate: true });
              }
            }
          };

          startInp?.addEventListener('change', handleDateChange);
          endInp?.addEventListener('change', handleDateChange);
          singleInp?.addEventListener('change', handleDateChange);
          opSel?.addEventListener('change', () => {
            item.operator = opSel.value;
            handleDateChange();
          });
        } else {
          // select
          const valEl = document.getElementById(`filter_val_${item.id}`);
          valEl?.addEventListener('change', (e) => {
            const val = e.target.value;
            if (val === '' || val === 'all') {
              setTabFilterState(null);
            } else {
              setTabFilterState({ op: '=', val });
            }
          });
        }
      }, 20);
    } else if (item.type === 'kpi') {
      const data = getFilteredData(item.dataSourceKey, [], item.tabId || DashboardState.present?.activeTabId || 'default');
      const numVals = data.map(r => Number(r[item.metricColumn])).filter(n => !isNaN(n));
      let sum = 0;
      if (item.aggregation === 'avg') {
        sum = numVals.length > 0 ? (numVals.reduce((a, b) => a + b, 0) / numVals.length) : 0;
      } else if (item.aggregation === 'count') {
        sum = numVals.length;
      } else if (item.aggregation === 'max') {
        sum = numVals.length > 0 ? Math.max(...numVals) : 0;
      } else if (item.aggregation === 'min') {
        sum = numVals.length > 0 ? Math.min(...numVals) : 0;
      } else {
        sum = numVals.reduce((a, b) => a + b, 0);
      }

      const kpiFs = item.fontSize || '30px';
      const kpiColor = item.color || 'var(--db-text-main)';
      body.innerHTML = `
        <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;text-align:center;">
          <div style="font-size:11px;text-transform:uppercase;color:var(--db-text-sub);font-weight:700;margin-bottom:4px;">${escapeHtml(item.name)}</div>
          <div style="font-size:${escapeHtml(kpiFs)};font-weight:800;color:${escapeHtml(kpiColor)};">${escapeHtml(item.prefix || '')}${new Intl.NumberFormat('vi-VN').format(Math.round(sum))}${escapeHtml(item.suffix || '')}</div>
        </div>
      `;
    } else if (item.type === 'heading') {
      const fFamily = item.fontFamily || 'Inter, sans-serif';
      const fSize = item.fontSize || '24px';
      const fWeight = item.fontWeight || '700';
      const fStyle = item.fontStyle || 'normal';
      const tDeco = item.textDecoration || 'none';
      const tColor = item.color || 'var(--db-text-main)';
      const tAlign = item.textAlign || 'left';
      const textVal = item.text !== undefined ? item.text : (item.name || 'Tiêu Đề');

      body.innerHTML = `
        <div style="height:100%;width:100%;display:flex;align-items:center;padding:4px;overflow:hidden;">
          <h2 class="db-text-content-editable" data-item-id="${item.id}" style="margin:0;width:100%;font-family:${fFamily};font-size:${fSize};font-weight:${fWeight};font-style:${fStyle};text-decoration:${tDeco};color:${tColor};text-align:${tAlign};word-break:break-word;line-height:1.3;" title="Nhấp đúp để sửa chữ trực tiếp">${escapeHtml(textVal)}</h2>
        </div>
      `;
      const textEl = body.querySelector('.db-text-content-editable');
      textEl?.addEventListener('dblclick', (e) => {
        e.stopPropagation();
        CanvaTextManager.startInlineEdit(item, textEl);
      });
    } else if (item.type === 'frame') {
      const bg = item.bgColor || 'rgba(255,255,255,0.4)';
      const borderW = item.borderWidth !== undefined ? item.borderWidth : 2;
      const borderS = item.borderStyle || 'dashed';
      const borderC = item.borderColor || '#cbd5e1';
      const border = borderW > 0 && borderS !== 'none' ? `${borderW}px ${borderS} ${borderC}` : 'none';
      const radius = `${item.borderRadius !== undefined ? item.borderRadius : 12}px`;
      const shadow = item.boxShadow || 'none';
      const opacity = item.opacity !== undefined ? (item.opacity / 100) : 1;

      body.innerHTML = `
        <div class="db-frame-container" style="background:${bg};border:${border};border-radius:${radius};box-shadow:${shadow};opacity:${opacity};width:100%;height:100%;box-sizing:border-box;padding:8px;font-size:12px;color:var(--db-text-sub);font-style:italic;">
          [Khung phân vùng: ${escapeHtml(item.name)}]
        </div>
      `;
    }
  }

  function attachItemDragAndResize(cardEl, item) {
    if (DashboardState.mode === 'preview') return;
    if (item.isLocked) return;

    const headerEl = cardEl.querySelector('.db-item-header');
    const miniActionsEl = cardEl.querySelector('.db-item-mini-actions');
    const dragTargets = [
      headerEl,
      cardEl.querySelector('.db-item-drag-grip'),
      cardEl.querySelector('.db-item-mini-grip'),
      miniActionsEl
    ].filter(Boolean);
    const resizeHandle = cardEl.querySelector('.db-item-resize-handle');

    dragTargets.forEach(target => {
      target.addEventListener('mousedown', (e) => {
        // Bỏ qua nếu click trúng các nút bấm thao tác trên header
        if (e.target.closest('.db-item-action-btn') || e.target.closest('.db-item-mini-btn')) return;

        // Tránh trùng lặp sự kiện giữa header và grip con
        if (e._dragHandled) return;
        e._dragHandled = true;
        e.stopPropagation();

        e.preventDefault();
        cardEl.classList.add('dragging');
        const zoom = DashboardState.zoom || 1.0;
        const preDragSnapshot = JSON.parse(JSON.stringify(DashboardState.present));
        const startMouseX = e.clientX;
        const startMouseY = e.clientY;
        const origX = item.pos.x;
        const origY = item.pos.y;

        function onMouseMove(ev) {
          const dx = (ev.clientX - startMouseX) / zoom;
          const dy = (ev.clientY - startMouseY) / zoom;
          const rawX = origX + dx;
          const rawY = origY + dy;
          const { snappedX, snappedY, guides } = calculateSnapping(
            { x: rawX, y: rawY, w: item.pos.w, h: item.pos.h },
            item.id
          );
          showSnappingGuides(guides);

          // Ràng buộc trong 1 ranh giới duy nhất của Dashboard (Top: 24px, Left: 24px, Right: 1336px)
          const minX = 24;
          const maxX = Math.max(minX, 1360 - item.pos.w - 24);
          const minY = 24;
          const newX = Math.max(minX, Math.min(maxX, snappedX));
          const newY = Math.max(minY, snappedY);
          cardEl.style.left = `${newX}px`;
          cardEl.style.top = `${newY}px`;
          item.pos.x = newX;
          item.pos.y = newY;
          item._tempX = newX;
          item._tempY = newY;

          // Khi kéo thành phần xuống sát hoặc vượt qua mép dưới, ranh giới dưới tự động mở rộng thêm một đoạn
          const canvas = document.getElementById('dashboardCanvas');
          if (canvas) {
            const itemBottom = newY + item.pos.h + 80;
            const curH = parseInt(canvas.style.minHeight || '800', 10);
            if (itemBottom > curH) {
              canvas.style.minHeight = `${itemBottom + 200}px`;
            }
          }

          // Tự động nhận diện khi kéo Filter lướt qua thanh top-filter-bar
          if (item.type === 'filter') {
            const topBar = document.getElementById('dbTopFilterBar');
            if (topBar) {
              const barRect = topBar.getBoundingClientRect();
              const isOver = ev.clientY >= (barRect.top - 20) && ev.clientY <= (barRect.bottom + 30);
              topBar.classList.toggle('drag-over', isOver);
            }
          }
        }

        function onMouseUp() {
          clearSnappingGuides();
          window.removeEventListener('mousemove', onMouseMove);
          window.removeEventListener('mouseup', onMouseUp);
          cardEl.classList.remove('dragging');

          const topBar = document.getElementById('dbTopFilterBar');
          const isOverTopBar = topBar && topBar.classList.contains('drag-over');
          if (topBar) topBar.classList.remove('drag-over');

          // Nếu thả filter vào vùng thanh trên cùng -> Tự động gán/ghim lên thanh cố định!
          if (item.type === 'filter' && isOverTopBar) {
            delete item._tempX;
            delete item._tempY;
            pinFilterToTopBar(item);
            triggerDebouncedAutoSave();
            return;
          }

          if (item._tempX !== undefined && item._tempY !== undefined) {
            const hasMoved = (item._tempX !== origX || item._tempY !== origY);
            item.pos.x = item._tempX;
            item.pos.y = item._tempY;
            delete item._tempX;
            delete item._tempY;
            if (hasMoved) {
              DashboardState.past.push(preDragSnapshot);
              if (DashboardState.past.length > 30) DashboardState.past.shift();
              DashboardState.future = [];
              DashboardState.isDirty = true;
              DashboardState.present.updatedAt = new Date().toISOString();
              DashboardState.updateUI();
              triggerDebouncedAutoSave();
            }
          }

          if (DashboardState.selectedItemId === item.id) {
            if (item.type === 'heading' || item.type === 'text') {
              CanvaTextManager.show(item, cardEl);
            }
            if (typeof FigmaPropertiesSidebar !== 'undefined') {
              FigmaPropertiesSidebar.syncUI(item);
            }
          }
        }

        window.addEventListener('mousemove', onMouseMove);
        window.addEventListener('mouseup', onMouseUp);
      });
    });

    if (resizeHandle) {
      resizeHandle.addEventListener('mousedown', (e) => {
        e.preventDefault();
        e.stopPropagation();
        const zoom = DashboardState.zoom || 1.0;
        const preResizeSnapshot = JSON.parse(JSON.stringify(DashboardState.present));
        const startMouseX = e.clientX;
        const startMouseY = e.clientY;
        const origW = item.pos.w;
        const origH = item.pos.h;

        function onMouseMove(ev) {
          const dx = (ev.clientX - startMouseX) / zoom;
          const dy = (ev.clientY - startMouseY) / zoom;
          const rawW = Math.max(180, Math.round((origW + dx) / 10) * 10);
          const rawH = Math.max(80, Math.round((origH + dy) / 10) * 10);

          const { snappedX, snappedY, guides } = calculateSnapping(
            { x: item.pos.x, y: item.pos.y, w: rawW, h: rawH },
            item.id
          );
          showSnappingGuides(guides);

          const maxW = Math.max(180, 1360 - item.pos.x - 20);
          const newW = Math.max(180, Math.min(maxW, rawW));
          const newH = Math.max(80, rawH);
          cardEl.style.width = `${newW}px`;
          cardEl.style.height = `${newH}px`;
          item.pos.w = newW;
          item.pos.h = newH;
          item._tempW = newW;
          item._tempH = newH;

          const chartEl = cardEl.querySelector('.db-chart-canvas');
          if (chartEl) {
            const chart = echarts.getInstanceByDom(chartEl);
            if (chart) chart.resize();
          }

          // Khi kéo co giãn kích thước thành phần xuống dưới, ranh giới dưới tự động mở rộng thêm
          const canvas = document.getElementById('dashboardCanvas');
          if (canvas) {
            const resizeBottom = item.pos.y + newH + 80;
            const curH = parseInt(canvas.style.minHeight || '800', 10);
            if (resizeBottom > curH) {
              canvas.style.minHeight = `${resizeBottom + 200}px`;
            }
          }
        }

        function onMouseUp() {
          clearSnappingGuides();
          window.removeEventListener('mousemove', onMouseMove);
          window.removeEventListener('mouseup', onMouseUp);

          if (item._tempW !== undefined && item._tempH !== undefined) {
            const hasResized = (item._tempW !== origW || item._tempH !== origH);
            item.pos.w = item._tempW;
            item.pos.h = item._tempH;
            delete item._tempW;
            delete item._tempH;
            if (hasResized) {
              DashboardState.past.push(preResizeSnapshot);
              if (DashboardState.past.length > 30) DashboardState.past.shift();
              DashboardState.future = [];
              DashboardState.isDirty = true;
              DashboardState.present.updatedAt = new Date().toISOString();
              DashboardState.updateUI();
              triggerDebouncedAutoSave();
            }
          }

          if (DashboardState.selectedItemId === item.id) {
            if (item.type === 'heading' || item.type === 'text') {
              CanvaTextManager.show(item, cardEl);
            }
            if (typeof FigmaPropertiesSidebar !== 'undefined') {
              FigmaPropertiesSidebar.syncUI(item);
            }
          }
        }

        window.addEventListener('mousemove', onMouseMove);
        window.addEventListener('mouseup', onMouseUp);
      });
    }
  }

  // ===========================================================================
  // 12.5 DASHBOARD TEMPLATES GALLERY & LOGIC
  // ===========================================================================
  const DashboardTemplates = {
    sales: {
      title: 'Báo Cáo Bán Hàng & Hoạt Động Kinh Doanh',
      description: 'Tổng hợp doanh số, tỷ trọng theo ngành hàng và bộ lọc khu vực',
      tabs: [{ id: 'tab_sales', name: 'Tổng Quan Doanh Thu' }],
      activeTabId: 'tab_sales',
      items: [
        {
          id: 'heading_title',
          type: 'heading',
          name: 'Tiêu Đề',
          text: 'Báo Cáo Doanh Thu Bán Hàng Toàn Quốc',
          fontSize: '22px',
          tabId: 'tab_sales',
          pos: { x: 30, y: 20, w: 560, h: 60 }
        },
        {
          id: 'filter_khuvuc',
          type: 'filter',
          name: 'Khu Vực',
          filterColumn: 'Khu Vực',
          tabId: 'tab_sales',
          pos: { x: 610, y: 20, w: 260, h: 60 }
        },
        {
          id: 'kpi_doanhthu',
          type: 'kpi',
          name: 'Tổng Doanh Thu',
          metricColumn: 'Doanh Thu',
          aggregation: 'sum',
          prefix: '',
          suffix: ' đ',
          tabId: 'tab_sales',
          pos: { x: 890, y: 20, w: 280, h: 60 }
        },
        {
          id: 'chart_sales_bar',
          type: 'chart',
          name: 'Doanh Thu Theo Tháng',
          dataSourceKey: 'Du_Lieu_Mau_Kinh_Doanh.xlsx::BaoCaoDoanhSo',
          code: ChartTemplates.bar('Du_Lieu_Mau_Kinh_Doanh.xlsx::BaoCaoDoanhSo'),
          tabId: 'tab_sales',
          pos: { x: 30, y: 100, w: 620, h: 360 }
        },
        {
          id: 'chart_cat_pie',
          type: 'chart',
          name: 'Tỷ Trọng Doanh Thu Theo Ngành Hàng',
          dataSourceKey: 'Du_Lieu_Mau_Kinh_Doanh.xlsx::BaoCaoDoanhSo',
          code: ChartTemplates.pie('Du_Lieu_Mau_Kinh_Doanh.xlsx::BaoCaoDoanhSo'),
          tabId: 'tab_sales',
          pos: { x: 670, y: 100, w: 500, h: 360 }
        }
      ]
    },
    finance: {
      title: 'Phân Tích Tài Chính & Tỷ Suất Lợi Nhuận',
      description: 'So sánh tương quan giữa Doanh thu, Chi phí và Lợi nhuận',
      tabs: [{ id: 'tab_finance', name: 'Tài Chính & Lợi Nhuận' }],
      activeTabId: 'tab_finance',
      items: [
        {
          id: 'heading_title',
          type: 'heading',
          name: 'Tiêu Đề',
          text: 'Phân Tích Cơ Cấu Tài Chính Doanh Nghiệp',
          fontSize: '22px',
          tabId: 'tab_finance',
          pos: { x: 30, y: 20, w: 560, h: 60 }
        },
        {
          id: 'kpi_loinhuan',
          type: 'kpi',
          name: 'Tổng Lợi Nhuận',
          metricColumn: 'Lợi Nhuận',
          aggregation: 'sum',
          prefix: '+',
          suffix: ' đ',
          tabId: 'tab_finance',
          pos: { x: 610, y: 20, w: 280, h: 60 }
        },
        {
          id: 'kpi_chiphi',
          type: 'kpi',
          name: 'Tổng Chi Phí',
          metricColumn: 'Chi Phí',
          aggregation: 'sum',
          prefix: '-',
          suffix: ' đ',
          tabId: 'tab_finance',
          pos: { x: 910, y: 20, w: 280, h: 60 }
        },
        {
          id: 'chart_finance_dual',
          type: 'chart',
          name: 'Doanh Thu vs Chi Phí & Lợi Nhuận',
          dataSourceKey: 'Du_Lieu_Mau_Kinh_Doanh.xlsx::BaoCaoDoanhSo',
          code: ChartTemplates.dual_axes('Du_Lieu_Mau_Kinh_Doanh.xlsx::BaoCaoDoanhSo'),
          tabId: 'tab_finance',
          pos: { x: 30, y: 100, w: 720, h: 380 }
        },
        {
          id: 'chart_kpi_gauge',
          type: 'chart',
          name: 'Tỷ Lệ Đạt Chỉ Tiêu Lợi Nhuận',
          dataSourceKey: 'Du_Lieu_Mau_Kinh_Doanh.xlsx::BaoCaoDoanhSo',
          code: ChartTemplates.gauge('Du_Lieu_Mau_Kinh_Doanh.xlsx::BaoCaoDoanhSo'),
          tabId: 'tab_finance',
          pos: { x: 770, y: 100, w: 420, h: 380 }
        }
      ]
    },
    operations: {
      title: 'Báo Cáo Hiệu Suất Vận Hành & Chuyển Đổi',
      description: 'Theo dõi đơn hàng, đánh giá đa chiều các khu vực và phễu bán hàng',
      tabs: [{ id: 'tab_operations', name: 'Vận Hành & Đơn Hàng' }],
      activeTabId: 'tab_operations',
      items: [
        {
          id: 'heading_title',
          type: 'heading',
          name: 'Tiêu Đề',
          text: 'Hiệu Suất Phân Phối & Tỷ Lệ Chuyển Đổi',
          fontSize: '22px',
          tabId: 'tab_operations',
          pos: { x: 30, y: 20, w: 560, h: 60 }
        },
        {
          id: 'kpi_donhang',
          type: 'kpi',
          name: 'Tổng Đơn Hàng',
          metricColumn: 'Số Đơn Hàng',
          aggregation: 'sum',
          prefix: '📦 ',
          suffix: ' đơn',
          tabId: 'tab_operations',
          pos: { x: 610, y: 20, w: 260, h: 60 }
        },
        {
          id: 'filter_khuvuc',
          type: 'filter',
          name: 'Khu Vực',
          filterColumn: 'Khu Vực',
          tabId: 'tab_operations',
          pos: { x: 890, y: 20, w: 280, h: 60 }
        },
        {
          id: 'chart_ops_radar',
          type: 'chart',
          name: 'Đánh Giá Toàn Diện Các Khu Vực',
          dataSourceKey: 'Du_Lieu_Mau_Kinh_Doanh.xlsx::BaoCaoDoanhSo',
          code: ChartTemplates.radar('Du_Lieu_Mau_Kinh_Doanh.xlsx::BaoCaoDoanhSo'),
          tabId: 'tab_operations',
          pos: { x: 30, y: 100, w: 600, h: 380 }
        },
        {
          id: 'chart_ops_funnel',
          type: 'chart',
          name: 'Phễu Chuyển Đổi Đơn Hàng',
          dataSourceKey: 'Du_Lieu_Mau_Kinh_Doanh.xlsx::BaoCaoDoanhSo',
          code: ChartTemplates.funnel('Du_Lieu_Mau_Kinh_Doanh.xlsx::BaoCaoDoanhSo'),
          tabId: 'tab_operations',
          pos: { x: 650, y: 100, w: 520, h: 380 }
        }
      ]
    },
    blank: {
      title: 'Dashboard Mới Chưa Đặt Tên',
      description: 'Dashboard canvas trống',
      tabs: [{ id: 'tab_1', name: 'Trang 1' }],
      activeTabId: 'tab_1',
      items: []
    }
  };

  function openTemplatesModal() {
    const modal = document.getElementById('templatesModal');
    if (modal) modal.classList.add('open');
  }

  function closeTemplatesModal() {
    const modal = document.getElementById('templatesModal');
    if (modal) modal.classList.remove('open');
    closeTemplatePreview();
  }

  function previewDashboardTemplate(key) {
    const tpl = DashboardTemplates[key];
    const drawer = document.getElementById('templatesPreviewDrawer');
    const content = document.getElementById('tplPreviewContent');
    const titleEl = document.getElementById('tplPreviewTitle');
    const descEl = document.getElementById('tplPreviewDesc');
    const applyBtn = document.getElementById('btnTplPreviewApply');
    if (!tpl || !drawer || !content) return;

    titleEl.textContent = `Chi Tiết Mẫu: ${tpl.title}`;
    descEl.textContent = tpl.description || 'Bố cục chuẩn với các chỉ số và đồ thị dựng sẵn';
    applyBtn.dataset.templateKey = key;

    let itemsHtml = '';
    const items = tpl.items || [];
    items.forEach(it => {
      const typeIcon = it.type === 'chart' ? '📊' : it.type === 'kpi' ? '📈' : it.type === 'filter' ? '🔍' : it.type === 'heading' ? '🔤' : '📦';
      const typeName = it.type === 'chart' ? 'Biểu Đồ ECharts' : it.type === 'kpi' ? 'Thẻ KPI' : it.type === 'filter' ? 'Bộ Lọc' : it.type === 'heading' ? 'Tiêu Đề' : it.type;
      itemsHtml += `
        <div class="db-tpl-widget-chip">
          <span>${typeIcon}</span>
          <span><strong>${escapeHtml(it.name || typeName)}</strong> (${it.pos.w} × ${it.pos.h}px)</span>
        </div>
      `;
    });

    content.innerHTML = `
      <div style="font-size:12px;font-weight:700;color:#0f172a;margin-bottom:6px;">Danh sách thành phần cấu hình (${items.length} phần tử, ${tpl.tabs?.length || 1} Tab):</div>
      <div class="db-tpl-widget-chip-list">${itemsHtml || '<span style="font-size:12px;color:#64748b;font-style:italic;">Canvas mới hoàn toàn trống</span>'}</div>
      <div style="margin-top:14px;padding:12px;background:#ffffff;border:1px dashed #cbd5e1;border-radius:8px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px;">
        <span style="font-size:12px;color:#334155;">💡 Mẫu này đã được tối ưu bố cục và mã ECharts, sẵn sàng liên kết với các sheet Excel của bạn.</span>
        <button class="db-btn-tool primary" id="btnApplyFromDrawer" style="padding:6px 16px;font-size:12px;font-weight:700;">⚡ Nạp Mẫu Này Ngay</button>
      </div>
    `;

    document.getElementById('btnApplyFromDrawer')?.addEventListener('click', () => {
      applyDashboardTemplate(key);
    });

    drawer.classList.add('open');
    drawer.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }

  function closeTemplatePreview() {
    const drawer = document.getElementById('templatesPreviewDrawer');
    if (drawer) drawer.classList.remove('open');
  }

  // ===========================================================================
  // 12.6 CANVAS BACKGROUND SWITCHER (TRẮNG, CHẤM BI, DỌC KẺ, KẺ NGANG, LƯỚI Ô VUÔNG)
  // ===========================================================================
  function setCanvasBackground(bgType) {
    const valid = ['white', 'dots', 'stripes', 'horizontal', 'grid'];
    if (!valid.includes(bgType)) bgType = 'dots';

    DashboardState.pushSnapshot();
    if (!DashboardState.present.theme) DashboardState.present.theme = {};
    DashboardState.present.theme.canvasBg = bgType;

    try {
      if (typeof localStorage !== 'undefined') {
        localStorage.setItem('mb_dashboard_canvas_bg', bgType);
      }
    } catch (_) {}

    applyCanvasBackgroundUI(bgType);

    // Lưu ngay lập tức vào catalog và session
    try {
      if (typeof SavedDashboardsCatalog !== 'undefined' && SavedDashboardsCatalog.saveCurrent) {
        SavedDashboardsCatalog.saveCurrent(false);
      }
      saveLastEditedSessionToLocalStorage();
    } catch (_) {}
    triggerDebouncedAutoSave();

    const bgNames = {
      white: 'Trắng trơn',
      dots: 'Chấm bi',
      stripes: 'Dọc kẻ',
      horizontal: 'Kẻ ngang',
      grid: 'Lưới ô vuông'
    };
    showToast(`Đã đổi và lưu nền Dashboard: ${bgNames[bgType] || bgType}!`, 'success');
  }

  function applyCanvasBackgroundUI(bgType) {
    const valid = ['white', 'dots', 'stripes', 'horizontal', 'grid'];
    if (!valid.includes(bgType)) {
      const cached = (typeof localStorage !== 'undefined' && localStorage.getItem('mb_dashboard_canvas_bg')) || null;
      bgType = (cached && valid.includes(cached)) ? cached : 'dots';
    }

    const wrapper = document.getElementById('canvasWrapper');
    const canvas = document.getElementById('dashboardCanvas');
    const classes = ['canvas-bg-white', 'canvas-bg-dots', 'canvas-bg-stripes', 'canvas-bg-horizontal', 'canvas-bg-grid'];

    if (wrapper) {
      wrapper.classList.remove(...classes);
      wrapper.classList.add(`canvas-bg-${bgType}`);
    }
    if (canvas) {
      canvas.classList.remove(...classes);
      canvas.classList.add(`canvas-bg-${bgType}`);
    }

    if (DashboardState.present) {
      if (!DashboardState.present.theme) DashboardState.present.theme = {};
      DashboardState.present.theme.canvasBg = bgType;
    }

    try {
      if (typeof localStorage !== 'undefined') {
        localStorage.setItem('mb_dashboard_canvas_bg', bgType);
      }
    } catch (_) {}

    // Đánh dấu active trong menu dropdown nếu có
    document.querySelectorAll('#dropdownCanvasBg [data-bg-type]').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.bgType === bgType);
    });
  }

  async function applyDashboardTemplate(key) {
    const tpl = DashboardTemplates[key];
    if (!tpl) return;
    if (DashboardState.isDirty) {
      const ok = await AppDialog.confirm('Bạn có thay đổi chưa lưu. Bạn có chắc muốn nạp mẫu mới này không?', 'Nạp Mẫu Dashboard', { isDanger: true, confirmText: 'Nạp Mẫu' });
      if (!ok) return;
    }
    DashboardState.pushSnapshot();
    DashboardState.present = {
      id: 'dashboard_' + Date.now(),
      title: tpl.title,
      description: tpl.description || '',
      createdAt: new Date().toISOString(),
      tabs: JSON.parse(JSON.stringify(tpl.tabs)),
      activeTabId: tpl.activeTabId,
      items: JSON.parse(JSON.stringify(tpl.items))
    };
    DashboardState.selectedItemId = null;
    DashboardState.activeFilters = {};
    DashboardState.isDirty = false;
    DashboardState.updateUI();

    const titleInput = document.getElementById('dashboardTitleInput');
    if (titleInput) titleInput.value = tpl.title;

    TabBarManager.apply();
    renderAllTabs();
    renderCanvas();
    closeTemplatesModal();
    showToast(`Đã áp dụng mẫu "${tpl.title}"!`, 'success');
  }

  // ===========================================================================
  // 12.8 ELEMENT CREATION HELPERS (CANVA / FIGMA STYLE)
  // ===========================================================================
  function addChartToCanvas() {
    syncItemPositionsFromDOM();
    DashboardState.pushSnapshot();
    const sources = DataStore.getAvailableSources();
    const defaultSource = sources[0] ? sources[0].key : 'Du_Lieu_Mau_Kinh_Doanh.xlsx::BaoCaoDoanhSo';
    const newId = 'chart_' + Date.now();
    const currentCount = DashboardState.present.items.filter(i => i.tabId === DashboardState.present.activeTabId).length;
    const pos = getNextAvailablePosition(560, 360);

    DashboardState.present.items.push({
      id: newId,
      type: 'chart',
      name: `Biểu Đồ ${currentCount + 1}`,
      dataSourceKey: defaultSource,
      code: ChartTemplates.bar(defaultSource),
      tabId: DashboardState.present.activeTabId,
      pos: pos
    });

    DashboardState.selectedItemId = newId;
    DashboardState.isDirty = true;
    renderCanvas();
    triggerDebouncedAutoSave();
    showToast('Đã thêm biểu đồ vào bên dưới Canvas! Nhấn "✏️ Sửa Biểu Đồ" để thiết kế code & demo.', 'success');
    setTimeout(() => {
      document.getElementById(`item_${newId}`)?.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }, 60);
  }

  function addShapeToCanvas(shapeType = 'rectangle') {
    syncItemPositionsFromDOM();
    DashboardState.pushSnapshot();
    const newId = 'shape_' + Date.now();

    let w = 320, h = 180;
    let name = 'Khối Nền';
    let borderRadius = 8;
    if (shapeType === 'circle') {
      w = 160; h = 160;
      name = 'Hình Tròn';
      borderRadius = 50;
    } else if (shapeType === 'line') {
      w = 420; h = 24;
      name = 'Đường Kẻ';
      borderRadius = 0;
    }

    const pos = getNextAvailablePosition(w, h);

    DashboardState.present.items.push({
      id: newId,
      type: 'shape',
      shapeType: shapeType,
      name: name,
      bgColor: shapeType === 'line' ? 'transparent' : '#ffffff',
      borderColor: '#cbd5e1',
      borderWidth: 1,
      borderRadius: borderRadius,
      tabId: DashboardState.present.activeTabId,
      pos: pos
    });

    DashboardState.selectedItemId = newId;
    DashboardState.isDirty = true;
    renderCanvas();
    triggerDebouncedAutoSave();
    showToast(`Đã thêm ${name} vào bên dưới Canvas!`, 'success');
    setTimeout(() => {
      document.getElementById(`item_${newId}`)?.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }, 60);
  }

  function addImageToCanvas(dataUrl, fileName) {
    syncItemPositionsFromDOM();
    DashboardState.pushSnapshot();
    const newId = 'img_' + Date.now();
    const pos = getNextAvailablePosition(320, 220);

    DashboardState.present.items.push({
      id: newId,
      type: 'image',
      name: fileName || 'Ảnh',
      imageUrl: dataUrl,
      borderRadius: 8,
      fit: 'contain',
      tabId: DashboardState.present.activeTabId,
      pos: pos
    });

    DashboardState.selectedItemId = newId;
    DashboardState.isDirty = true;
    renderCanvas();
    triggerDebouncedAutoSave();
    showToast('Đã đưa ảnh vào bên dưới Canvas!', 'success');
    setTimeout(() => {
      document.getElementById(`item_${newId}`)?.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }, 60);
  }

  // ===========================================================================
  // 13. EVENT LISTENERS INITIALIZATION
  // ===========================================================================
  function initEventListeners() {
    // Return to hub
    document.getElementById('btnBackToHub')?.addEventListener('click', (e) => {
      e.preventDefault();
      openExitConfirmModal();
    });

    // Excel Multi-File Upload & Manager Trigger from Header Button
    const headerExcelInput = document.getElementById('excelHeaderFileInput');
    document.getElementById('btnUploadExcelFiles')?.addEventListener('click', () => {
      const btn = document.getElementById('btnUploadExcelFiles');
      if (btn?.dataset.mode === 'manage') {
        openExcelManagerModal();
      } else {
        headerExcelInput?.click();
      }
    });
    headerExcelInput?.addEventListener('change', (e) => {
      if (e.target.files && e.target.files.length > 0) {
        handleExcelFilesUpload(e.target.files);
        e.target.value = '';
      }
    });

    // Excel Modal (Popup Quản Lý File) Handlers
    document.getElementById('btnCloseExcelModal')?.addEventListener('click', closeExcelManagerModal);
    document.getElementById('btnCloseExcelModalFooter')?.addEventListener('click', closeExcelManagerModal);
    document.getElementById('excelFileInput')?.addEventListener('change', (e) => {
      if (e.target.files && e.target.files.length > 0) {
        handleExcelFilesUpload(e.target.files);
        e.target.value = '';
      }
    });

    // Drag and drop for modal drop zone
    const dropZone = document.getElementById('excelModalDropZone');
    if (dropZone) {
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
        const dtFiles = e.dataTransfer?.files;
        if (dtFiles && dtFiles.length > 0) {
          handleExcelFilesUpload(dtFiles);
        }
      });
    }

    // Modal footer: Add chart with current source
    document.getElementById('btnExcelModalAddChart')?.addEventListener('click', () => {
      closeExcelManagerModal();
      ChartStudioModal.open();
    });

    // Dropdown Add Elements (Canva / Figma Style)
    const addElementsDropdown = document.getElementById('dropdownAddElements');
    document.getElementById('btnAddElementTrigger')?.addEventListener('click', (e) => {
      e.stopPropagation();
      addElementsDropdown?.classList.toggle('open');
    });

    document.getElementById('menuAddChartCanvas')?.addEventListener('click', () => {
      addElementsDropdown?.classList.remove('open');
      addChartToCanvas();
    });

    document.getElementById('menuAddFilter')?.addEventListener('click', () => {
      addElementsDropdown?.classList.remove('open');
      WidgetConfigModal.open('filter');
    });

    document.getElementById('menuAddTopFilterBar')?.addEventListener('click', () => {
      addElementsDropdown?.classList.remove('open');
      const curTab = DashboardState.present.activeTabId || 'default';
      const cfg = TopFilterBarManager.getConfig(curTab);
      if (cfg.enabled) {
        showToast('ℹ️ Thanh Bộ Lọc Tổng Hợp đã được bật trên Tab này rồi!', 'info');
        return;
      }
      DashboardState.pushSnapshot();
      TopFilterBarManager.saveConfig({
        enabled: true,
        isSticky: true,
        isCollapsed: false,
        pos: { x: 24, y: 24, w: 1312 }
      }, curTab);
      DashboardState.isDirty = true;
      renderCanvas();
      triggerDebouncedAutoSave();
      showToast('✓ Đã thêm Thanh Bộ Lọc Tổng Hợp vào Tab này!', 'success');
    });

    document.getElementById('menuAddHeading')?.addEventListener('click', () => {
      addElementsDropdown?.classList.remove('open');
      WidgetConfigModal.open('heading');
    });

    document.getElementById('menuAddText')?.addEventListener('click', () => {
      addElementsDropdown?.classList.remove('open');
      WidgetConfigModal.open('text');
    });

    document.getElementById('menuAddShapeRect')?.addEventListener('click', () => {
      addElementsDropdown?.classList.remove('open');
      addShapeToCanvas('rectangle');
    });

    document.getElementById('menuAddShapeCircle')?.addEventListener('click', () => {
      addElementsDropdown?.classList.remove('open');
      addShapeToCanvas('circle');
    });

    document.getElementById('menuAddShapeLine')?.addEventListener('click', () => {
      addElementsDropdown?.classList.remove('open');
      addShapeToCanvas('line');
    });

    document.getElementById('menuAddFrame')?.addEventListener('click', () => {
      addElementsDropdown?.classList.remove('open');
      WidgetConfigModal.open('frame');
    });

    document.getElementById('menuAddImage')?.addEventListener('click', () => {
      addElementsDropdown?.classList.remove('open');
      document.getElementById('imageElementFileInput')?.click();
    });

    document.getElementById('imageElementFileInput')?.addEventListener('change', (e) => {
      const file = e.target.files?.[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = (ev) => {
        addImageToCanvas(ev.target.result, file.name);
      };
      reader.readAsDataURL(file);
      e.target.value = '';
    });

    document.getElementById('menuAddButton')?.addEventListener('click', () => {
      addElementsDropdown?.classList.remove('open');
      WidgetConfigModal.open('button');
    });

    // Templates Modal & Preview Triggers
    document.getElementById('btnOpenTemplatesModal')?.addEventListener('click', openTemplatesModal);
    document.getElementById('btnCloseTemplatesModal')?.addEventListener('click', closeTemplatesModal);
    document.getElementById('btnTplPreviewClose')?.addEventListener('click', closeTemplatePreview);
    document.getElementById('btnTplPreviewApply')?.addEventListener('click', (e) => {
      const key = e.target.dataset.templateKey;
      if (key) applyDashboardTemplate(key);
    });

    // Preview template buttons
    document.querySelectorAll('.db-tpl-preview-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const key = btn.dataset.previewKey;
        if (key) previewDashboardTemplate(key);
      });
    });

    // Apply template buttons
    document.querySelectorAll('.db-tpl-apply-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const key = btn.dataset.templateKey;
        if (key) applyDashboardTemplate(key);
      });
    });

    // Canvas Background Dropdown
    const canvasBgDropdown = document.getElementById('dropdownCanvasBg');
    document.getElementById('btnCanvasBgTrigger')?.addEventListener('click', (e) => {
      e.stopPropagation();
      canvasBgDropdown?.classList.toggle('open');
    });

    document.querySelectorAll('#dropdownCanvasBg [data-bg-type]').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const bgType = btn.dataset.bgType;
        if (bgType) setCanvasBackground(bgType);
        canvasBgDropdown?.classList.remove('open');
      });
    });

    // Nút Lưu Dashboard: Lần đầu chọn thư mục máy tính, từ lần 2 tự động lưu trực tiếp
    document.getElementById('btnSaveDashboardPackage')?.addEventListener('click', async () => {
      await saveAndExportDashboardPackage();
    });

    window.addEventListener('click', (e) => {
      addElementsDropdown?.classList.remove('open');
      canvasBgDropdown?.classList.remove('open');
      document.querySelectorAll('.db-filter-multiselect-dropdown.open').forEach(d => d.classList.remove('open'));
      const tplModal = document.getElementById('templatesModal');
      if (tplModal && e.target === tplModal) {
        closeTemplatesModal();
      }
      const excelModal = document.getElementById('excelFilesModal');
      if (excelModal && e.target === excelModal) {
        closeExcelManagerModal();
      }
      const importModal = document.getElementById('dashboardImportModal');
      if (importModal && e.target === importModal) {
        closeImportModal();
      }
      const saveModal = document.getElementById('saveConfirmModal');
      if (saveModal && e.target === saveModal) {
        closeSaveConfirmModal();
      }
    });

    // Theme toggle
    document.getElementById('btnToggleTheme')?.addEventListener('click', toggleTheme);

    // AI Prompt Studio triggers
    document.getElementById('btnOpenAiPromptNav')?.addEventListener('click', () => {
      AiPromptStudio.open();
    });
    document.getElementById('btnCeOpenAiPrompt')?.addEventListener('click', () => {
      const title = document.getElementById('ceChartTitleInput')?.value;
      AiPromptStudio.open(null, title);
    });
    document.getElementById('btnCloseAiPromptModal')?.addEventListener('click', () => {
      AiPromptStudio.close();
    });
    document.getElementById('aiSourceSelect')?.addEventListener('change', (e) => {
      AiPromptStudio.updateSchemaPreview(e.target.value);
      AiPromptStudio.generatePrompt();
    });
    document.getElementById('btnGenerateAiPrompt')?.addEventListener('click', () => {
      AiPromptStudio.generatePrompt();
    });
    document.getElementById('btnCopyAiPrompt')?.addEventListener('click', () => {
      AiPromptStudio.copyPrompt();
    });
    document.getElementById('btnClearAiCode')?.addEventListener('click', () => {
      document.getElementById('aiCodePasteArea').value = '';
    });
    document.getElementById('btnApplyAiCodeToEditor')?.addEventListener('click', () => {
      AiPromptStudio.applyAiCodeToEditor();
    });

    // Quick prompt chips
    document.querySelectorAll('.ai-chip').forEach(chip => {
      chip.addEventListener('click', () => {
        const req = chip.dataset.req;
        if (req) {
          document.getElementById('aiUserRequirementInput').value = req;
          AiPromptStudio.generatePrompt();
        }
      });
    });

    // Widget Modal
    document.getElementById('btnCloseWidgetModal')?.addEventListener('click', () => WidgetConfigModal.close());
    document.getElementById('btnCancelWidgetModal')?.addEventListener('click', () => WidgetConfigModal.close());
    document.getElementById('btnSaveWidgetModal')?.addEventListener('click', () => WidgetConfigModal.save());

    // Saved Dashboards Modal
    document.getElementById('btnCloseSavedDashboardsModal')?.addEventListener('click', () => SavedDashboardsCatalog.closeModal());
    document.getElementById('btnModalSaveCurrentSnapshot')?.addEventListener('click', () => {
      SavedDashboardsCatalog.saveCurrent(true);
      SavedDashboardsCatalog.renderList();
    });
    document.getElementById('btnModalImportJsonFile')?.addEventListener('click', () => {
      document.getElementById('importJsonInput')?.click();
      SavedDashboardsCatalog.closeModal();
    });
    document.getElementById('savedDashboardsList')?.addEventListener('click', async (e) => {
      const loadBtn = e.target.closest('[data-load-id]');
      if (loadBtn) {
        SavedDashboardsCatalog.loadById(loadBtn.dataset.loadId);
        return;
      }
      const exportBtn = e.target.closest('[data-export-id]');
      if (exportBtn) {
        const list = SavedDashboardsCatalog.getAll();
        const item = list.find(d => d.id === exportBtn.dataset.exportId);
        if (item) {
          const titleSlug = (item.title || 'dashboard').trim().toLowerCase().replace(/[^a-z0-9]+/g, '_');
          downloadBlob(new Blob([JSON.stringify(item, null, 2)], { type: 'application/json' }), `dashboard_${titleSlug}_config.json`);
        }
        return;
      }
      const delBtn = e.target.closest('[data-del-id]');
      if (delBtn) {
        const ok = await AppDialog.confirm('Xóa bản lưu này khỏi danh sách?', 'Xóa Bản Lưu', { isDanger: true, confirmText: 'Xóa' });
        if (ok) {
          SavedDashboardsCatalog.deleteById(delBtn.dataset.delId);
        }
        return;
      }
    });

    // Undo / Redo & Keyboard Shortcuts
    document.getElementById('btnUndo')?.addEventListener('click', () => DashboardState.undo());
    document.getElementById('btnRedo')?.addEventListener('click', () => DashboardState.redo());

    // Zoom Controls
    document.getElementById('btnZoomIn')?.addEventListener('click', () => setZoom(DashboardState.zoom + 0.1));
    document.getElementById('btnZoomOut')?.addEventListener('click', () => setZoom(DashboardState.zoom - 0.1));
    document.getElementById('btnZoomReset')?.addEventListener('click', () => setZoom(1.0));

    // Alignment Dropdown
    const alignDropdown = document.getElementById('alignDropdown');
    document.getElementById('btnAlignTrigger')?.addEventListener('click', (e) => {
      e.stopPropagation();
      alignDropdown?.classList.toggle('open');
    });
    window.addEventListener('click', () => {
      alignDropdown?.classList.remove('open');
    });

    document.getElementById('btnAlignLeft')?.addEventListener('click', () => {
      alignDropdown?.classList.remove('open');
      alignWidgets('left');
    });
    document.getElementById('btnAlignRight')?.addEventListener('click', () => {
      alignDropdown?.classList.remove('open');
      alignWidgets('right');
    });
    document.getElementById('btnAlignTop')?.addEventListener('click', () => {
      alignDropdown?.classList.remove('open');
      alignWidgets('top');
    });
    document.getElementById('btnAlignBottom')?.addEventListener('click', () => {
      alignDropdown?.classList.remove('open');
      alignWidgets('bottom');
    });
    document.getElementById('btnAlignCenter')?.addEventListener('click', () => {
      alignDropdown?.classList.remove('open');
      alignWidgets('center');
    });
    document.getElementById('btnDistributeH')?.addEventListener('click', () => {
      alignDropdown?.classList.remove('open');
      alignWidgets('distributeH');
    });

    window.addEventListener('keydown', async (e) => {
      const activeTag = document.activeElement?.tagName?.toLowerCase();
      const isInputFocused = activeTag === 'input' || activeTag === 'textarea';

      // Canva text shortcut formatting (Ctrl+B, Ctrl+I, Ctrl+U)
      if ((e.ctrlKey || e.metaKey) && DashboardState.selectedItemId && !isInputFocused) {
        const item = DashboardState.present.items.find(i => i.id === DashboardState.selectedItemId);
        if (item && (item.type === 'heading' || item.type === 'text')) {
          const keyLower = e.key.toLowerCase();
          if (keyLower === 'b') {
            e.preventDefault();
            CanvaTextManager.toggleBold();
            return;
          } else if (keyLower === 'i') {
            e.preventDefault();
            CanvaTextManager.toggleItalic();
            return;
          } else if (keyLower === 'u') {
            e.preventDefault();
            CanvaTextManager.toggleUnderline();
            return;
          }
        }
      }

      if ((e.ctrlKey || e.metaKey) && e.key === 'z') {
        e.preventDefault();
        DashboardState.undo();
      } else if ((e.ctrlKey || e.metaKey) && (e.key === 'y' || (e.shiftKey && e.key === 'Z'))) {
        e.preventDefault();
        DashboardState.redo();
      } else if ((e.ctrlKey || e.metaKey) && e.key === 's') {
        e.preventDefault();
        saveAndExportDashboardPackage();
      } else if ((e.ctrlKey || e.metaKey) && (e.key === '=' || e.key === '+')) {
        e.preventDefault();
        setZoom(DashboardState.zoom + 0.1);
      } else if ((e.ctrlKey || e.metaKey) && (e.key === '-' || e.key === '_')) {
        e.preventDefault();
        setZoom(DashboardState.zoom - 0.1);
      } else if ((e.ctrlKey || e.metaKey) && e.key === '0') {
        e.preventDefault();
        setZoom(1.0);
      } else if (e.key === 'Delete' && !isInputFocused && DashboardState.selectedItemId && DashboardState.mode !== 'preview') {
        const item = DashboardState.present.items.find(i => i.id === DashboardState.selectedItemId);
        if (item) {
          const ok = await AppDialog.confirm(`Xóa widget "${item.name || item.id}"?`, 'Xác Nhận Xóa', { isDanger: true, confirmText: 'Xóa' });
          if (ok) {
            syncItemPositionsFromDOM();
            DashboardState.pushSnapshot();
            DashboardState.present.items = DashboardState.present.items.filter(i => i.id !== DashboardState.selectedItemId);
            DashboardState.selectedItemId = null;
            DashboardState.isDirty = true;
            renderCanvas();
            triggerDebouncedAutoSave();
            showToast('Đã xóa widget', 'info');
          }
        }
      } else if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key) && !isInputFocused && DashboardState.selectedItemId && DashboardState.mode !== 'preview') {
        const item = DashboardState.present.items.find(i => i.id === DashboardState.selectedItemId);
        if (item && item.pos) {
          e.preventDefault();
          const step = e.shiftKey ? 1 : 10;
          let newX = item.pos.x;
          let newY = item.pos.y;
          if (e.key === 'ArrowUp') newY = Math.max(24, newY - step);
          if (e.key === 'ArrowDown') newY = newY + step;
          if (e.key === 'ArrowLeft') newX = Math.max(24, newX - step);
          if (e.key === 'ArrowRight') newX = Math.min(1360 - item.pos.w - 24, newX + step);
          item.pos.x = newX;
          item.pos.y = newY;
          const cardEl = document.getElementById(`item_${item.id}`);
          if (cardEl) {
            cardEl.style.left = `${newX}px`;
            cardEl.style.top = `${newY}px`;
          }
          const canvas = document.getElementById('dashboardCanvas');
          if (canvas) {
            const itemBottom = newY + item.pos.h + 80;
            const curH = parseInt(canvas.style.minHeight || '800', 10);
            if (itemBottom > curH) {
              canvas.style.minHeight = `${itemBottom + 200}px`;
            }
          }
          DashboardState.pushSnapshot();
          DashboardState.isDirty = true;
          triggerDebouncedAutoSave();
        }
      } else if (e.key === 'Escape' && DashboardState.mode === 'preview') {
        document.getElementById('btnTogglePreview')?.click();
      }
    });

    // Title inline edit
    const titleInput = document.getElementById('dashboardTitleInput');
    if (titleInput) {
      // 1. Khi đang gõ phím: lập tức cập nhật State và tiêu đề tab trình duyệt, tránh bị updateUI ghi đè
      titleInput.addEventListener('input', () => {
        const val = titleInput.value;
        DashboardState.present.title = val;
        DashboardState.isDirty = true;
        document.title = (val.trim() || 'Dashboard') + ' - Superset Dashboard & BI Studio';
        const dirtyBadge = document.getElementById('dirtyBadge');
        if (dirtyBadge) {
          dirtyBadge.textContent = '● Chưa lưu thay đổi';
          dirtyBadge.className = 'db-dirty-badge unsaved';
        }
        triggerDebouncedAutoSave();
      });

      // 2. Khi ấn Enter: tự động blur để hoàn tất chỉnh sửa
      titleInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
          titleInput.blur();
        }
      });

      // 3. Khi blur hoặc change: đẩy snapshot và đồng bộ chuẩn hóa
      const finalizeTitleChange = () => {
        const val = titleInput.value.trim() || 'Dashboard';
        titleInput.value = val;
        if (DashboardState.present.title !== val) {
          DashboardState.pushSnapshot();
          DashboardState.present.title = val;
        }
        DashboardState.isDirty = true;
        DashboardState.updateUI();
        triggerDebouncedAutoSave();
      };

      titleInput.addEventListener('change', finalizeTitleChange);
      titleInput.addEventListener('blur', finalizeTitleChange);

      document.getElementById('btnEditDashboardTitle')?.addEventListener('click', () => {
        titleInput.focus();
        titleInput.select();
      });
    }

    // Add Tab
    document.getElementById('btnAddTab')?.addEventListener('click', async () => {
      const name = await AppDialog.prompt('Nhập tên Tab mới:', `Trang ${DashboardState.present.tabs.length + 1}`, 'Thêm Tab Mới');
      if (!name) return;
      syncItemPositionsFromDOM();
      DashboardState.pushSnapshot();
      const newId = 'tab_' + Date.now();
      DashboardState.present.tabs.push({ id: newId, name: name });
      switchDashboardTab(newId);
      DashboardState.isDirty = true;
      triggerDebouncedAutoSave();
    });

    // Tab Bar Configuration triggers
    document.getElementById('btnConfigTabs')?.addEventListener('click', () => {
      if (typeof TabConfigSidebar !== 'undefined') {
        TabConfigSidebar.open();
      } else if (typeof FigmaPropertiesSidebar !== 'undefined') {
        FigmaPropertiesSidebar.openTabConfig();
      }
    });
    document.getElementById('btnCollapseTabsBar')?.addEventListener('click', () => {
      TabBarManager.toggleCollapse();
    });
    document.getElementById('btnCloseTabConfigModal')?.addEventListener('click', () => {
      TabBarManager.closeModal();
    });
    document.getElementById('btnCancelTabConfig')?.addEventListener('click', () => {
      TabBarManager.closeModal();
    });
    document.getElementById('btnApplyTabConfig')?.addEventListener('click', () => {
      TabBarManager.saveFromModal();
    });
    document.getElementById('btnResetTabConfig')?.addEventListener('click', () => {
      TabBarManager.resetDefaults();
    });

    // Palette Color selection inside modal
    document.querySelectorAll('#tcColorPalette .tc-color-dot').forEach(dot => {
      dot.addEventListener('click', () => {
        document.querySelectorAll('#tcColorPalette .tc-color-dot').forEach(d => d.classList.remove('active'));
        dot.classList.add('active');
      });
    });

    // Close tab config modal on click outside backdrop
    window.addEventListener('click', (e) => {
      const tcModal = document.getElementById('tabBarConfigModal');
      if (tcModal && e.target === tcModal) {
        TabBarManager.closeModal();
      }
    });

    // Preview Mode Toggle
    document.getElementById('btnTogglePreview')?.addEventListener('click', () => {
      const isPreview = DashboardState.mode === 'preview';
      DashboardState.mode = isPreview ? 'edit' : 'preview';
      document.body.classList.toggle('preview-mode', !isPreview);
      const btn = document.getElementById('btnTogglePreview');
      if (btn) {
        btn.classList.toggle('active-mode', !isPreview);
        btn.innerHTML = !isPreview ? '✏️ Thoát Xem' : '👁️ Xem Trước';
      }
      if (!isPreview) {
        // Vào chế độ Xem trước (dạng website):
        DashboardState.selectedItemId = null;
        document.querySelectorAll('.db-item-box').forEach(b => b.classList.remove('selected'));
        CanvaTextManager.hide();
        FigmaPropertiesSidebar.close();
        DashboardState.pan = { x: 0, y: 0 };
        DashboardState.zoom = 1.0;
        const wrapper = document.getElementById('canvasWrapper');
        if (wrapper) {
          wrapper.scrollLeft = 0;
          wrapper.scrollTop = 0;
        }
        updateCanvasTransform();
      } else {
        updateCanvasTransform();
      }
      renderCanvas();
    });

    // Floating Exit Preview Button
    document.getElementById('btnFloatingExitPreview')?.addEventListener('click', () => {
      document.getElementById('btnTogglePreview')?.click();
    });

    // Add Chart button
    document.getElementById('btnAddChart')?.addEventListener('click', () => {
      ChartStudioModal.open();
    });

    // Add Widget Dropdown
    const widgetDropdown = document.getElementById('addWidgetDropdown');
    document.getElementById('btnAddWidgetTrigger')?.addEventListener('click', (e) => {
      e.stopPropagation();
      widgetDropdown?.classList.toggle('open');
    });
    window.addEventListener('click', (e) => {
      widgetDropdown?.classList.remove('open');
      if (!e.target.closest('.db-filter-multiselect-dropdown') && !e.target.closest('.db-filter-multiselect-btn')) {
        DashboardState._openDropdownFilterId = null;
        document.querySelectorAll('.db-filter-multiselect-dropdown.open').forEach(d => {
          d.classList.remove('open');
          const pBox = d.closest('.db-item-box') || d.closest('.db-top-filter-item-wrapper');
          if (pBox) {
            pBox.classList.remove('dropdown-open');
            if (pBox.dataset.originalZIndex !== undefined) {
              pBox.style.zIndex = pBox.dataset.originalZIndex;
            } else {
              pBox.style.zIndex = '';
            }
          }
          const tb = d.closest('.db-top-filter-bar');
          if (tb) {
            tb.classList.remove('dropdown-open');
            if (tb.dataset.originalZIndex !== undefined) {
              tb.style.zIndex = tb.dataset.originalZIndex;
            } else {
              tb.style.zIndex = '';
            }
          }
        });
      }
    });

    // Lắng nghe sự kiện cuộn trang (Window, Document và Canvas Wrapper) để ghim thanh Bộ Lọc Tổng Hợp ở đỉnh trình duyệt khi xem trước
    let isTickingSticky = false;
    const handleScrollSticky = () => {
      if (!isTickingSticky) {
        window.requestAnimationFrame(() => {
          updateTopFilterBarStickyPosition();
          isTickingSticky = false;
        });
        isTickingSticky = true;
      }
    };
    window.addEventListener('scroll', handleScrollSticky, { passive: true });
    document.addEventListener('scroll', handleScrollSticky, { passive: true });
    const canvasWrapperEl = document.getElementById('canvasWrapper');
    if (canvasWrapperEl) {
      canvasWrapperEl.addEventListener('scroll', handleScrollSticky, { passive: true });
    }
    window.addEventListener('resize', () => {
      updateCanvasTransform();
      updateTopFilterBarStickyPosition();
    });

    document.getElementById('menuAddFilter')?.addEventListener('click', () => {
      widgetDropdown?.classList.remove('open');
      WidgetConfigModal.open('filter');
    });
    document.getElementById('menuAddTopFilterBar')?.addEventListener('click', () => {
      widgetDropdown?.classList.remove('open');
      const curTab = DashboardState.present.activeTabId || 'default';
      const cfg = TopFilterBarManager.getConfig(curTab);
      if (cfg.enabled) {
        showToast('ℹ️ Thanh Bộ Lọc Tổng Hợp đã được bật trên Tab này rồi!', 'info');
        return;
      }
      DashboardState.pushSnapshot();
      TopFilterBarManager.saveConfig({
        enabled: true,
        isSticky: true,
        isCollapsed: false,
        pos: { x: 24, y: 24, w: 1312 }
      }, curTab);
      DashboardState.isDirty = true;
      renderCanvas();
      triggerDebouncedAutoSave();
      showToast('✓ Đã thêm Thanh Bộ Lọc Tổng Hợp vào Tab này!', 'success');
    });
    document.getElementById('menuAddHeading')?.addEventListener('click', () => {
      widgetDropdown?.classList.remove('open');
      WidgetConfigModal.open('heading');
    });
    document.getElementById('menuAddFrame')?.addEventListener('click', () => {
      widgetDropdown?.classList.remove('open');
      WidgetConfigModal.open('frame');
    });

    // Excel Files upload
    const excelModal = document.getElementById('excelFilesModal');
    document.getElementById('btnManageExcelFiles')?.addEventListener('click', () => {
      excelModal?.classList.add('open');
      renderLoadedExcelFilesList();
    });
    document.getElementById('btnCloseExcelModal')?.addEventListener('click', () => {
      excelModal?.classList.remove('open');
    });
    document.getElementById('excelFileInput')?.addEventListener('change', (e) => {
      handleExcelFilesUpload(e.target.files);
    });

    // =========================================================================
    // IMPORT DASHBOARD (NẠP LẠI ĐỂ TIẾP TỤC CHỈNH SỬA)
    // =========================================================================
    // 1. Nút "Nạp Dashboard" trên Header
    document.getElementById('btnImportDashboard')?.addEventListener('click', () => {
      openImportModal();
    });

    // 2. Các nút bên trong modal Nạp Dashboard
    document.getElementById('btnCloseImportModal')?.addEventListener('click', () => {
      closeImportModal();
    });
    document.getElementById('btnImportFilePicker')?.addEventListener('click', () => {
      document.getElementById('importUnifiedInput')?.click();
    });
    document.getElementById('btnImportFolderPicker')?.addEventListener('click', () => {
      closeImportModal();
      importDashboardFolder();
    });

    // 3. Thẻ input tệp hợp nhất (.json, .html, .htm, .zip)
    document.getElementById('importUnifiedInput')?.addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (!file) return;
      const nameLower = file.name.toLowerCase();
      if (nameLower.endsWith('.zip')) {
        importDashboardZipFile(file);
      } else if (nameLower.endsWith('.html') || nameLower.endsWith('.htm')) {
        importDashboardHtmlFile(file);
      } else if (nameLower.endsWith('.json')) {
        importDashboardJsonFile(file);
      } else {
        showToast('Định dạng tệp không được hỗ trợ. Vui lòng chọn tệp .json, .html hoặc .zip!', 'warning');
      }
      e.target.value = '';
    });

    // 4. Thẻ input JSON cũ (tương thích ngược)
    const jsonInput = document.getElementById('importJsonInput');
    document.getElementById('btnImportJson')?.addEventListener('click', () => jsonInput?.click());
    jsonInput?.addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (!file) return;
      importDashboardJsonFile(file);
      e.target.value = '';
    });

    // 5. Các nút mở Import từ Hub / Menu
    document.getElementById('btnHeaderNewDashboard')?.addEventListener('click', () => createNewDashboard());
    document.getElementById('menuImportFolderAction')?.addEventListener('click', () => importDashboardFolder());
    document.getElementById('menuImportJsonFileAction')?.addEventListener('click', () => document.getElementById('importUnifiedInput')?.click());
    document.getElementById('btnModalImportFolder')?.addEventListener('click', () => importDashboardFolder());
    document.getElementById('btnModalImportJsonFile')?.addEventListener('click', () => document.getElementById('importUnifiedInput')?.click());

    // 6. Thẻ input chọn thư mục Fallback
    document.getElementById('folderPickerInput')?.addEventListener('change', async (e) => {
      const files = Array.from(e.target.files || []);
      if (files.length === 0) return;
      await restoreDashboardFromFolderFiles(files);
      e.target.value = '';
    });

    // Chart Studio modal events
    document.getElementById('btnCeCancel')?.addEventListener('click', () => ChartStudioModal.close());
    document.getElementById('btnCeSave')?.addEventListener('click', () => ChartStudioModal.save());
    document.getElementById('btnCeRun')?.addEventListener('click', () => ChartStudioModal.runAndPreview());
    document.getElementById('btnCeFormatCode')?.addEventListener('click', () => ChartStudioModal.formatCode());
    document.getElementById('btnCeClearTerminal')?.addEventListener('click', () => clearCeTerminal());
    document.getElementById('btnCeOpenAiPrompt')?.addEventListener('click', () => AiPromptStudio.open());

    // Chart preset template buttons
    document.querySelectorAll('.ce-tpl-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const type = btn.dataset.tpl;
        if (ChartTemplates[type]) {
          const ta = document.getElementById('ceCodeTextarea');
          if (ta) {
            ta.value = ChartTemplates[type]();
            updateCeCodeHighlight();
          }
          ChartStudioModal.runAndPreview();
        }
      });
    });

    // Canvas Widget Actions (Edit code, Edit widget, Duplicate, Delete, Toggle Header)
    document.getElementById('dashboardCanvas')?.addEventListener('click', async (e) => {
      const btn = e.target.closest('.db-item-action-btn') || e.target.closest('.db-item-mini-btn');
      if (!btn) return;
      const action = btn.dataset.action;
      const itemId = btn.dataset.id;

      if (action === 'open-sidebar') {
        const item = DashboardState.present.items.find(i => i.id === itemId);
        if (item) {
          DashboardState.selectedItemId = itemId;
          document.querySelectorAll('.db-item-box').forEach(b => b.classList.remove('selected'));
          document.getElementById(`item_${itemId}`)?.classList.add('selected');
          if (item.type === 'heading' || item.type === 'text') {
            CanvaTextManager.show(item, document.getElementById(`item_${itemId}`));
          }
          FigmaPropertiesSidebar.open(itemId);
        }
      } else if (action === 'edit') {
        ChartStudioModal.open(itemId);
      } else if (action === 'edit-widget') {
        const item = DashboardState.present.items.find(i => i.id === itemId);
        if (item) {
          DashboardState.selectedItemId = itemId;
          document.querySelectorAll('.db-item-box').forEach(b => b.classList.remove('selected'));
          document.getElementById(`item_${itemId}`)?.classList.add('selected');
          FigmaPropertiesSidebar.open(itemId);
        }
      } else if (action === 'toggle-header') {
        const item = DashboardState.present.items.find(i => i.id === itemId);
        if (item) {
          syncItemPositionsFromDOM();
          DashboardState.pushSnapshot();
          item.hideHeader = !item.hideHeader;
          DashboardState.isDirty = true;
          renderCanvas();
          triggerDebouncedAutoSave();
          showToast(item.hideHeader ? `Đã tắt Header của "${item.name || item.id}"` : `Đã bật lại Header của "${item.name || item.id}"`, 'info');
        }
      } else if (action === 'toggle-lock') {
        const item = DashboardState.present.items.find(i => i.id === itemId);
        if (item) {
          syncItemPositionsFromDOM();
          DashboardState.pushSnapshot();
          item.isLocked = !item.isLocked;
          DashboardState.isDirty = true;
          renderCanvas();
          triggerDebouncedAutoSave();
          showToast(item.isLocked ? `Đã khóa widget "${item.name}"` : `Đã mở khóa widget "${item.name}"`, 'info');
        }
      } else if (action === 'z-up') {
        const item = DashboardState.present.items.find(i => i.id === itemId);
        if (item) {
          syncItemPositionsFromDOM();
          DashboardState.pushSnapshot();
          item.pos.zIndex = (item.pos.zIndex || 10) + 2;
          DashboardState.isDirty = true;
          renderCanvas();
          triggerDebouncedAutoSave();
          showToast('Đã đưa widget lên trên', 'info');
        }
      } else if (action === 'z-down') {
        const item = DashboardState.present.items.find(i => i.id === itemId);
        if (item) {
          syncItemPositionsFromDOM();
          DashboardState.pushSnapshot();
          item.pos.zIndex = Math.max(1, (item.pos.zIndex || 10) - 2);
          DashboardState.isDirty = true;
          renderCanvas();
          triggerDebouncedAutoSave();
          showToast('Đã đưa widget xuống dưới', 'info');
        }
      } else if (action === 'pin-top') {
        const item = DashboardState.present.items.find(i => i.id === itemId);
        if (item) {
          pinFilterToTopBar(item);
          triggerDebouncedAutoSave();
        }
      } else if (action === 'duplicate') {
        syncItemPositionsFromDOM();
        const orig = DashboardState.present.items.find(i => i.id === itemId);
        if (orig) {
          DashboardState.pushSnapshot();
          const dupId = orig.type + '_' + Date.now();
          DashboardState.present.items.push({
            ...JSON.parse(JSON.stringify(orig)),
            id: dupId,
            name: `${orig.name} (Bản sao)`,
            pos: { ...orig.pos, x: orig.pos.x + 30, y: orig.pos.y + 30 }
          });
          DashboardState.isDirty = true;
          renderCanvas();
          triggerDebouncedAutoSave();
          showToast('Đã nhân bản!', 'success');
        }
      } else if (action === 'delete') {
        const ok = await AppDialog.confirm('Bạn có chắc muốn xóa phần tử này?', 'Xác Nhận Xóa', { isDanger: true, confirmText: 'Xóa' });
        if (ok) {
          syncItemPositionsFromDOM();
          DashboardState.pushSnapshot();
          DashboardState.present.items = DashboardState.present.items.filter(i => i.id !== itemId);
          if (DashboardState.selectedItemId === itemId) DashboardState.selectedItemId = null;
          DashboardState.isDirty = true;
          renderCanvas();
          triggerDebouncedAutoSave();
          showToast('Đã xóa', 'info');
        }
      }
    });

    // Confirmation modal buttons
    document.getElementById('btnConfirmSaveDirectory')?.addEventListener('click', async () => {
      const modal = document.getElementById('saveConfirmModal');
      const onAfter = modal?._onAfterSave;
      const success = await saveAndExportDashboardPackage(true);
      if (success && typeof onAfter === 'function') onAfter();
    });
    document.getElementById('btnConfirmDownloadZip')?.addEventListener('click', async () => {
      const modal = document.getElementById('saveConfirmModal');
      const onAfter = modal?._onAfterSave;
      const success = await exportDashboardZipPackage();
      if (success && typeof onAfter === 'function') onAfter();
    });
    document.getElementById('btnExportJsonDirect')?.addEventListener('click', () => {
      saveDashboardConfigJson();
    });
    document.getElementById('btnExportHtmlDirect')?.addEventListener('click', () => {
      exportStandaloneHtml();
    });
    document.getElementById('btnCloseSaveConfirmModal')?.addEventListener('click', closeSaveConfirmModal);
    document.getElementById('btnCancelSaveModal')?.addEventListener('click', closeSaveConfirmModal);
    document.getElementById('btnExitWithoutSaving')?.addEventListener('click', () => {
      closeSaveConfirmModal();
      window.location.href = 'index.html';
    });

    window.addEventListener('beforeunload', (e) => {
      if (DashboardState.isDirty) {
        e.preventDefault();
        e.returnValue = '';
      }
    });
  }

  // ===========================================================================
  // 13.5 FIGMA-STYLE FLOATING CANVAS TOOLBAR CONTROLLER
  // ===========================================================================
  function initFigmaFloatingToolbar() {
    const toolbar = document.getElementById('figmaFloatingToolbar');
    const collapseBtn = document.getElementById('fftCollapseToggle');
    const dragGrip = document.getElementById('fftDragGrip');
    if (!toolbar) return;
    if (toolbar.dataset.initialized === 'true') return;
    toolbar.dataset.initialized = 'true';

    const updateCollapseBtnUI = () => {
      const isCollapsed = toolbar.classList.contains('is-collapsed');
      if (collapseBtn) {
        collapseBtn.title = isCollapsed ? 'Mở rộng thanh công cụ' : 'Thu nhỏ thanh công cụ';
        collapseBtn.setAttribute('aria-expanded', (!isCollapsed).toString());
      }
    };

    // 1. Khôi phục trạng thái thu nhỏ / mở rộng
    try {
      const savedCollapsed = localStorage.getItem('figma_toolbar_collapsed');
      if (savedCollapsed === 'true') {
        toolbar.classList.add('is-collapsed');
      }
    } catch (_) {}
    updateCollapseBtnUI();

    // 2. Toggle Thu Nhỏ / Mở Rộng
    collapseBtn?.addEventListener('click', (e) => {
      e.stopPropagation();
      toolbar.classList.toggle('is-collapsed');
      const isCollapsed = toolbar.classList.contains('is-collapsed');
      try {
        localStorage.setItem('figma_toolbar_collapsed', isCollapsed ? 'true' : 'false');
      } catch (_) {}
      updateCollapseBtnUI();
      // Đóng dropdown nếu đang mở
      document.getElementById('dropdownAddElements')?.classList.remove('open');
      document.getElementById('dropdownCanvasBg')?.classList.remove('open');
    });

    // 3. Khôi phục vị trí đã lưu từ localStorage
    try {
      const savedPos = JSON.parse(localStorage.getItem('figma_toolbar_pos') || 'null');
      if (savedPos && typeof savedPos.left === 'number' && typeof savedPos.top === 'number') {
        const maxLeft = Math.max(10, window.innerWidth - 120);
        const maxTop = Math.max(10, window.innerHeight - 60);
        const clampedLeft = Math.min(Math.max(10, savedPos.left), maxLeft);
        const clampedTop = Math.min(Math.max(10, savedPos.top), maxTop);
        toolbar.style.left = `${clampedLeft}px`;
        toolbar.style.top = `${clampedTop}px`;
        toolbar.style.bottom = 'auto';
        toolbar.style.transform = 'none';

        if (clampedTop < window.innerHeight / 2) {
          toolbar.classList.add('opens-down');
        } else {
          toolbar.classList.remove('opens-down');
        }
      }
    } catch (_) {}

    // 4. Di chuyển tự do (Draggable)
    let isDragging = false;
    let startMouseX = 0;
    let startMouseY = 0;
    let origLeft = 0;
    let origTop = 0;

    function onMouseDown(e) {
      // Chỉ cho kéo khi click vào grip hoặc vùng trống của toolbar
      if (e.target.closest('button') || e.target.closest('.db-dropdown-menu') || e.target.closest('input')) {
        return;
      }
      e.preventDefault();
      isDragging = true;
      toolbar.classList.add('is-dragging');

      const rect = toolbar.getBoundingClientRect();
      startMouseX = e.clientX;
      startMouseY = e.clientY;
      origLeft = rect.left;
      origTop = rect.top;

      toolbar.style.bottom = 'auto';
      toolbar.style.transform = 'none';
      toolbar.style.left = `${origLeft}px`;
      toolbar.style.top = `${origTop}px`;

      window.addEventListener('mousemove', onMouseMove);
      window.addEventListener('mouseup', onMouseUp);
    }

    function onMouseMove(e) {
      if (!isDragging) return;
      const dx = e.clientX - startMouseX;
      const dy = e.clientY - startMouseY;

      let newLeft = origLeft + dx;
      let newTop = origTop + dy;

      const maxLeft = window.innerWidth - toolbar.offsetWidth - 10;
      const maxTop = window.innerHeight - toolbar.offsetHeight - 10;

      newLeft = Math.max(10, Math.min(maxLeft, newLeft));
      newTop = Math.max(10, Math.min(maxTop, newTop));

      toolbar.style.left = `${newLeft}px`;
      toolbar.style.top = `${newTop}px`;

      if (newTop < window.innerHeight / 2) {
        toolbar.classList.add('opens-down');
      } else {
        toolbar.classList.remove('opens-down');
      }
    }

    function onMouseUp() {
      if (!isDragging) return;
      isDragging = false;
      toolbar.classList.remove('is-dragging');
      window.removeEventListener('mousemove', onMouseMove);
      window.removeEventListener('mouseup', onMouseUp);

      const rect = toolbar.getBoundingClientRect();
      try {
        localStorage.setItem('figma_toolbar_pos', JSON.stringify({ left: Math.round(rect.left), top: Math.round(rect.top) }));
      } catch (_) {}
    }

    dragGrip?.addEventListener('mousedown', onMouseDown);
    toolbar.addEventListener('mousedown', onMouseDown);

    // Khi thay đổi kích thước cửa sổ
    window.addEventListener('resize', () => {
      if (DashboardState.mode === 'preview') {
        updateCanvasTransform();
      }
      const rect = toolbar.getBoundingClientRect();
      if (rect.right > window.innerWidth - 10) {
        toolbar.style.left = `${Math.max(10, window.innerWidth - toolbar.offsetWidth - 10)}px`;
      }
      if (rect.bottom > window.innerHeight - 10) {
        toolbar.style.top = `${Math.max(10, window.innerHeight - toolbar.offsetHeight - 10)}px`;
      }
    });

    // 5. Nút Zoom & Undo / Redo trên thanh Figma Toolbar
    document.getElementById('btnZoomIn')?.addEventListener('click', () => setZoom((DashboardState.zoom || 1.0) + 0.1));
    document.getElementById('btnZoomOut')?.addEventListener('click', () => setZoom((DashboardState.zoom || 1.0) - 0.1));
    document.getElementById('btnZoomReset')?.addEventListener('click', () => resetZoomAndPan());
    document.getElementById('btnFocusTopBoundary')?.addEventListener('click', () => {
      DashboardState.pan = { x: 0, y: 0 };
      setZoom(1.0);
      document.getElementById('canvasWrapper')?.scrollTo({ top: 0, left: 0, behavior: 'smooth' });
      showToast('Đã căn chỉnh về Ranh Giới Trên của Dashboard', 'info');
    });
    document.getElementById('btnUndo')?.addEventListener('click', () => DashboardState.undo());
    document.getElementById('btnRedo')?.addEventListener('click', () => DashboardState.redo());
  }

  // ===========================================================================
  // 14. INITIAL BOOTSTRAP
  // ===========================================================================
  function bootstrapDefaultWorkspace() {
    document.documentElement.setAttribute('data-theme', 'light');
    const initialBg = (typeof localStorage !== 'undefined' && localStorage.getItem('mb_dashboard_canvas_bg')) || DashboardState.present.theme?.canvasBg || 'dots';
    DashboardState.present.theme.canvasBg = initialBg;
    applyCanvasBackgroundUI(initialBg);
    DataStore.initDefault();
    updateDataSourcesSummary();

    // Default sample items if empty
    if (DashboardState.present.items.length === 0) {
      const defaultSource = 'Du_Lieu_Mau_Kinh_Doanh.xlsx::BaoCaoDoanhSo';

      // 1. Heading
      DashboardState.present.items.push({
        id: 'heading_title',
        type: 'heading',
        name: 'Tiêu Đề',
        text: 'Báo Cáo Hoạt Động & Phân Tích Kinh Doanh',
        fontSize: '22px',
        tabId: 'tab_overview',
        pos: { x: 30, y: 20, w: 580, h: 60 }
      });

      // 2. Filter Widget
      DashboardState.present.items.push({
        id: 'filter_khuvuc',
        type: 'filter',
        name: 'Khu Vực',
        filterColumn: 'Khu Vực',
        tabId: 'tab_overview',
        pos: { x: 630, y: 20, w: 260, h: 60 }
      });

      // 3. Bar Chart (ECharts Code)
      DashboardState.present.items.push({
        id: 'chart_sales_bar',
        type: 'chart',
        name: 'Doanh Thu Theo Tháng',
        dataSourceKey: defaultSource,
        code: ChartTemplates.bar(defaultSource),
        tabId: 'tab_overview',
        pos: { x: 30, y: 100, w: 640, h: 360 }
      });

      // 5. Pie/Donut Chart (ECharts Code)
      DashboardState.present.items.push({
        id: 'chart_cat_pie',
        type: 'chart',
        name: 'Tỷ Trọng Doanh Thu Theo Ngành Hàng',
        dataSourceKey: defaultSource,
        code: ChartTemplates.pie(defaultSource),
        tabId: 'tab_overview',
        pos: { x: 690, y: 100, w: 500, h: 360 }
      });
    }
  }

  function getPendingImportFromIndexedDB() {
    return new Promise((resolve) => {
      let isSettled = false;
      const settle = (val) => {
        if (!isSettled) {
          isSettled = true;
          resolve(val);
        }
      };

      try {
        const req = indexedDB.open('mb_dashboard_hub_db', 3);
        req.onupgradeneeded = (e) => {
          const db = e.target.result;
          if (!db.objectStoreNames.contains('pending_import')) {
            db.createObjectStore('pending_import', { keyPath: 'id' });
          }
          if (!db.objectStoreNames.contains('active_workspace')) {
            db.createObjectStore('active_workspace', { keyPath: 'id' });
          }
        };
        req.onsuccess = (e) => {
          try {
            const db = e.target.result;
            if (!db.objectStoreNames.contains('pending_import')) {
              settle(null);
              return;
            }
            const tx = db.transaction('pending_import', 'readwrite');
            const store = tx.objectStore('pending_import');
            const getReq = store.get('latest_import');
            getReq.onsuccess = () => {
              const res = getReq.result;
              if (res) {
                try { store.delete('latest_import'); } catch (_) {}
              }
              settle(res || null);
            };
            getReq.onerror = () => settle(null);
          } catch (err) {
            settle(null);
          }
        };
        req.onerror = () => settle(null);
        req.onblocked = () => settle(null);
      } catch (e) {
        settle(null);
      }

      // Đảm bảo không bao giờ bị nghẽn (tối đa 800ms)
      setTimeout(() => settle(null), 800);
    });
  }

  async function getPendingImport() {
    try {
      const dbResult = await getPendingImportFromIndexedDB();
      if (dbResult && dbResult.configJsonStr) {
        return dbResult;
      }
    } catch (_) {}

    try {
      const localJson = localStorage.getItem('mb_dashboard_pending_import_json');
      if (localJson) {
        localStorage.removeItem('mb_dashboard_pending_import_json');
        return {
          id: 'latest_import',
          configJsonStr: localJson,
          excelEntries: []
        };
      }
    } catch (_) {}

    return null;
  }

  function applyPendingImportData(pending) {
    if (!pending || !pending.configJsonStr) return false;
    try {
      const configJson = JSON.parse(pending.configJsonStr);
      let loadedExcelCount = 0;

      if (pending.excelEntries && pending.excelEntries.length > 0) {
        DataStore.files = {}; // Làm mới để xóa sạch dữ liệu mẫu mặc định
        pending.excelEntries.forEach(entry => {
          try {
            const isCsv = entry.name.toLowerCase().endsWith('.csv');
            let workbook;
            if (isCsv) {
              const decoder = new TextDecoder('utf-8');
              const csvText = decoder.decode(entry.buffer);
              workbook = XLSX.read(csvText, { type: 'string', raw: false });
            } else {
              workbook = XLSX.read(new Uint8Array(entry.buffer), { type: 'array' });
            }
            const isLargeFile = entry.buffer.byteLength >= PARQUET_THRESHOLD_BYTES;
            const sheetsData = {};
            (workbook.SheetNames || []).forEach(sName => {
              const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sName], { defval: '' });
              if (isLargeFile) {
                const parquetTable = ParquetEngine.encodeToParquet(rows, { fileName: entry.name, sheetName: sName });
                sheetsData[sName] = ParquetEngine.createVirtualArray(parquetTable);
              } else {
                sheetsData[sName] = rows;
              }
            });
            const sizeStr = entry.buffer.byteLength >= 1048576 
              ? (entry.buffer.byteLength / 1048576).toFixed(1) + ' MB' 
              : Math.max(1, Math.round(entry.buffer.byteLength / 1024)) + ' KB';

            DataStore.files[entry.name] = {
              name: entry.name,
              sizeStr: isLargeFile ? `${sizeStr} (📦 Parquet)` : sizeStr,
              uploadedAt: new Date().toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit' }),
              sheets: sheetsData,
              rawBytes: isCsv ? null : entry.buffer,
              isParquet: isLargeFile
            };
            loadedExcelCount++;
          } catch (err) {
            console.warn('Lỗi đọc Excel từ Hub:', entry.name, err);
          }
        });

        if (DataStore.files['Du_Lieu_Mau_Kinh_Doanh.xlsx'] && Object.keys(DataStore.files).length > 1) {
          delete DataStore.files['Du_Lieu_Mau_Kinh_Doanh.xlsx'];
        }
      }

      const ok = loadDashboardFromConfig(configJson, {
        successMsg: loadedExcelCount > 0
          ? `✓ Đã tái hiện 100% Dashboard từ thư mục ("${configJson.title || 'Dashboard'}", ${loadedExcelCount} file Excel)!`
          : `✓ Đã nạp thành công Dashboard "${configJson.title || 'Dashboard'}"!`
      });

      if (pending && pending.dirHandle) {
        activeDashboardFolderHandle = pending.dirHandle;
        activeDashboardFolderName = pending.folderName || pending.dirHandle.name || '';
        saveActiveWorkspaceHandle(activeDashboardFolderHandle, activeDashboardFolderName);
      }
      if (ok) {
        saveLastEditedSessionToLocalStorage();
      }
      return ok;
    } catch (parseErr) {
      console.error('Lỗi nạp dashboard từ pending import:', parseErr);
      return false;
    }
  }

  async function init() {
    AppDialog.init();
    CanvaTextManager.init();
    initFigmaFloatingToolbar();
    initFigmaCanvasPan();
    document.documentElement.setAttribute('data-theme', 'light');

    const urlParams = new URLSearchParams(window.location.search);
    const mode = urlParams.get('mode') || urlParams.get('action');

    let restored = false;

    // Nếu người dùng KHÔNG yêu cầu 'new' (tạo mới từ đầu):
    if (mode !== 'new') {
      // 0. Nếu có tham số ?file= hoặc ?config= trên URL
      const directFileUrl = urlParams.get('file') || urlParams.get('config');
      if (directFileUrl) {
        try {
          const res = await fetch(directFileUrl);
          if (res.ok) {
            const cfg = await res.json();
            restored = loadDashboardFromConfig(cfg, {
              successMsg: `✓ Đã nạp thành công Dashboard từ "${directFileUrl}"!`
            });
          }
        } catch (_) {}
      }

      // 1. Kiểm tra Pending Import từ Hub (với IndexedDB v3 và localStorage fallback)
      if (!restored) {
        try {
          const pending = await getPendingImport();
          if (pending && pending.configJsonStr) {
            restored = applyPendingImportData(pending);
          }
        } catch (importErr) {
          console.warn('Lỗi kiểm tra pending import:', importErr);
        }
      }

      // 2. Nếu không có pending import từ Hub, thử khôi phục từ active directory handle nếu có
      if (!restored) {
        try {
          const savedWs = await getActiveWorkspaceHandle();
          if (savedWs && savedWs.dirHandle) {
            activeDashboardFolderHandle = savedWs.dirHandle;
            activeDashboardFolderName = savedWs.folderName || savedWs.dirHandle.name || '';
            try {
              let perm = 'prompt';
              if (activeDashboardFolderHandle.queryPermission) {
                perm = await activeDashboardFolderHandle.queryPermission({ mode: 'read' });
              }
              if (perm === 'granted') {
                const filesArray = [];
                for await (const entry of activeDashboardFolderHandle.values()) {
                  if (entry.kind === 'file') {
                    filesArray.push(await entry.getFile());
                  }
                }
                if (filesArray.length > 0) {
                  restored = await restoreDashboardFromFolderFiles(filesArray);
                }
              }
            } catch (_) {}
          }
        } catch (_) {}
      }

      // 3. Khôi phục từ phiên làm việc gần nhất (Last Edited Session) trong localStorage
      if (!restored) {
        try {
          const lastSessionStr = localStorage.getItem('mb_dashboard_last_edited_session');
          if (lastSessionStr) {
            const lastConfig = JSON.parse(lastSessionStr);
            if (lastConfig && (Array.isArray(lastConfig.items) || Array.isArray(lastConfig.tabs) || lastConfig.title)) {
              try {
                const cachedFilesStr = localStorage.getItem('mb_dashboard_last_files_data');
                if (cachedFilesStr) {
                  const cachedFiles = JSON.parse(cachedFilesStr);
                  if (cachedFiles && Object.keys(cachedFiles).length > 0) {
                    DataStore.files = cachedFiles;
                  }
                }
              } catch (_) {}

              restored = loadDashboardFromConfig(lastConfig, {
                successMsg: `✓ Đã tự động tái hiện Dashboard vừa chỉnh sửa gần nhất ("${lastConfig.title || 'Dashboard'}")!`
              });
            }
          }
        } catch (sessErr) {
          console.warn('Khôi phục session gần nhất thất bại:', sessErr);
        }
      }

      // 4. Khôi phục từ SavedDashboardsCatalog nếu có bản lưu
      if (!restored) {
        try {
          const savedList = SavedDashboardsCatalog.getAll();
          if (Array.isArray(savedList) && savedList.length > 0 && savedList[0]) {
            restored = loadDashboardFromConfig(savedList[0], {
              successMsg: `✓ Đã tái hiện Dashboard đã lưu gần nhất: "${savedList[0].title || 'Dashboard'}"`
            });
          }
        } catch (_) {}
      }
    }

    // Nếu không có bất kỳ cấu hình nào hoặc người dùng chọn mode === 'new'
    if (!restored) {
      bootstrapDefaultWorkspace();
      renderAllTabs();
      renderCanvas();
      DashboardState.updateUI();
      updateExcelButtonState();
    }

    // Áp dụng TabBar, Properties Sidebar và gắn kết các sự kiện
    TabBarManager.apply();
    FigmaPropertiesSidebar.init();
    TabConfigSidebar.init();
    initEventListeners();

    if (mode === 'new') {
      createNewDashboard(true);
    } else if (mode === 'import_folder') {
      setTimeout(() => {
        importDashboardFolder();
      }, 400);
    }
  }

  window.__DashboardBuilder = {
    DataStore,
    DashboardState,
    getFilteredData,
    renderCanvas,
    generateStandaloneDashboardHTML,
    exportStandaloneHtml,
    WidgetConfigModal,
    FigmaPropertiesSidebar
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

})();
