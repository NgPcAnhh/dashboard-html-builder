/**
 * json-storage.js - Quản lý Lưu Trữ Dữ Liệu Qua File JSON
 * Thay thế hoàn toàn LocalStorage:
 * - Lưu trữ dữ liệu cấu hình, theme, pinned tools và custom tools vào file JSON trong thư mục data/
 * - Hỗ trợ File System Access API để lưu trực tiếp vào data/app_storage.json
 * - Cung cấp tính năng Nạp (Import) và Xuất (Export) file JSON trực tiếp
 * - Tương thích 100% với môi trường offline file:/// thông qua data/storage.js
 */

(function () {
  'use strict';

  // Dữ liệu mặc định ban đầu
  const DEFAULT_DATA = {
    version: "1.0",
    lastUpdated: new Date().toISOString(),
    settings: {
      theme: "dark",
      sidebarCollapsed: false
    },
    pinnedToolIds: [
      "excel-csv-terminal",
      "excel-schema-matcher",
      "excel-column-normalizer"
    ],
    customTools: []
  };

  // Khởi tạo từ window.APP_JSON_STORAGE nếu có sẵn từ data/storage.js
  let _data = (window.APP_JSON_STORAGE && typeof window.APP_JSON_STORAGE === 'object')
    ? JSON.parse(JSON.stringify(window.APP_JSON_STORAGE))
    : JSON.parse(JSON.stringify(DEFAULT_DATA));

  // Đảm bảo các thuộc tính bắt buộc luôn tồn tại
  if (!_data.settings) _data.settings = { ...DEFAULT_DATA.settings };
  if (!Array.isArray(_data.pinnedToolIds)) _data.pinnedToolIds = [...DEFAULT_DATA.pinnedToolIds];
  if (!Array.isArray(_data.customTools)) _data.customTools = [];

  let _dataDirHandle = null;

  const JsonStorage = {
    /**
     * Lấy giá trị theo key
     */
    get: function (key, defaultValue) {
      if (key === 'theme') {
        return (_data.settings && _data.settings.theme) || defaultValue || 'dark';
      }
      if (key === 'sidebarCollapsed') {
        return (_data.settings && typeof _data.settings.sidebarCollapsed === 'boolean')
          ? _data.settings.sidebarCollapsed
          : (defaultValue !== undefined ? defaultValue : false);
      }
      if (key in _data) {
        return _data[key];
      }
      return defaultValue !== undefined ? defaultValue : null;
    },

    /**
     * Ghi giá trị theo key và tự động đồng bộ
     */
    set: function (key, value, autoSave = true) {
      if (key === 'theme') {
        if (!_data.settings) _data.settings = {};
        _data.settings.theme = value;
      } else if (key === 'sidebarCollapsed') {
        if (!_data.settings) _data.settings = {};
        _data.settings.sidebarCollapsed = !!value;
      } else {
        _data[key] = value;
      }

      _data.lastUpdated = new Date().toISOString();

      if (autoSave && _dataDirHandle) {
        this.saveToLinkedDirectory().catch(e => console.warn('Lỗi tự động lưu vào folder data:', e));
      }

      window.dispatchEvent(new CustomEvent('json-storage-changed', {
        detail: { key, value, data: _data }
      }));
    },

    /**
     * Lấy toàn bộ dữ liệu JSON hiện tại
     */
    getAllData: function () {
      return JSON.parse(JSON.stringify(_data));
    },

    /**
     * Cập nhật toàn bộ dữ liệu từ đối tượng JSON mới
     */
    setAllData: function (newData) {
      if (!newData || typeof newData !== 'object') return false;
      _data = {
        version: newData.version || "1.0",
        lastUpdated: new Date().toISOString(),
        settings: {
          theme: (newData.settings && newData.settings.theme) || _data.settings.theme || "dark",
          sidebarCollapsed: (newData.settings && typeof newData.settings.sidebarCollapsed === 'boolean')
            ? newData.settings.sidebarCollapsed
            : _data.settings.sidebarCollapsed
        },
        pinnedToolIds: Array.isArray(newData.pinnedToolIds) ? newData.pinnedToolIds : _data.pinnedToolIds,
        customTools: Array.isArray(newData.customTools) ? newData.customTools : _data.customTools
      };

      window.dispatchEvent(new CustomEvent('json-storage-updated', {
        detail: { data: _data }
      }));

      return true;
    },

    /**
     * Liên kết thư mục data/ qua File System Access API
     */
    linkDataDirectory: async function () {
      if (!('showDirectoryPicker' in window)) {
        alert("Trình duyệt này không hỗ trợ File System Access API. Bạn vẫn có thể dùng tính năng Xuất/Nạp file JSON!");
        return false;
      }

      try {
        const handle = await window.showDirectoryPicker();
        let targetHandle = handle;

        // Nếu người dùng chọn thư mục cha của 'data', tự động tìm hoặc mở folder 'data'
        if (handle.name !== 'data') {
          try {
            targetHandle = await handle.getDirectoryHandle('data', { create: true });
          } catch (err) {
            targetHandle = handle;
          }
        }

        _dataDirHandle = targetHandle;

        // Thử đọc file app_storage.json có sẵn trong thư mục
        try {
          const fileHandle = await _dataDirHandle.getFileHandle('app_storage.json');
          const file = await fileHandle.getFile();
          const text = await file.text();
          const parsed = JSON.parse(text);
          this.setAllData(parsed);
          console.log(`[JSON Storage] Đã nạp thành công dữ liệu từ folder: ${targetHandle.name}/app_storage.json`);
        } catch (e) {
          // Nếu file chưa có thì ghi file hiện tại vào thư mục
          await this.saveToLinkedDirectory();
        }

        window.dispatchEvent(new CustomEvent('json-storage-folder-linked', {
          detail: { folderName: targetHandle.name }
        }));

        return true;
      } catch (err) {
        if (err.name !== 'AbortError') {
          console.error('[JSON Storage] Lỗi liên kết thư mục:', err);
        }
        return false;
      }
    },

    /**
     * Ghi dữ liệu vào thư mục data/ đã liên kết
     */
    saveToLinkedDirectory: async function () {
      if (!_dataDirHandle) return false;
      try {
        const jsonContent = JSON.stringify(_data, null, 2);
        
        // Ghi file app_storage.json
        const fileHandle = await _dataDirHandle.getFileHandle('app_storage.json', { create: true });
        const writable = await fileHandle.createWritable();
        await writable.write(jsonContent);
        await writable.close();

        // Đồng thời ghi file storage.js để hỗ trợ nạp offline file:///
        try {
          const jsContent = `/**\n * data/storage.js - Bản sao đồng bộ dữ liệu JSON\n * Cập nhật: ${new Date().toLocaleString('vi-VN')}\n */\nwindow.APP_JSON_STORAGE = ${jsonContent};\n`;
          const jsHandle = await _dataDirHandle.getFileHandle('storage.js', { create: true });
          const jsWritable = await jsHandle.createWritable();
          await jsWritable.write(jsContent);
          await jsWritable.close();
        } catch (eJs) {
          // Bỏ qua nếu không ghi được storage.js
        }

        console.log(`[JSON Storage] Đã lưu thành công dữ liệu vào folder: ${_dataDirHandle.name}/app_storage.json`);
        return true;
      } catch (err) {
        console.error('[JSON Storage] Lỗi khi ghi file vào thư mục:', err);
        return false;
      }
    },

    /**
     * Xuất và tải file app_storage.json về máy tính
     */
    exportJsonFile: function (fileName = 'app_storage.json') {
      const jsonContent = JSON.stringify(_data, null, 2);
      const blob = new Blob([jsonContent], { type: 'application/json;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      console.log(`[JSON Storage] Đã tải xuống file: ${fileName}`);
    },

    /**
     * Nạp dữ liệu từ một file JSON do người dùng chọn
     */
    importJsonFile: function (file) {
      return new Promise((resolve, reject) => {
        if (!file) return reject(new Error("Chưa chọn file"));
        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            const parsed = JSON.parse(e.target.result);
            const success = this.setAllData(parsed);
            if (success) {
              if (_dataDirHandle) {
                this.saveToLinkedDirectory().catch(err => console.warn(err));
              }
              resolve(_data);
            } else {
              reject(new Error("Cấu trúc file JSON không hợp lệ"));
            }
          } catch (err) {
            reject(err);
          }
        };
        reader.onerror = () => reject(new Error("Lỗi khi đọc file"));
        reader.readAsText(file);
      });
    },

    /**
     * Kiểm tra xem đã liên kết thư mục chưa
     */
    isFolderLinked: function () {
      return !!_dataDirHandle;
    }
  };

  window.JsonStorage = JsonStorage;
})();
