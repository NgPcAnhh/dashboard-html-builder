/**
 * =============================================================================
 * THEME MANAGER ENGINE - LIGHT & DARK MODE
 * Phối màu lấy cảm hứng từ 2 trang phục Polo MB Bank:
 * - DARK MODE: Áo Polo MB Xanh Navy đậm (#090e1a / #0f172a / #162344)
 *   phối sọc tay áo Royal Blue (#2563eb), Cyan (#38bdf8) & Ngôi sao đỏ (#ef4444)
 * - LIGHT MODE: Áo Polo MB Trắng tinh khôi (#ffffff / #f4f7fb)
 *   phối cổ áo MB Royal Blue (#004ea8), viền Cyan (#0284c7) & Ngôi sao đỏ (#ef4444)
 * =============================================================================
 */

(function () {
  'use strict';

  const STORAGE_KEY = 'app_theme_mode'; // 'light' | 'dark' | 'system'

  /**
   * Lấy cấu hình theme đã lưu từ JSON Storage hoặc mặc định là 'dark'
   */
  function getSavedTheme() {
    if (window.JsonStorage) {
      return window.JsonStorage.get('theme', 'dark');
    }
    if (window.APP_JSON_STORAGE && window.APP_JSON_STORAGE.settings && window.APP_JSON_STORAGE.settings.theme) {
      return window.APP_JSON_STORAGE.settings.theme;
    }
    return 'dark';
  }

  /**
   * Xác định xem hệ điều hành đang dùng Dark hay Light mode
   */
  function getSystemPreference() {
    if (window.matchMedia && window.matchMedia('(prefers-color-scheme: light)').matches) {
      return 'light';
    }
    return 'dark';
  }

  /**
   * Phân giải theme thực tế ('light' hoặc 'dark')
   */
  function resolveEffectiveTheme(theme) {
    if (theme === 'system') {
      return getSystemPreference();
    }
    return theme === 'light' ? 'light' : 'dark';
  }

  /**
   * Cập nhật toàn bộ các nút Toggle và nhãn trên giao diện
   */
  function updateThemeUI(theme, effectiveTheme) {
    // 1. Cập nhật nhãn / icon cho các nút Toggle nhanh
    document.querySelectorAll('.btn-theme-toggle, #vizThemeToggle').forEach(btn => {
      const iconSpan = btn.querySelector('.theme-icon');
      const textSpan = btn.querySelector('.theme-text');
      
      if (effectiveTheme === 'dark') {
        if (iconSpan) iconSpan.textContent = '🌙';
        if (textSpan) textSpan.textContent = 'Dark Mode';
        btn.setAttribute('title', 'Đang dùng Chế độ Tối (Dark Mode). Nhấp để đổi sang Chế độ Sáng');
        btn.classList.add('theme-dark-active');
        btn.classList.remove('theme-light-active');
      } else {
        if (iconSpan) iconSpan.textContent = '☀️';
        if (textSpan) textSpan.textContent = 'Light Mode';
        btn.setAttribute('title', 'Đang dùng Chế độ Sáng (Light Mode). Nhấp để đổi sang Chế độ Tối');
        btn.classList.add('theme-light-active');
        btn.classList.remove('theme-dark-active');
      }
    });

    // 2. Cập nhật các thẻ chọn trong màn hình Setting (nếu có trên trang)
    document.querySelectorAll('.theme-card-option').forEach(card => {
      const mode = card.getAttribute('data-mode');
      if (mode === theme) {
        card.classList.add('active');
        const radio = card.querySelector('input[type="radio"]');
        if (radio) radio.checked = true;
      } else {
        card.classList.remove('active');
        const radio = card.querySelector('input[type="radio"]');
        if (radio) radio.checked = false;
      }
    });

    // 3. Cập nhật Pill trạng thái trong Setting
    const pillLabel = document.getElementById('themeCurrentLabel');
    if (pillLabel) {
      if (theme === 'system') {
        pillLabel.textContent = `Tự động (${effectiveTheme === 'dark' ? 'Đang là Chế độ Tối' : 'Đang là Chế độ Sáng'})`;
      } else if (theme === 'light') {
        pillLabel.textContent = 'Đang dùng: Chế độ Sáng (Light Mode)';
      } else {
        pillLabel.textContent = 'Đang dùng: Chế độ Tối (Dark Mode)';
      }
    }
  }

  /**
   * Áp dụng theme lên thẻ <html> và lưu vào Storage
   */
  function applyTheme(theme, save = true) {
    const effectiveTheme = resolveEffectiveTheme(theme);
    
    // Gán thuộc tính lên html
    document.documentElement.setAttribute('data-theme', effectiveTheme);
    document.documentElement.setAttribute('data-theme-setting', theme);

    if (save) {
      if (window.JsonStorage) {
        window.JsonStorage.set('theme', theme);
      } else if (window.APP_JSON_STORAGE && window.APP_JSON_STORAGE.settings) {
        window.APP_JSON_STORAGE.settings.theme = theme;
      }
    }

    // Cập nhật các phần tử UI
    updateThemeUI(theme, effectiveTheme);

    // Bắn Custom Event để các script khác nếu cần có thể lắng nghe
    window.dispatchEvent(new CustomEvent('app-theme-changed', {
      detail: { theme, effectiveTheme }
    }));
  }

  /**
   * Đổi nhanh qua lại giữa Dark và Light
   */
  function toggleTheme() {
    const current = getSavedTheme();
    const effective = resolveEffectiveTheme(current);
    const nextTheme = effective === 'dark' ? 'light' : 'dark';
    applyTheme(nextTheme, true);
    return nextTheme;
  }

  // 1. Áp dụng ngay lập tức trước khi render để chống chớp sáng/tối (FOUC)
  const initialTheme = getSavedTheme();
  const initialEffective = resolveEffectiveTheme(initialTheme);
  document.documentElement.setAttribute('data-theme', initialEffective);
  document.documentElement.setAttribute('data-theme-setting', initialTheme);

  // 2. Lắng nghe thay đổi chế độ của hệ điều hành nếu đang ở chế độ 'system'
  if (window.matchMedia) {
    window.matchMedia('(prefers-color-scheme: light)').addEventListener('change', () => {
      if (getSavedTheme() === 'system') {
        applyTheme('system', false);
      }
    });
  }

  // 3. Đồng bộ real-time giữa các tab trình duyệt
  window.addEventListener('storage', (e) => {
    if (e.key === STORAGE_KEY && e.newValue) {
      applyTheme(e.newValue, false);
    }
  });

  /**
   * Hiển thị thông báo Toast linh hoạt trên mọi trang
   */
  function notifyToast(message) {
    if (typeof window.showToast === 'function') {
      window.showToast(message);
      return;
    }
    const copyToast = document.getElementById('copyToast');
    const copyToastMsg = document.getElementById('copyToastMsg');
    if (copyToast && copyToastMsg) {
      copyToastMsg.textContent = message;
      copyToast.classList.add('show');
      setTimeout(() => copyToast.classList.remove('show'), 2400);
      return;
    }
    let toast = document.getElementById('themeToastNotification');
    if (!toast) {
      toast = document.createElement('div');
      toast.id = 'themeToastNotification';
      toast.style.cssText = 'position:fixed;bottom:24px;right:24px;padding:10px 18px;border-radius:10px;background:var(--primary);color:#fff;font-size:12.5px;font-weight:700;box-shadow:0 8px 24px rgba(0,0,0,0.35);z-index:99999;transition:all 0.3s ease;opacity:0;transform:translateY(20px);pointer-events:none;display:flex;align-items:center;gap:8px;';
      document.body.appendChild(toast);
    }
    toast.textContent = message;
    toast.style.opacity = '1';
    toast.style.transform = 'translateY(0)';
    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(20px)';
    }, 2400);
  }

  /**
   * Gắn sự kiện click vào các thành phần UI sau khi DOM sẵn sàng
   */
  function bindUI() {
    // Nút toggle nhanh
    document.querySelectorAll('.btn-theme-toggle, #vizThemeToggle').forEach(btn => {
      if (btn.dataset.themeBound) return;
      btn.dataset.themeBound = 'true';

      btn.addEventListener('click', (e) => {
        e.preventDefault();
        const next = toggleTheme();
        notifyToast(next === 'light' ? '☀️ Đã chuyển sang Giao diện Sáng (Light Mode)' : '🌙 Đã chuyển sang Giao diện Tối (Dark Mode)');
      });
    });

    // Thẻ chọn trong Settings View
    document.querySelectorAll('.theme-card-option').forEach(card => {
      if (card.dataset.themeBound) return;
      card.dataset.themeBound = 'true';

      card.addEventListener('click', () => {
        const mode = card.getAttribute('data-mode');
        if (mode) {
          applyTheme(mode, true);
          const labels = {
            dark: '🌙 Đã kích hoạt Chế độ Tối (Dark Mode)',
            light: '☀️ Đã kích hoạt Chế độ Sáng (Light Mode)',
            system: '💻 Đã kích hoạt Chế độ Tự Động theo Thiết Bị'
          };
          notifyToast(labels[mode] || 'Đã cập nhật giao diện');
        }
      });
    });

    // Cập nhật lại UI sau khi gắn kết
    const currentTheme = getSavedTheme();
    updateThemeUI(currentTheme, resolveEffectiveTheme(currentTheme));
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', bindUI);
  } else {
    bindUI();
  }

  window.addEventListener('json-storage-updated', () => {
    const updatedTheme = getSavedTheme();
    applyTheme(updatedTheme, false);
  });

  // Xuất API toàn cục
  window.AppTheme = {
    getTheme: getSavedTheme,
    getEffectiveTheme: () => resolveEffectiveTheme(getSavedTheme()),
    setTheme: (t) => applyTheme(t, true),
    toggleTheme: toggleTheme,
    bindUI: bindUI
  };

})();
