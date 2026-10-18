/**
 * TOOL HUB & WORKSPACE MANAGER - JAVASCRIPT CONTROLLER
 * Handles:
 * - Registered Tools Catalog
 * - Collapsible Sidebar with LocalStorage persistence
 * - Real-time Search, Category Filter, and Sorting
 * - Pagination (configurable items per page, dynamic page numbers)
 * - Sidebar Tool Pinning & Order Rearrangement (Drag-and-Drop + Up/Down) with LocalStorage sync
 * - View switching between Tools Management and Settings
 */

(function () {
  'use strict';

  // =========================================================================
  // 1. REGISTRY OF TOOLS (CATALOG)
  // =========================================================================
  // =========================================================================
  // 1. REGISTRY OF TOOLS (CATALOG) - DUY NHẤT 2 CÔNG CỤ CỐT LÕI
  // =========================================================================
  const REGISTERED_TOOLS = [
    {
      id: 'excel-csv-terminal',
      name: 'Xử lý dữ liệu bằng Javascript',
      category: 'data',
      categoryLabel: 'Xử lý Dữ liệu',
      icon: '📊',
      desc: 'Môi trường lập trình JS xử lý file Excel, CSV dung lượng lớn, tạo mock data, chạy script phân tích và biến đổi dữ liệu tự động với SheetJS.',
      tags: ['Excel', 'CSV', 'SheetJS', 'Scripting', 'Clean Data'],
      badge: 'Công Cụ 1 • Sẵn Sàng',
      isPrimary: true,
      url: 'data_processing.html',
      author: 'Data Engine'
    },
    {
      id: 'excel-schema-matcher',
      name: 'Kiểm Tra Đối Soát 2 File Excel',
      category: 'data',
      categoryLabel: 'Xử lý Dữ liệu',
      icon: '🔍',
      desc: 'Tự động đối soát, mapping và chẩn đoán lỗi format giữa 2 file Excel (File Mẫu vs File Gửi Mới). Xử lý lệch dòng tiêu đề, đảo vị trí cột, sai tên trường và xuất file chuẩn hóa.',
      tags: ['Excel', 'Đối Soát', 'Schema Matcher', 'Fuzzy Mapping', 'Auto Clean'],
      badge: 'Công Cụ 2 • Sẵn Sàng',
      isPrimary: true,
      url: 'excel_matcher.html',
      author: 'Data Quality Engine'
    },
    {
      id: 'excel-column-normalizer',
      name: 'Chuẩn Hóa & Sắp Xếp Cột Theo Kỳ Gốc',
      category: 'data',
      categoryLabel: 'Xử lý Dữ liệu',
      icon: '🔄',
      desc: 'Tự động phát hiện, ánh xạ (mapping) và sắp xếp lại toàn bộ thứ tự cột, chuẩn hóa tên trường của file kỳ mới theo đúng 100% format chuẩn của file kỳ gốc.',
      tags: ['Excel', 'Chuẩn Hóa', 'Sắp Xếp Cột', 'Auto Mapping', 'Đổi Tên Trường'],
      badge: 'Công Cụ 3 • Mới',
      isPrimary: true,
      url: 'excel_normalizer.html',
      author: 'Data Normalizer Engine'
    },
    {
      id: 'echarts-data-visualizer',
      name: 'Trực Quan Hóa Dữ Liệu (ECharts Studio)',
      category: 'data',
      categoryLabel: 'Xử lý Dữ liệu',
      icon: '📈',
      desc: 'Môi trường live compiler biên dịch code JavaScript dùng thư viện Apache ECharts để demo và dựng biểu đồ tương tác thời gian thực từ Excel/CSV, mock data hoặc JSON tùy biến.',
      tags: ['ECharts', 'Data Visualization', 'Biểu Đồ', 'JavaScript', 'Live Compiler'],
      badge: 'Công Cụ 4 • Mới',
      isPrimary: true,
      url: 'echarts_visualizer.html',
      author: 'Visualization Engine'
    },
    {
      id: 'superset-dashboard-builder',
      name: 'BI Studio',
      category: 'data',
      categoryLabel: 'Xử lý Dữ liệu',
      icon: '📊',
      desc: 'Xây dựng Dashboard BI: Xử lý dữ liệu Excel bằng JavaScript & ECharts không qua database, trợ lý System Prompt AI cho người dùng Low-Code, lưu cấu hình JSON để chỉnh sửa lại và xuất file HTML độc lập trỏ vào file Excel.',
      tags: ['Dashboard', 'ECharts', 'BI', 'Excel', 'Lưới 12 Cột', 'Bộ Lọc', 'AI Prompt', 'JSON Config'],
      badge: 'Công Cụ 5 • BI Studio',
      isPrimary: true,
      url: 'dashboard_builder.html',
      author: 'BI Analytics Engine'
    },
    {
      id: 'excel-data-masker',
      name: 'Che Mờ & Bảo Mật Dữ Liệu (Excel & CSV)',
      category: 'data',
      categoryLabel: 'Xử lý Dữ liệu',
      icon: '🎭',
      desc: 'Tự động che mờ (masking), ẩn thông tin nhạy cảm trên tất cả các sheet Excel hoặc CSV. Tùy chọn giữ nguyên dòng đầu làm Header, che mờ < 5 ký tự bằng *****, ô > 5 ký tự giữ 3 ký tự đầu và che mờ phần sau (ví dụ: abc****).',
      tags: ['Excel', 'CSV', 'Che Mờ Dữ Liệu', 'Data Masking', 'Bảo Mật', 'Header Option'],
      badge: 'Công Cụ 6 • Mới',
      isPrimary: true,
      url: 'data_masker.html',
      author: 'Data Security Engine'
    }
  ];

  // =========================================================================
  // 2. DEFAULT PINNED TOOLS & STORAGE KEYS
  // =========================================================================
  const STORAGE_KEY_PINNED = 'hub_sidebar_pinned_tools';
  const DEFAULT_PINNED_IDS = ['excel-csv-terminal', 'excel-schema-matcher', 'excel-column-normalizer'];

  // =========================================================================
  // 2.1. CUSTOM TOOLS STORAGE HANDLERS (LƯU TRỮ VÀO FILE JSON TRONG THƯ MỤC DATA)
  // =========================================================================
  function loadCustomTools() {
    if (window.JsonStorage) {
      const list = window.JsonStorage.get('customTools', []);
      if (Array.isArray(list)) return list;
    }
    if (window.APP_JSON_STORAGE && Array.isArray(window.APP_JSON_STORAGE.customTools)) {
      return window.APP_JSON_STORAGE.customTools;
    }
    return [];
  }

  function saveCustomTools(toolsList) {
    if (window.JsonStorage) {
      window.JsonStorage.set('customTools', toolsList);
    } else if (window.APP_JSON_STORAGE) {
      window.APP_JSON_STORAGE.customTools = toolsList;
    }
  }

  function getAllTools() {
    const customList = loadCustomTools();
    return [...REGISTERED_TOOLS, ...customList];
  }

  // =========================================================================
  // 3. APPLICATION STATE
  // =========================================================================
  const state = {
    pinnedToolIds: loadPinnedTools(),
    sidebarCollapsed: window.JsonStorage ? window.JsonStorage.get('sidebarCollapsed', false) : false,
    currentView: 'tools', // 'tools' | 'settings'
    searchQuery: '',
    selectedCategory: 'all',
    sortMode: 'az', // 'az' | 'za' | 'popular'
    currentPage: 1,
    pageSize: 6, // Hiển thị 6 tools / trang để chứa cả các tool tùy chỉnh
    availableSearchQuery: ''
  };

  function loadPinnedTools() {
    const allTools = getAllTools();
    let validIds = [];
    if (window.JsonStorage) {
      const saved = window.JsonStorage.get('pinnedToolIds', null);
      if (Array.isArray(saved) && saved.length > 0) {
        validIds = saved.filter(id => allTools.some(t => t.id === id));
      }
    }
    if (validIds.length === 0 && window.APP_JSON_STORAGE && Array.isArray(window.APP_JSON_STORAGE.pinnedToolIds)) {
      validIds = window.APP_JSON_STORAGE.pinnedToolIds.filter(id => allTools.some(t => t.id === id));
    }
    if (validIds.length > 0) return validIds;

    const defaults = [...DEFAULT_PINNED_IDS];
    if (window.JsonStorage) {
      window.JsonStorage.set('pinnedToolIds', defaults);
    }
    return defaults;
  }

  function savePinnedTools() {
    if (window.JsonStorage) {
      window.JsonStorage.set('pinnedToolIds', state.pinnedToolIds);
      showToast('Đã lưu cấu hình Sidebar vào File JSON');
    }
  }

  // =========================================================================
  // 4. DOM ELEMENTS CACHE
  // =========================================================================
  const elements = {
    sidebar: document.getElementById('sidebar'),
    sidebarToggleBtn: document.getElementById('sidebarToggleBtn'),
    sidebarPinnedList: document.getElementById('sidebarPinnedList'),
    sidebarPinnedCount: document.getElementById('sidebarPinnedCount'),
    navTools: document.getElementById('navTools'),
    navSettings: document.getElementById('navSettings'),

    viewToolsManager: document.getElementById('viewToolsManager'),
    viewSettings: document.getElementById('viewSettings'),

    // Tools Manager View elements
    searchInput: document.getElementById('searchInput'),
    searchClearBtn: document.getElementById('searchClearBtn'),
    categoryPills: document.getElementById('categoryPills'),
    sortSelect: document.getElementById('sortSelect'),
    toolsGrid: document.getElementById('toolsGrid'),
    statTotalTools: document.getElementById('statTotalTools'),
    statPinnedTools: document.getElementById('statPinnedTools'),

    // Pagination elements
    paginationWrapper: document.getElementById('paginationWrapper'),
    paginationInfo: document.getElementById('paginationInfo'),
    paginationControls: document.getElementById('paginationControls'),
    pageSizeSelect: document.getElementById('pageSizeSelect'),

    // Settings View elements
    pinnedReorderList: document.getElementById('pinnedReorderList'),
    pinnedCountBadge: document.getElementById('pinnedCountBadge'),
    availableToolsList: document.getElementById('availableToolsList'),
    availableSearchInput: document.getElementById('availableSearchInput'),
    btnResetDefault: document.getElementById('btnResetDefault'),

    // Toast & Modal
    toastNotification: document.getElementById('toastNotification'),
    toastMessage: document.getElementById('toastMessage'),
    modalOverlay: document.getElementById('modalOverlay'),
    modalTitle: document.getElementById('modalTitle'),
    modalDesc: document.getElementById('modalDesc'),
    modalTags: document.getElementById('modalTags'),
    modalActionBtn: document.getElementById('modalActionBtn'),
    modalCloseBtn: document.getElementById('modalCloseBtn')
  };

  // =========================================================================
  // 5. SIDEBAR COLLAPSE / EXPAND LOGIC (OPTIONAL / IF PRESENT)
  // =========================================================================
  function initSidebar() {
    if (!elements.sidebar || !elements.sidebarToggleBtn) return;
    if (state.sidebarCollapsed) {
      elements.sidebar.classList.add('collapsed');
      elements.sidebarToggleBtn.innerHTML = '▶';
      elements.sidebarToggleBtn.title = 'Mở rộng Sidebar';
    } else {
      elements.sidebar.classList.remove('collapsed');
      elements.sidebarToggleBtn.innerHTML = '◀';
      elements.sidebarToggleBtn.title = 'Thu gọn Sidebar';
    }

    elements.sidebarToggleBtn.addEventListener('click', toggleSidebar);
  }

  function toggleSidebar() {
    if (!elements.sidebar || !elements.sidebarToggleBtn) return;
    state.sidebarCollapsed = !state.sidebarCollapsed;
    elements.sidebar.classList.toggle('collapsed', state.sidebarCollapsed);
    elements.sidebarToggleBtn.innerHTML = state.sidebarCollapsed ? '▶' : '◀';
    elements.sidebarToggleBtn.title = state.sidebarCollapsed ? 'Mở rộng Sidebar' : 'Thu gọn Sidebar';
    if (window.JsonStorage) {
      window.JsonStorage.set('sidebarCollapsed', state.sidebarCollapsed);
    }
  }

  // =========================================================================
  // 6. VIEW SWITCHING (IF SEPARATE VIEWS EXIST)
  // =========================================================================
  function switchView(viewName) {
    state.currentView = viewName;

    if (viewName === 'tools') {
      if (elements.viewToolsManager) elements.viewToolsManager.classList.add('active');
      if (elements.viewSettings) elements.viewSettings.classList.remove('active');
      if (elements.navTools) elements.navTools.classList.add('active');
      if (elements.navSettings) elements.navSettings.classList.remove('active');
      renderToolsGrid();
    } else if (viewName === 'settings') {
      const settingSection = document.getElementById('settingSection');
      if (settingSection) {
        settingSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
        return;
      }
      if (elements.viewToolsManager) elements.viewToolsManager.classList.remove('active');
      if (elements.viewSettings) elements.viewSettings.classList.add('active');
      if (elements.navTools) elements.navTools.classList.remove('active');
      if (elements.navSettings) elements.navSettings.classList.add('active');
      renderSettingsView();
    }
  }

  // =========================================================================
  // 7. RENDER PINNED TOOLS IN SIDEBAR
  // =========================================================================
  function renderSidebarPinnedTools() {
    if (!elements.sidebarPinnedList) return;
    elements.sidebarPinnedList.innerHTML = '';
    const allTools = getAllTools();
    const pinnedTools = state.pinnedToolIds
      .map(id => allTools.find(t => t.id === id))
      .filter(Boolean);

    if (elements.sidebarPinnedCount) {
      elements.sidebarPinnedCount.textContent = `(${pinnedTools.length})`;
    }
    if (elements.statPinnedTools) {
      elements.statPinnedTools.textContent = pinnedTools.length;
    }

    if (pinnedTools.length === 0) {
      elements.sidebarPinnedList.innerHTML = `
        <div class="pinned-empty-hint">
          Chưa ghim tool nào<br>
          <a href="javascript:void(0)" id="linkGoSettings" style="color: var(--primary); text-decoration: underline; font-size: 11px;">Vào Cài đặt để ghim</a>
        </div>
      `;
      const link = document.getElementById('linkGoSettings');
      if (link) {
        link.addEventListener('click', () => switchView('settings'));
      }
      return;
    }

    pinnedTools.forEach(tool => {
      const item = document.createElement('a');
      item.className = `pinned-tool-item ${tool.isPrimary ? 'primary-tool' : ''}`;
      item.href = 'javascript:void(0);';
      item.setAttribute('data-tooltip', `${tool.name} (${tool.categoryLabel})`);
      item.title = tool.name;

      item.innerHTML = `
        <span class="pinned-tool-icon">${tool.icon}</span>
        <span class="pinned-tool-name">${escapeHtml(tool.name)}</span>
      `;

      item.addEventListener('click', (e) => {
        e.preventDefault();
        launchTool(tool);
      });

      elements.sidebarPinnedList.appendChild(item);
    });
  }

  // =========================================================================
  // 7.1. DYNAMIC CATEGORY PILLS RENDERER
  // =========================================================================
  function renderCategoryPills() {
    if (!elements.categoryPills) return;
    const allTools = getAllTools();

    // Map các danh mục và đếm số lượng
    const catMap = new Map();
    // Khởi tạo các danh mục tiêu chuẩn
    catMap.set('data', { label: 'Xử lý Dữ liệu Excel', count: 0 });
    catMap.set('ai', { label: 'Trợ lý AI & Chatbot', count: 0 });
    catMap.set('web', { label: 'Tiện ích Web', count: 0 });

    allTools.forEach(tool => {
      const cat = tool.category || 'other';
      const label = tool.categoryLabel || 'Tiện ích khác';
      if (!catMap.has(cat)) {
        catMap.set(cat, { label: label, count: 0 });
      }
      catMap.get(cat).count++;
    });

    let pillsHtml = `
      <button class="category-pill ${state.selectedCategory === 'all' ? 'active' : ''}" data-category="all">
        Tất cả (${allTools.length})
      </button>
    `;

    catMap.forEach((info, key) => {
      if (info.count > 0 || key === 'data') {
        const isActive = state.selectedCategory === key ? 'active' : '';
        pillsHtml += `
          <button class="category-pill ${isActive}" data-category="${key}">
            ${escapeHtml(info.label)} (${info.count})
          </button>
        `;
      }
    });

    pillsHtml += `
      <a href="#settingSection" class="category-pill" id="pillSettingsLink" style="text-decoration:none;">
        ⚙️ Cài Đặt Giao Diện
      </a>
    `;

    elements.categoryPills.innerHTML = pillsHtml;
  }

  // =========================================================================
  // 8. TOOLS MANAGER SCREEN - SEARCH, FILTER & PAGINATION
  // =========================================================================
  function getFilteredAndSortedTools() {
    let list = getAllTools();

    // Filter by Search Query
    if (state.searchQuery.trim()) {
      const q = state.searchQuery.trim().toLowerCase();
      list = list.filter(tool =>
        (tool.name && tool.name.toLowerCase().includes(q)) ||
        (tool.desc && tool.desc.toLowerCase().includes(q)) ||
        (tool.categoryLabel && tool.categoryLabel.toLowerCase().includes(q)) ||
        (tool.tags && tool.tags.some(t => t.toLowerCase().includes(q))) ||
        (tool.url && tool.url.toLowerCase().includes(q))
      );
    }

    // Filter by Category
    if (state.selectedCategory !== 'all') {
      list = list.filter(tool => tool.category === state.selectedCategory);
    }

    // Sort
    if (state.sortMode === 'az') {
      list.sort((a, b) => a.name.localeCompare(b.name, 'vi'));
    } else if (state.sortMode === 'za') {
      list.sort((a, b) => b.name.localeCompare(a.name, 'vi'));
    } else if (state.sortMode === 'popular') {
      // Primary tools first, then custom tools, then rest
      list.sort((a, b) => {
        if (a.isPrimary && !b.isPrimary) return -1;
        if (!a.isPrimary && b.isPrimary) return 1;
        const aPinned = state.pinnedToolIds.includes(a.id);
        const bPinned = state.pinnedToolIds.includes(b.id);
        if (aPinned && !bPinned) return -1;
        if (!aPinned && bPinned) return 1;
        return a.name.localeCompare(b.name, 'vi');
      });
    }

    return list;
  }

  function renderToolsGrid() {
    const allTools = getAllTools();
    const customCount = allTools.filter(t => t.isCustom).length;
    const filteredTools = getFilteredAndSortedTools();
    const totalItems = filteredTools.length;
    const totalPages = Math.ceil(totalItems / state.pageSize) || 1;

    if (state.currentPage > totalPages) {
      state.currentPage = totalPages;
    }
    if (state.currentPage < 1) {
      state.currentPage = 1;
    }

    // Slice for current page
    const startIndex = (state.currentPage - 1) * state.pageSize;
    const paginatedTools = filteredTools.slice(startIndex, startIndex + state.pageSize);

    // Update stats
    if (elements.statTotalTools) {
      elements.statTotalTools.textContent = allTools.length;
    }
    const statCustom = document.getElementById('statCustomTools');
    if (statCustom) {
      statCustom.textContent = customCount;
    }

    // Render Grid
    elements.toolsGrid.innerHTML = '';

    if (paginatedTools.length === 0) {
      elements.toolsGrid.innerHTML = `
        <div class="empty-state">
          <div class="empty-icon">🔍</div>
          <h3 class="empty-title">Không tìm thấy công cụ phù hợp</h3>
          <p class="empty-desc">Không có công cụ nào khớp với từ khóa "<b>${escapeHtml(state.searchQuery)}</b>". Thử tìm kiếm với từ khóa khác hoặc bấm nút thêm tool mới.</p>
          <div style="display: flex; gap: 10px; justify-content: center; flex-wrap: wrap;">
            <button class="btn-launch-tool" id="btnResetSearch">Xóa bộ lọc tìm kiếm</button>
            <button class="btn-add-tool-cta" id="btnEmptyAddTool">＋ Thêm Tool Mới</button>
          </div>
        </div>
      `;
      const btnReset = document.getElementById('btnResetSearch');
      if (btnReset) {
        btnReset.addEventListener('click', () => {
          state.searchQuery = '';
          state.selectedCategory = 'all';
          elements.searchInput.value = '';
          elements.searchClearBtn.style.display = 'none';
          renderCategoryPills();
          renderToolsGrid();
        });
      }
      const btnEmptyAdd = document.getElementById('btnEmptyAddTool');
      if (btnEmptyAdd) {
        btnEmptyAdd.addEventListener('click', () => openCustomToolModal());
      }
      renderPagination(0, 1);
      return;
    }

    paginatedTools.forEach(tool => {
      const isPinned = state.pinnedToolIds.includes(tool.id);
      const card = document.createElement('div');
      card.className = `tool-card ${tool.isPrimary ? 'primary-tool' : ''}`;

      const badgeClass = tool.isCustom ? 'badge-custom' : (tool.isPrimary ? 'badge-primary' : 'badge-category');
      const badgeText = tool.badge || (tool.isCustom ? 'Tùy Chỉnh' : tool.categoryLabel);

      card.innerHTML = `
        <div>
          <div class="tool-card-header">
            <div class="tool-icon-box">${tool.icon || '⚡'}</div>
            <div class="tool-badge-row">
              <span class="badge ${badgeClass}">${escapeHtml(badgeText)}</span>
              <div class="btn-card-action-group">
                ${tool.isCustom ? `
                  <button class="btn-card-action-small btn-card-edit" data-action="edit" data-tool-id="${tool.id}" title="Chỉnh sửa công cụ này">✏️</button>
                  <button class="btn-card-action-small btn-card-delete" data-action="delete" data-tool-id="${tool.id}" title="Xóa công cụ này khỏi File JSON">🗑️</button>
                ` : ''}
                <button class="btn-card-pin ${isPinned ? 'pinned' : ''}" data-tool-id="${tool.id}" title="${isPinned ? 'Bỏ ghim khỏi Sidebar' : 'Ghim vào Sidebar'}">
                  ${isPinned ? '★' : '☆'}
                </button>
              </div>
            </div>
          </div>
          <div class="tool-card-body">
            <h3 class="tool-title">${escapeHtml(tool.name)}</h3>
            <p class="tool-desc">${escapeHtml(tool.desc)}</p>
            <div class="tool-tags">
              ${(tool.tags || []).map(t => `<span class="tag-item">#${escapeHtml(t)}</span>`).join('')}
            </div>
          </div>
        </div>
        ${tool.id === 'superset-dashboard-builder' || (tool.url && tool.url.toLowerCase().includes('dashboard')) || (tool.name && tool.name.toLowerCase().includes('dashboard')) ? `
          <div class="tool-card-footer" style="flex-direction: column; align-items: stretch; gap: 10px;">
            <div class="tool-meta" style="justify-content: space-between; width: 100%;">
              <span>🏷️ ${escapeHtml(tool.categoryLabel || 'Tiện ích')}</span>
              <span style="font-size:11px; color:var(--text-dim);">📊 100% Client-side JSON</span>
            </div>
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 8px; width: 100%;">
              <button class="btn-launch-tool" data-action="launch-new-dashboard" data-tool-id="${tool.id}" style="justify-content: center; padding: 10px 8px; font-size: 12px; font-weight: 700; background: linear-gradient(135deg, #0284c7, #0369a1); text-align: center; border-radius: 8px;" title="Tạo Dashboard mới để vào màn hình edit">
                ✨ Tạo Dashboard Mới
              </button>
              <button class="btn-launch-tool outline" data-action="launch-import-dashboard" data-tool-id="${tool.id}" style="justify-content: center; padding: 10px 8px; font-size: 12px; font-weight: 700; border: 1.5px solid #38bdf8; color: #38bdf8; background: rgba(56, 189, 248, 0.08); text-align: center; border-radius: 8px;" title="Import Dashboard: Mở thư mục chứa file JSON cấu hình và Excel để nạp dữ liệu tái tạo lại Dashboard">
                📥 Import Dashboard
              </button>
            </div>
          </div>
        ` : `
          <div class="tool-card-footer">
            <div class="tool-meta">
              <span>🏷️ ${escapeHtml(tool.categoryLabel || 'Tiện ích')}</span>
              ${tool.url ? `<span style="font-size:11px; color:var(--text-dim); max-width:160px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;" title="${escapeHtml(tool.url)}">🔗 ${escapeHtml(tool.url)}</span>` : ''}
            </div>
            <button class="btn-launch-tool ${tool.isPrimary || tool.isCustom ? '' : 'outline'}" data-action="launch" data-tool-id="${tool.id}">
              ${tool.isPrimary || tool.isCustom ? '▶ Mở công cụ' : 'Khám phá tool'}
            </button>
          </div>
        `}
      `;

      // Quick Pin Toggle
      const pinBtn = card.querySelector('.btn-card-pin');
      if (pinBtn) {
        pinBtn.addEventListener('click', (e) => {
          e.stopPropagation();
          toggleToolPin(tool.id);
        });
      }

      // Edit Button (cho custom tool)
      const editBtn = card.querySelector('.btn-card-edit');
      if (editBtn) {
        editBtn.addEventListener('click', (e) => {
          e.stopPropagation();
          openCustomToolModal(tool);
        });
      }

      // Delete Button (cho custom tool)
      const deleteBtn = card.querySelector('.btn-card-delete');
      if (deleteBtn) {
        deleteBtn.addEventListener('click', (e) => {
          e.stopPropagation();
          deleteCustomTool(tool.id);
        });
      }

      // Dashboard Button 1: Tạo Dashboard Mới
      const newDashBtn = card.querySelector('[data-action="launch-new-dashboard"]');
      if (newDashBtn) {
        newDashBtn.addEventListener('click', () => {
          window.location.href = 'dashboard_builder.html?mode=new';
        });
      }

      // Dashboard Button 2: Import Dashboard (Mở thư mục chứa file JSON và Excel)
      const importDashBtn = card.querySelector('[data-action="launch-import-dashboard"]') || card.querySelector('[data-action="launch-import-folder"]');
      if (importDashBtn) {
        importDashBtn.addEventListener('click', async (e) => {
          e.preventDefault();
          await handleHubImportDashboard();
        });
      }

      // Launch button (cho các tool tiêu chuẩn)
      const launchBtn = card.querySelector('[data-action="launch"]');
      if (launchBtn) {
        launchBtn.addEventListener('click', () => {
          launchTool(tool);
        });
      }

      elements.toolsGrid.appendChild(card);
    });

    renderPagination(totalItems, totalPages);
  }

  // =========================================================================
  // 9. PAGINATION RENDERER
  // =========================================================================
  function renderPagination(totalItems, totalPages) {
    if (totalItems === 0) {
      elements.paginationWrapper.style.display = 'none';
      return;
    }
    elements.paginationWrapper.style.display = 'flex';

    const startItem = (state.currentPage - 1) * state.pageSize + 1;
    const endItem = Math.min(state.currentPage * state.pageSize, totalItems);

    elements.paginationInfo.textContent = `Hiển thị ${startItem} - ${endItem} trong tổng số ${totalItems} công cụ (Trang ${state.currentPage}/${totalPages})`;

    elements.paginationControls.innerHTML = '';

    // Previous Button
    const prevBtn = document.createElement('button');
    prevBtn.className = 'page-btn';
    prevBtn.innerHTML = '‹ Trước';
    prevBtn.disabled = state.currentPage <= 1;
    prevBtn.addEventListener('click', () => {
      if (state.currentPage > 1) {
        state.currentPage--;
        renderToolsGrid();
      }
    });
    elements.paginationControls.appendChild(prevBtn);

    // Numbered buttons
    for (let i = 1; i <= totalPages; i++) {
      // Logic for ellipsis if many pages
      if (totalPages > 6 && Math.abs(i - state.currentPage) > 2 && i !== 1 && i !== totalPages) {
        if (i === 2 || i === totalPages - 1) {
          const ellipsis = document.createElement('span');
          ellipsis.className = 'page-btn';
          ellipsis.style.border = 'none';
          ellipsis.style.background = 'transparent';
          ellipsis.textContent = '...';
          elements.paginationControls.appendChild(ellipsis);
        }
        continue;
      }

      const numBtn = document.createElement('button');
      numBtn.className = `page-btn ${i === state.currentPage ? 'active' : ''}`;
      numBtn.textContent = i;
      numBtn.addEventListener('click', () => {
        state.currentPage = i;
        renderToolsGrid();
      });
      elements.paginationControls.appendChild(numBtn);
    }

    // Next Button
    const nextBtn = document.createElement('button');
    nextBtn.className = 'page-btn';
    nextBtn.innerHTML = 'Tiếp ›';
    nextBtn.disabled = state.currentPage >= totalPages;
    nextBtn.addEventListener('click', () => {
      if (state.currentPage < totalPages) {
        state.currentPage++;
        renderToolsGrid();
      }
    });
    elements.paginationControls.appendChild(nextBtn);
  }

  // =========================================================================
  // 10. SETTINGS VIEW - PINNING & REORDERING TOOLS
  // =========================================================================
  function renderSettingsView() {
    renderPinnedReorderList();
    renderAvailableToolsList();
  }

  function renderPinnedReorderList() {
    elements.pinnedReorderList.innerHTML = '';
    const allTools = getAllTools();
    const pinnedTools = state.pinnedToolIds
      .map(id => allTools.find(t => t.id === id))
      .filter(Boolean);

    elements.pinnedCountBadge.textContent = `${pinnedTools.length} tools`;

    if (pinnedTools.length === 0) {
      elements.pinnedReorderList.innerHTML = `
        <div class="empty-state" style="padding: 30px 10px;">
          <div class="empty-icon" style="font-size: 32px;">📌</div>
          <h4 style="color: #fff; font-size: 14px; margin-bottom: 6px;">Chưa gán tool nào vào sidebar</h4>
          <p style="color: var(--text-dim); font-size: 12px;">Hãy bấm nút "+ Ghim vào Sidebar" ở danh sách bên phải để thêm các công cụ thường dùng.</p>
        </div>
      `;
      return;
    }

    pinnedTools.forEach((tool, index) => {
      const item = document.createElement('div');
      item.className = 'reorder-item';
      item.draggable = true;
      item.dataset.toolId = tool.id;
      item.dataset.index = index;

      item.innerHTML = `
        <span class="drag-handle" title="Giữ và kéo thả chuột để đổi thứ tự">⋮⋮</span>
        <span class="reorder-index">#${index + 1}</span>
        <span class="reorder-icon">${tool.icon || '⚡'}</span>
        <div class="reorder-info">
          <div class="reorder-name">${escapeHtml(tool.name)}</div>
          <div class="reorder-sub">${escapeHtml(tool.categoryLabel || '')} ${tool.isPrimary ? '• Tool cốt lõi' : (tool.isCustom ? '• Tool tự thêm' : '')}</div>
        </div>
        <div class="reorder-actions">
          <button class="btn-move btn-move-up" title="Di chuyển lên" ${index === 0 ? 'disabled' : ''}>▲</button>
          <button class="btn-move btn-move-down" title="Di chuyển xuống" ${index === pinnedTools.length - 1 ? 'disabled' : ''}>▼</button>
          <button class="btn-unpin" title="Bỏ ghim khỏi sidebar">✕</button>
        </div>
      `;

      // Move Up
      const btnUp = item.querySelector('.btn-move-up');
      btnUp.addEventListener('click', (e) => {
        e.stopPropagation();
        moveToolOrder(index, index - 1);
      });

      // Move Down
      const btnDown = item.querySelector('.btn-move-down');
      btnDown.addEventListener('click', (e) => {
        e.stopPropagation();
        moveToolOrder(index, index + 1);
      });

      // Unpin
      const btnUnpin = item.querySelector('.btn-unpin');
      btnUnpin.addEventListener('click', (e) => {
        e.stopPropagation();
        toggleToolPin(tool.id, false);
      });

      // Drag and Drop support
      setupDragAndDrop(item, index);

      elements.pinnedReorderList.appendChild(item);
    });
  }

  // Drag and Drop implementation
  let draggedIndex = null;

  function setupDragAndDrop(element, index) {
    element.addEventListener('dragstart', (e) => {
      draggedIndex = index;
      element.classList.add('dragging');
      e.dataTransfer.effectAllowed = 'move';
      e.dataTransfer.setData('text/plain', index);
    });

    element.addEventListener('dragend', () => {
      element.classList.remove('dragging');
      draggedIndex = null;
    });

    element.addEventListener('dragover', (e) => {
      e.preventDefault();
      e.dataTransfer.dropEffect = 'move';
    });

    element.addEventListener('drop', (e) => {
      e.preventDefault();
      const targetIndex = index;
      if (draggedIndex !== null && draggedIndex !== targetIndex) {
        moveToolOrder(draggedIndex, targetIndex);
      }
    });
  }

  function moveToolOrder(fromIndex, toIndex) {
    if (fromIndex < 0 || fromIndex >= state.pinnedToolIds.length) return;
    if (toIndex < 0 || toIndex >= state.pinnedToolIds.length) return;

    const item = state.pinnedToolIds.splice(fromIndex, 1)[0];
    state.pinnedToolIds.splice(toIndex, 0, item);

    savePinnedTools();
    renderSidebarPinnedTools();
    renderSettingsView();
  }

  function renderAvailableToolsList() {
    elements.availableToolsList.innerHTML = '';
    const q = state.availableSearchQuery.trim().toLowerCase();
    const allTools = getAllTools();

    const filtered = allTools.filter(tool => {
      if (!q) return true;
      return (tool.name && tool.name.toLowerCase().includes(q)) ||
        (tool.desc && tool.desc.toLowerCase().includes(q)) ||
        (tool.categoryLabel && tool.categoryLabel.toLowerCase().includes(q));
    });

    if (filtered.length === 0) {
      elements.availableToolsList.innerHTML = `
        <div style="padding: 20px; text-align: center; color: var(--text-dim); font-size: 13px;">
          Không tìm thấy công cụ nào với từ khóa "${escapeHtml(q)}"
        </div>
      `;
      return;
    }

    filtered.forEach(tool => {
      const isPinned = state.pinnedToolIds.includes(tool.id);
      const row = document.createElement('div');
      row.className = 'available-tool-item';

      row.innerHTML = `
        <div class="available-tool-left">
          <span class="available-tool-icon">${tool.icon || '⚡'}</span>
          <div style="min-width: 0;">
            <div class="available-tool-title">${escapeHtml(tool.name)}</div>
            <div class="available-tool-desc">${escapeHtml(tool.categoryLabel || 'Tiện ích')} ${tool.isCustom ? '• Tool tự thêm' : ''}</div>
          </div>
        </div>
        <button class="btn-pin-toggle ${isPinned ? 'is-pinned' : ''}">
          ${isPinned ? '✓ Đã ghim' : '＋ Ghim vào Sidebar'}
        </button>
      `;

      const btnToggle = row.querySelector('.btn-pin-toggle');
      btnToggle.addEventListener('click', () => {
        toggleToolPin(tool.id);
      });

      elements.availableToolsList.appendChild(row);
    });
  }

  function toggleToolPin(toolId, explicitState) {
    const index = state.pinnedToolIds.indexOf(toolId);
    const shouldPin = explicitState !== undefined ? explicitState : index === -1;

    if (shouldPin && index === -1) {
      state.pinnedToolIds.push(toolId);
    } else if (!shouldPin && index !== -1) {
      state.pinnedToolIds.splice(index, 1);
    }

    savePinnedTools();
    renderSidebarPinnedTools();

    if (state.currentView === 'settings') {
      renderSettingsView();
    } else if (state.currentView === 'tools') {
      renderToolsGrid();
    }
  }

  // =========================================================================
  // 11. TOOL LAUNCHER & MODAL DEMO
  // =========================================================================
  function launchTool(tool) {
    if (tool.url) {
      const url = tool.url.trim();
      const isExternal = url.startsWith('http://') || url.startsWith('https://');

      // Mở trong tab mới nếu là website ngoài hoặc người dùng đã chọn mở tab mới
      if (tool.openInNewTab || isExternal) {
        window.open(url, '_blank');
        showToast(`Đang mở "${tool.name}" trong tab mới...`);
        return;
      }

      // Mở file HTML nội bộ hoặc liên kết trong workspace
      window.location.href = url;
      return;
    }

    // Với tool catalog chưa có URL: hiển thị modal preview
    elements.modalTitle.innerHTML = `${tool.icon} ${tool.name}`;
    elements.modalDesc.textContent = tool.desc;
    elements.modalTags.innerHTML = `
      <div style="margin-bottom: 8px; color: var(--text-dim); font-size: 11px;">DANH MỤC: <b>${tool.categoryLabel}</b></div>
      <div class="tool-tags" style="margin-bottom: 12px;">
        ${(tool.tags || []).map(t => `<span class="tag-item">#${t}</span>`).join('')}
      </div>
      <div style="background: rgba(255,255,255,0.04); padding: 12px; border-radius: 8px; border: 1px dashed var(--border-subtle); font-size: 12.5px; color: var(--text-muted);">
        💡 <b>Trạng thái:</b> Module công cụ xử lý dữ liệu. Nhấn nút bên dưới để mở công cụ.
      </div>
    `;

    elements.modalActionBtn.textContent = '🚀 Mở Công Cụ Này';
    elements.modalActionBtn.onclick = () => {
      if (tool.url) window.location.href = tool.url;
      closeModal();
    };

    elements.modalOverlay.classList.add('open');
  }

  function closeModal() {
    elements.modalOverlay.classList.remove('open');
  }

  // =========================================================================
  // 11.1. CUSTOM TOOL MODAL HANDLERS (THÊM / CHỈNH SỬA / XÓA)
  // =========================================================================
  const PRESET_TEMPLATES = {
    chatgpt: {
      name: 'ChatGPT',
      url: 'https://chatgpt.com',
      icon: '🤖',
      category: 'ai',
      tags: 'AI, OpenAI, Chatbot',
      desc: 'Trợ lý trí tuệ nhân tạo ChatGPT từ OpenAI hỗ trợ giải đáp, xử lý dữ liệu và lập trình.',
      newTab: true
    },
    gemini: {
      name: 'Google Gemini',
      url: 'https://gemini.google.com',
      icon: '🔮',
      category: 'ai',
      tags: 'AI, Gemini, Google',
      desc: 'Mô hình AI từ Google hỗ trợ tra cứu, phân tích tài liệu và viết nội dung thông minh.',
      newTab: true
    },
    claude: {
      name: 'Claude AI',
      url: 'https://claude.ai',
      icon: '🧠',
      category: 'ai',
      tags: 'AI, Claude, Anthropic',
      desc: 'Trợ lý AI Claude từ Anthropic chuyên sâu về tư duy logic, phân tích mã nguồn và văn bản.',
      newTab: true
    },
    localhtml: {
      name: 'Tool HTML Nội Bộ',
      url: 'custom_tool.html',
      icon: '📄',
      category: 'data',
      tags: 'HTML, Offline, File Cục Bộ',
      desc: 'Công cụ lập trình offline mở trực tiếp file HTML được lưu trong thư mục máy tính.',
      newTab: false
    }
  };

  function openCustomToolModal(toolToEdit = null) {
    const modal = document.getElementById('customToolModalOverlay');
    const heading = document.getElementById('customToolModalHeading');
    const icon = document.getElementById('customToolModalIcon');
    const editIdInput = document.getElementById('customToolEditId');
    const nameInput = document.getElementById('toolNameInput');
    const urlInput = document.getElementById('toolUrlInput');
    const iconInput = document.getElementById('toolIconInput');
    const catSelect = document.getElementById('toolCategorySelect');
    const tagsInput = document.getElementById('toolTagsInput');
    const descInput = document.getElementById('toolDescInput');
    const openNewTabCheckbox = document.getElementById('toolOpenNewTabCheckbox');
    const autoPinCheckbox = document.getElementById('toolAutoPinCheckbox');

    if (!modal) return;

    if (toolToEdit) {
      heading.textContent = 'Chỉnh Sửa Công Cụ';
      icon.textContent = '✏️';
      editIdInput.value = toolToEdit.id;
      nameInput.value = toolToEdit.name || '';
      urlInput.value = toolToEdit.url || '';
      iconInput.value = toolToEdit.icon || '⚡';
      catSelect.value = toolToEdit.category || 'ai';
      tagsInput.value = (toolToEdit.tags || []).join(', ');
      descInput.value = toolToEdit.desc || '';
      openNewTabCheckbox.checked = toolToEdit.openInNewTab !== false;
      autoPinCheckbox.checked = state.pinnedToolIds.includes(toolToEdit.id);
    } else {
      heading.textContent = 'Thêm Công Cụ Mới (Link / HTML)';
      icon.textContent = '➕';
      editIdInput.value = '';
      nameInput.value = '';
      urlInput.value = '';
      iconInput.value = '⚡';
      catSelect.value = 'ai';
      tagsInput.value = '';
      descInput.value = '';
      openNewTabCheckbox.checked = true;
      autoPinCheckbox.checked = true;
    }

    modal.classList.add('open');
    nameInput.focus();
  }

  function closeCustomToolModal() {
    const modal = document.getElementById('customToolModalOverlay');
    if (modal) modal.classList.remove('open');
  }

  function handleCustomToolSubmit(e) {
    if (e) e.preventDefault();

    const editId = document.getElementById('customToolEditId').value.trim();
    const name = document.getElementById('toolNameInput').value.trim();
    let url = document.getElementById('toolUrlInput').value.trim();
    const icon = document.getElementById('toolIconInput').value.trim() || '⚡';
    const category = document.getElementById('toolCategorySelect').value;
    const catLabelMap = {
      ai: 'Trợ lý AI & Chatbot',
      data: 'Xử lý Dữ liệu',
      web: 'Tiện ích Web & Tra cứu',
      report: 'Báo cáo & Thống kê',
      other: 'Công cụ Tiện ích khác'
    };
    const categoryLabel = catLabelMap[category] || 'Tiện ích khác';
    const tagsRaw = document.getElementById('toolTagsInput').value.trim();
    const tags = tagsRaw ? tagsRaw.split(',').map(t => t.trim()).filter(Boolean) : ['Custom'];
    const desc = document.getElementById('toolDescInput').value.trim() || 'Công cụ tùy chỉnh được thêm vào trình duyệt.';
    const openInNewTab = document.getElementById('toolOpenNewTabCheckbox').checked;
    const autoPin = document.getElementById('toolAutoPinCheckbox').checked;

    if (!name) {
      alert('Vui lòng nhập tên công cụ!');
      document.getElementById('toolNameInput').focus();
      return;
    }
    if (!url) {
      alert('Vui lòng nhập đường dẫn URL hoặc file HTML!');
      document.getElementById('toolUrlInput').focus();
      return;
    }

    // Tự động thêm https:// nếu người dùng nhập dạng domain (vd: chatgpt.com, google.com)
    if (!url.startsWith('http://') && !url.startsWith('https://') && !url.startsWith('file://') && !url.endsWith('.html') && !url.includes('/')) {
      if (url.includes('.')) {
        url = 'https://' + url;
      }
    }

    const customTools = loadCustomTools();
    let targetToolId = editId;

    if (editId) {
      // Cập nhật tool đã có
      const index = customTools.findIndex(t => t.id === editId);
      if (index !== -1) {
        customTools[index] = {
          ...customTools[index],
          name,
          url,
          icon,
          category,
          categoryLabel,
          tags,
          desc,
          openInNewTab,
          updatedAt: Date.now()
        };
      }
      showToast(`Đã cập nhật công cụ "${name}" thành công!`);
    } else {
      // Thêm mới
      targetToolId = 'custom-' + Date.now();
      const newTool = {
        id: targetToolId,
        name,
        url,
        icon,
        category,
        categoryLabel,
        tags,
        desc,
        badge: 'Tùy Chỉnh • ' + (url.startsWith('http') ? 'Web Link' : 'File Cục Bộ'),
        isPrimary: false,
        isCustom: true,
        openInNewTab,
        createdAt: Date.now()
      };
      customTools.unshift(newTool);
      showToast(`Đã thêm công cụ "${name}" vào File JSON thành công!`);
    }

    saveCustomTools(customTools);

    // Xử lý Pin
    if (autoPin && !state.pinnedToolIds.includes(targetToolId)) {
      state.pinnedToolIds.push(targetToolId);
      savePinnedTools();
    } else if (!autoPin && state.pinnedToolIds.includes(targetToolId)) {
      const pIdx = state.pinnedToolIds.indexOf(targetToolId);
      state.pinnedToolIds.splice(pIdx, 1);
      savePinnedTools();
    }

    closeCustomToolModal();
    renderSidebarPinnedTools();
    renderCategoryPills();
    renderToolsGrid();
    if (state.currentView === 'settings') {
      renderSettingsView();
    }
  }

  function deleteCustomTool(toolId) {
    const customTools = loadCustomTools();
    const tool = customTools.find(t => t.id === toolId);
    if (!tool) return;

    const confirmMsg = `Bạn có chắc chắn muốn xóa công cụ "${tool.name}" khỏi File JSON không?`;
    if (!confirm(confirmMsg)) return;

    // Lọc bỏ tool khỏi danh sách custom
    const updated = customTools.filter(t => t.id !== toolId);
    saveCustomTools(updated);

    // Gỡ khỏi pinned nếu có
    const pinIndex = state.pinnedToolIds.indexOf(toolId);
    if (pinIndex !== -1) {
      state.pinnedToolIds.splice(pinIndex, 1);
      savePinnedTools();
      renderSidebarPinnedTools();
    }

    renderCategoryPills();
    renderToolsGrid();
    if (state.currentView === 'settings') {
      renderSettingsView();
    }
    showToast(`Đã xóa công cụ "${tool.name}" thành công!`);
  }

  // =========================================================================
  // 12. TOAST NOTIFICATION
  // =========================================================================
  let toastTimer = null;
  function showToast(message) {
    if (toastTimer) clearTimeout(toastTimer);
    if (elements.toastMessage && elements.toastNotification) {
      elements.toastMessage.textContent = message;
      elements.toastNotification.classList.add('show');
      toastTimer = setTimeout(() => {
        elements.toastNotification.classList.remove('show');
      }, 2400);
    }
  }
  window.showToast = showToast;

  // Utility: Escape HTML
  function escapeHtml(str) {
    if (!str) return '';
    return str.replace(/[&<>"']/g, function (m) {
      return {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#39;'
      }[m];
    });
  }

  // =========================================================================
  // 13. EVENT LISTENERS SETUP
  // =========================================================================
  function setupEventListeners() {
    // Navigation items (if separate views exist)
    if (elements.navTools) {
      elements.navTools.addEventListener('click', () => switchView('tools'));
    }
    if (elements.navSettings) {
      elements.navSettings.addEventListener('click', () => switchView('settings'));
    }

    // Smooth scroll buttons for Settings
    const btnScrollSettings = document.getElementById('btnScrollSettings');
    if (btnScrollSettings) {
      btnScrollSettings.addEventListener('click', (e) => {
        e.preventDefault();
        const sec = document.getElementById('settingSection');
        if (sec) sec.scrollIntoView({ behavior: 'smooth', block: 'start' });
      });
    }

    const pillSettingsLink = document.getElementById('pillSettingsLink');
    if (pillSettingsLink) {
      pillSettingsLink.addEventListener('click', (e) => {
        e.preventDefault();
        const sec = document.getElementById('settingSection');
        if (sec) sec.scrollIntoView({ behavior: 'smooth', block: 'start' });
      });
    }

    // Search bar
    if (elements.searchInput) {
      elements.searchInput.addEventListener('input', (e) => {
        state.searchQuery = e.target.value;
        state.currentPage = 1;
        if (elements.searchClearBtn) elements.searchClearBtn.style.display = state.searchQuery ? 'block' : 'none';
        renderToolsGrid();
      });
    }

    if (elements.searchClearBtn) {
      elements.searchClearBtn.addEventListener('click', () => {
        if (elements.searchInput) elements.searchInput.value = '';
        state.searchQuery = '';
        state.currentPage = 1;
        elements.searchClearBtn.style.display = 'none';
        renderToolsGrid();
      });
    }

    // Category pills click (delegation)
    if (elements.categoryPills) {
      elements.categoryPills.addEventListener('click', (e) => {
        const pill = e.target.closest('.category-pill');
        if (!pill || pill.id === 'pillSettingsLink') return;
        document.querySelectorAll('.category-pill').forEach(p => p.classList.remove('active'));
        pill.classList.add('active');
        state.selectedCategory = pill.dataset.category || 'all';
        state.currentPage = 1;
        renderToolsGrid();
      });
    }

    // Sort select
    if (elements.sortSelect) {
      elements.sortSelect.addEventListener('change', (e) => {
        state.sortMode = e.target.value;
        renderToolsGrid();
      });
    }

    // Page size select
    if (elements.pageSizeSelect) {
      elements.pageSizeSelect.addEventListener('change', (e) => {
        state.pageSize = parseInt(e.target.value, 10) || 6;
        state.currentPage = 1;
        renderToolsGrid();
      });
    }

    // Settings available tools search
    if (elements.availableSearchInput) {
      elements.availableSearchInput.addEventListener('input', (e) => {
        state.availableSearchQuery = e.target.value;
        renderAvailableToolsList();
      });
    }

    // Reset default pinned tools
    if (elements.btnResetDefault) {
      elements.btnResetDefault.addEventListener('click', () => {
        if (confirm('Bạn có chắc muốn khôi phục danh sách tools ghim trên sidebar về mặc định?')) {
          state.pinnedToolIds = [...DEFAULT_PINNED_IDS];
          savePinnedTools();
          renderSidebarPinnedTools();
          renderSettingsView();
        }
      });
    }

    // Preview Modal close
    if (elements.modalCloseBtn) {
      elements.modalCloseBtn.addEventListener('click', closeModal);
    }
    if (elements.modalOverlay) {
      elements.modalOverlay.addEventListener('click', (e) => {
        if (e.target === elements.modalOverlay) closeModal();
      });
    }

    // =========================================================================
    // CUSTOM TOOL MODAL LISTENERS
    // =========================================================================
    const btnOpenAddNav = document.getElementById('btnOpenAddToolModalNav');
    if (btnOpenAddNav) {
      btnOpenAddNav.addEventListener('click', () => openCustomToolModal());
    }

    const btnOpenAddCta = document.getElementById('btnOpenAddToolModal');
    if (btnOpenAddCta) {
      btnOpenAddCta.addEventListener('click', () => openCustomToolModal());
    }

    const customToolModalCloseBtn = document.getElementById('customToolModalCloseBtn');
    if (customToolModalCloseBtn) {
      customToolModalCloseBtn.addEventListener('click', closeCustomToolModal);
    }

    const customToolCancelBtn = document.getElementById('customToolCancelBtn');
    if (customToolCancelBtn) {
      customToolCancelBtn.addEventListener('click', closeCustomToolModal);
    }

    const customModalOverlay = document.getElementById('customToolModalOverlay');
    if (customModalOverlay) {
      customModalOverlay.addEventListener('click', (e) => {
        if (e.target === customModalOverlay) closeCustomToolModal();
      });
    }

    // Form submit
    const customToolForm = document.getElementById('customToolForm');
    if (customToolForm) {
      customToolForm.addEventListener('submit', handleCustomToolSubmit);
    }

    // Preset chips
    document.querySelectorAll('.preset-chip').forEach(btn => {
      btn.addEventListener('click', () => {
        const key = btn.dataset.preset;
        const preset = PRESET_TEMPLATES[key];
        if (!preset) return;
        document.getElementById('toolNameInput').value = preset.name;
        document.getElementById('toolUrlInput').value = preset.url;
        document.getElementById('toolIconInput').value = preset.icon;
        document.getElementById('toolCategorySelect').value = preset.category;
        document.getElementById('toolTagsInput').value = preset.tags;
        document.getElementById('toolDescInput').value = preset.desc;
        document.getElementById('toolOpenNewTabCheckbox').checked = preset.newTab;
      });
    });

    // Quick icon options
    document.querySelectorAll('.icon-option').forEach(opt => {
      opt.addEventListener('click', () => {
        const ic = opt.dataset.icon;
        if (ic) document.getElementById('toolIconInput').value = ic;
      });
    });

    // Escape key
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        if (elements.modalOverlay && elements.modalOverlay.classList.contains('open')) {
          closeModal();
        }
        if (customModalOverlay && customModalOverlay.classList.contains('open')) {
          closeCustomToolModal();
        }
      }
    });

    // JSON Storage Controls (Xuất / Nạp / Liên kết thư mục data/)
    const btnExportJson = document.getElementById('btnExportJson');
    if (btnExportJson) {
      btnExportJson.addEventListener('click', () => {
        if (window.JsonStorage) {
          window.JsonStorage.exportJsonFile();
          showToast('Đã tải xuống file app_storage.json!');
        }
      });
    }

    const btnImportJson = document.getElementById('btnImportJson');
    const jsonFileInput = document.getElementById('jsonFileInput');
    if (btnImportJson && jsonFileInput) {
      btnImportJson.addEventListener('click', () => {
        jsonFileInput.value = '';
        jsonFileInput.click();
      });

      jsonFileInput.addEventListener('change', async (e) => {
        const file = e.target.files && e.target.files[0];
        if (!file) return;
        try {
          if (window.JsonStorage) {
            await window.JsonStorage.importJsonFile(file);
            showToast(`Đã nạp thành công dữ liệu từ file "${file.name}"!`);
          }
        } catch (err) {
          alert('Lỗi nạp file JSON: ' + err.message);
        }
      });
    }

    const btnLinkDataFolder = document.getElementById('btnLinkDataFolder');
    if (btnLinkDataFolder) {
      btnLinkDataFolder.addEventListener('click', async () => {
        if (window.JsonStorage) {
          const ok = await window.JsonStorage.linkDataDirectory();
          if (ok) {
            showToast('Đã liên kết với thư mục data/! Mọi thay đổi sẽ tự động lưu vào file JSON');
            const syncBadge = document.getElementById('storageSyncBadge');
            if (syncBadge) {
              syncBadge.innerHTML = `<span style="color: #4ec9b0;">●</span> Đã kết nối trực tiếp thư mục data/`;
            }
          }
        }
      });
    }

    // Lắng nghe sự kiện cập nhật dữ liệu từ JSON Storage
    window.addEventListener('json-storage-updated', () => {
      state.pinnedToolIds = loadPinnedTools();
      renderSidebarPinnedTools();
      renderCategoryPills();
      renderToolsGrid();
      if (state.currentView === 'settings') {
        renderSettingsView();
      }
    });

    window.addEventListener('json-storage-folder-linked', (e) => {
      const folderName = (e.detail && e.detail.folderName) || 'data';
      const syncBadge = document.getElementById('storageSyncBadge');
      if (syncBadge) {
        syncBadge.innerHTML = `<span style="color: #4ec9b0;">●</span> Đã kết nối folder ${folderName}/ (Tự động lưu)`;
      }
    });
  }

  // =========================================================================
  // 13.5. DASHBOARD IMPORT FROM FOLDER HANDLER
  // =========================================================================
  function storePendingImportInIndexedDB(data) {
    return new Promise((resolve) => {
      let isSettled = false;
      const settle = () => {
        if (!isSettled) {
          isSettled = true;
          resolve();
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
            const tx = db.transaction('pending_import', 'readwrite');
            const store = tx.objectStore('pending_import');
            store.put({ id: 'latest_import', ...data, timestamp: Date.now() });
            tx.oncomplete = () => settle();
            tx.onerror = () => settle();
          } catch (_) {
            settle();
          }
        };
        req.onerror = () => settle();
        req.onblocked = () => settle();
      } catch (_) {
        settle();
      }

      setTimeout(() => settle(), 1500);
    });
  }

  async function handleHubImportDashboard() {
    // 1. Thử dùng File System Access API (showDirectoryPicker) trực tiếp từ click event
    if (window.showDirectoryPicker) {
      try {
        const dirHandle = await window.showDirectoryPicker();
        let configJsonStr = null;
        const excelEntries = [];
        const candidateJsonFiles = [];
        const htmlFiles = [];

        for await (const entry of dirHandle.values()) {
          if (entry.kind === 'file') {
            const file = await entry.getFile();
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
              const buffer = await file.arrayBuffer();
              excelEntries.push({ name: file.name, buffer });
            } else if (nameLower.endsWith('.html') || nameLower.endsWith('.htm')) {
              htmlFiles.push(file);
            }
          }
        }

        // Đọc tất cả các file JSON ứng viên và tìm file cấu hình mới nhất
        const validJsonCandidates = [];
        for (const candidate of candidateJsonFiles) {
          try {
            const text = await candidate.file.text();
            const parsed = JSON.parse(text);
            if (parsed && (parsed.items || parsed.tabs || parsed.title)) {
              const modTime = (parsed.updatedAt ? new Date(parsed.updatedAt).getTime() : 0) || (candidate.file.lastModified || 0);
              validJsonCandidates.push({
                text: text,
                modTime: modTime,
                score: candidate.score
              });
            }
          } catch (_) {}
        }

        validJsonCandidates.sort((a, b) => {
          if (b.modTime !== a.modTime) return b.modTime - a.modTime;
          return b.score - a.score;
        });

        if (validJsonCandidates.length > 0) {
          configJsonStr = validJsonCandidates[0].text;
        }

        // Nếu không có file JSON, tìm trong file HTML
        if (!configJsonStr && htmlFiles.length > 0) {
          for (const hf of htmlFiles) {
            try {
              const hText = await hf.text();
              const cfgMarker = 'const CONFIG =';
              const filesMarker = 'let FILES =';
              const cfgIdx = hText.indexOf(cfgMarker);
              const fIdx = hText.indexOf(filesMarker);
              if (cfgIdx !== -1 && fIdx !== -1 && fIdx > cfgIdx) {
                let cStr = hText.substring(cfgIdx + cfgMarker.length, fIdx).trim();
                if (cStr.endsWith(';')) cStr = cStr.slice(0, -1).trim();
                const parsed = JSON.parse(cStr);
                if (parsed && (parsed.items || parsed.tabs || parsed.title)) {
                  configJsonStr = JSON.stringify(parsed);
                  break;
                }
              }
            } catch (_) {}
          }
        }

        if (!configJsonStr) {
          alert('Không tìm thấy file cấu hình Dashboard hợp lệ (*_config.json hoặc index.html) trong thư mục được chọn. Vui lòng chọn thư mục chứa file cấu hình JSON và file Excel dữ liệu!');
          return;
        }

        try {
          localStorage.setItem('mb_dashboard_pending_import_json', configJsonStr);
        } catch (_) {}

        try {
          await storePendingImportInIndexedDB({ configJsonStr, excelEntries, dirHandle, folderName: dirHandle.name });
        } catch (_) {
          await storePendingImportInIndexedDB({ configJsonStr, excelEntries, folderName: dirHandle.name });
        }
        window.location.href = 'dashboard_builder.html?mode=imported';
        return;
      } catch (err) {
        if (err.name === 'AbortError') return;
        console.warn('showDirectoryPicker error in hub:', err);
      }
    }

    // 2. Fallback dùng thẻ input thư mục
    let folderInput = document.getElementById('hubFolderPickerInput');
    if (!folderInput) {
      folderInput = document.createElement('input');
      folderInput.type = 'file';
      folderInput.id = 'hubFolderPickerInput';
      folderInput.webkitdirectory = true;
      folderInput.directory = true;
      folderInput.multiple = true;
      folderInput.style.display = 'none';
      document.body.appendChild(folderInput);

      folderInput.addEventListener('change', async (ev) => {
        const files = Array.from(ev.target.files || []);
        if (files.length === 0) return;
        let configJsonStr = null;
        const excelEntries = [];
        const candidateJsonFiles = [];
        const htmlFiles = [];

        for (const file of files) {
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
            const buffer = await file.arrayBuffer();
            excelEntries.push({ name: file.name, buffer });
          } else if (nameLower.endsWith('.html') || nameLower.endsWith('.htm')) {
            htmlFiles.push(file);
          }
        }

        const validJsonCandidates = [];
        for (const candidate of candidateJsonFiles) {
          try {
            const text = await candidate.file.text();
            const parsed = JSON.parse(text);
            if (parsed && (parsed.items || parsed.tabs || parsed.title)) {
              const modTime = (parsed.updatedAt ? new Date(parsed.updatedAt).getTime() : 0) || (candidate.file.lastModified || 0);
              validJsonCandidates.push({
                text: text,
                modTime: modTime,
                score: candidate.score
              });
            }
          } catch (_) {}
        }

        validJsonCandidates.sort((a, b) => {
          if (b.modTime !== a.modTime) return b.modTime - a.modTime;
          return b.score - a.score;
        });

        if (validJsonCandidates.length > 0) {
          configJsonStr = validJsonCandidates[0].text;
        }

        if (!configJsonStr && htmlFiles.length > 0) {
          for (const hf of htmlFiles) {
            try {
              const hText = await hf.text();
              const cfgMarker = 'const CONFIG =';
              const filesMarker = 'let FILES =';
              const cfgIdx = hText.indexOf(cfgMarker);
              const fIdx = hText.indexOf(filesMarker);
              if (cfgIdx !== -1 && fIdx !== -1 && fIdx > cfgIdx) {
                let cStr = hText.substring(cfgIdx + cfgMarker.length, fIdx).trim();
                if (cStr.endsWith(';')) cStr = cStr.slice(0, -1).trim();
                const parsed = JSON.parse(cStr);
                if (parsed && (parsed.items || parsed.tabs || parsed.title)) {
                  configJsonStr = JSON.stringify(parsed);
                  break;
                }
              }
            } catch (_) {}
          }
        }

        if (!configJsonStr) {
          alert('Không tìm thấy file cấu hình Dashboard hợp lệ (*_config.json hoặc index.html) trong thư mục được chọn. Vui lòng chọn thư mục chứa file cấu hình JSON và file Excel dữ liệu!');
          return;
        }

        try {
          localStorage.setItem('mb_dashboard_pending_import_json', configJsonStr);
        } catch (_) {}

        await storePendingImportInIndexedDB({ configJsonStr, excelEntries });
        window.location.href = 'dashboard_builder.html?mode=imported';
      });
    }
    folderInput.value = '';
    folderInput.click();
  }

  // =========================================================================
  // 14. INITIALIZATION
  // =========================================================================
  function init() {
    if (elements.sidebar) initSidebar();
    setupEventListeners();
    if (elements.sidebarPinnedList) renderSidebarPinnedTools();
    renderCategoryPills();
    renderToolsGrid();
  }

  // Run on DOM ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

})();
