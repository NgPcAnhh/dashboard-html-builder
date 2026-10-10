const fs = require('fs');

const code = fs.readFileSync('js/dashboard-builder.js', 'utf8');
const startMarker = 'function generateStandaloneDashboardHTML() {';
const startIndex = code.indexOf(startMarker);
const endMarker = 'function prepareDashboardConfigJson() {';
const endIndex = code.indexOf(endMarker, startIndex);
const moduleEscapeIndex = code.lastIndexOf('function escapeHtml(str) {', endIndex);

const funcCode = code.substring(startIndex, moduleEscapeIndex).trim();

const config = JSON.parse(fs.readFileSync('storage/Báo_Cáo_Dashboard_Mới/dashboard_config.json', 'utf8'));

const mockScope = {
  DashboardState: {
    present: config,
    activeFilters: {}
  },
  DataStore: {
    files: config.embeddedData || {},
    getAllFilesObject: () => config.embeddedData || {}
  },
  TabBarManager: {
    getConfig: () => config.tabBarConfig || {}
  },
  document: {
    getElementById: (id) => {
      if (id === 'canvasWrapper') return { className: 'canvas-bg-dots' };
      if (id === 'dashboardTitleInput') return { value: config.title };
      return null;
    }
  },
  prepareDashboardConfigJson: () => {},
  escapeHtml: (str) => {
    if (!str) return '';
    return String(str).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
  }
};

const evalFn = new Function('DashboardState', 'DataStore', 'TabBarManager', 'document', 'prepareDashboardConfigJson', 'escapeHtml', 'return (' + funcCode + ')()');
const generatedHtml = evalFn(mockScope.DashboardState, mockScope.DataStore, mockScope.TabBarManager, mockScope.document, mockScope.prepareDashboardConfigJson, mockScope.escapeHtml);

fs.writeFileSync('storage/Báo_Cáo_Dashboard_Mới/index.html', generatedHtml, 'utf8');
console.log('Successfully generated storage/Báo_Cáo_Dashboard_Mới/index.html! Length:', generatedHtml.length);
