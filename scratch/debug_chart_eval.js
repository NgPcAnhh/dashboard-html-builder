const fs = require('fs');
const html = fs.readFileSync('storage/Báo_Cáo_Dashboard_Mới/index.html', 'utf8');

// Find the script tag containing CONFIG and evaluateEChartsCode
const scriptMatch = html.match(/<script>([\s\S]*?)<\/script>[\s\r\n]*<\/body>/i);
if (!scriptMatch) {
  console.log('No script match found');
  process.exit(1);
}

const scriptContent = scriptMatch[1];
// Let's extract CONFIG
const configMatch = scriptContent.match(/const CONFIG = ({[\s\S]*?});[\s\r\n]*let FILES = ({[\s\S]*?});/);
if (!configMatch) {
  console.log('No CONFIG match found');
  process.exit(1);
}

const config = JSON.parse(configMatch[1]);
console.log('CONFIG loaded, items count:', config.items.length);

const files = JSON.parse(configMatch[2]);
console.log('FILES loaded, files count:', Object.keys(files).length);


// Let's run evaluateEChartsCode exactly as written in index.html
function evaluateEChartsCode(codeStr, data) {
  if (!codeStr || !codeStr.trim()) return { error: 'Không có mã cấu hình ECharts.' };
  try {
    const rawFiles = files;
    const getFile = (fn) => files[fn] || null;
    const getRawSheet = (fn, sn) => (files[fn] && files[fn][sn]) ? files[fn][sn] : [];
    // Mock minimal echarts graphic LinearGradient
    const echarts = {
      graphic: {
        LinearGradient: function(x, y, x2, y2, colorStops) {
          return { type: 'linear', x, y, x2, y2, colorStops };
        }
      }
    };
    const fn = new Function('data', 'rawFiles', 'getFile', 'getRawSheet', 'echarts', 'console',
      'try { ' + codeStr + ' } catch (e) { return { __error: e.stack || e.message }; }'
    );
    const result = fn(data, rawFiles, getFile, getRawSheet, echarts, console);
    if (result && result.__error) return { error: result.__error };
    return { option: result };
  } catch (err) {
    return { error: err.message };
  }
}

config.items.forEach((item, idx) => {
  if (item.type === 'chart') {
    let rawData = [];
    if (item.dataSourceKey) {
      const parts = item.dataSourceKey.split('::');
      const fName = parts[0];
      const sName = parts[1];
      if (files[fName] && files[fName][sName]) {
        rawData = files[fName][sName];
      }
    }
    const res = evaluateEChartsCode(item.code, rawData);
    if (res.error) {
      console.log('Item ' + idx + ' (' + item.id + ' - ' + item.name + ') ERROR:\n' + res.error);
    } else {
      console.log('Item ' + idx + ' (' + item.id + ' - ' + item.name + ') OK');
    }
  }
});
