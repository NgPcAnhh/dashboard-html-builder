const fs = require('fs');
const html = fs.readFileSync('scratch/dumped_dom.html', 'utf8');

const canvasMatches = html.match(/<canvas/g) || [];
console.log('Total <canvas> elements rendered:', canvasMatches.length);

const chartBoxes = html.match(/db-item-chart/g) || [];
console.log('Total chart item boxes:', chartBoxes.length);

const kpiBoxes = html.match(/db-item-kpi/g) || [];
console.log('Total KPI item boxes:', kpiBoxes.length);

const filterBoxes = html.match(/db-item-filter/g) || [];
console.log('Total filter boxes:', filterBoxes.length);

const errors = html.match(/db-chart-error-banner/g) || [];
console.log('Chart error banners:', errors.length);

const errMatches = html.match(/<div class="db-chart-error-banner"[^>]*>([\s\S]*?)<\/div>/g) || [];
errMatches.forEach((e, idx) => console.log('Error banner ' + idx + ':', e));

const tabs = html.match(/db-tab-pill/g) || [];
console.log('Total tab pills:', tabs.length);

