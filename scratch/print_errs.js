const fs = require('fs');
const html = fs.readFileSync('scratch/dumped_dom.html', 'utf8');
const regex = /<div class="db-chart-error-banner"[^>]*>([\s\S]*?)<\/div>/g;
let m;
let i = 0;
while ((m = regex.exec(html)) !== null) {
  console.log('--- ERROR ' + (i++) + ' ---');
  console.log(m[1]);
}
