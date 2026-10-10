const { spawn } = require('child_process');
const http = require('http');

async function testHeadless() {
  const edgePath = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';
  const targetUrl = 'file:///C:/Users/Admin/Desktop/small%20tools/data_process_tools/storage/Báo_Cáo_Dashboard_Mới/index.html';
  
  const edge = spawn(edgePath, [
    '--headless',
    '--remote-debugging-port=9222',
    '--disable-gpu',
    targetUrl
  ]);

  await new Promise(r => setTimeout(r, 1500));

  // Query pages
  http.get('http://127.0.0.1:9222/json', (res) => {
    let data = '';
    res.on('data', chunk => data += chunk);
    res.on('end', () => {
      try {
        const pages = JSON.parse(data);
        console.log('Pages found:', pages.length);
        const page = pages.find(p => p.type === 'page');
        if (page && page.webSocketDebuggerUrl) {
          const ws = new WebSocket(page.webSocketDebuggerUrl);
          ws.onopen = () => {
            console.log('Connected to CDP');
            ws.send(JSON.stringify({ id: 1, method: 'Runtime.enable' }));
            ws.send(JSON.stringify({ id: 2, method: 'Console.enable' }));
            ws.send(JSON.stringify({ id: 3, method: 'Runtime.evaluate', params: {
              expression: `(new Promise(resolve => {
                const tabs = document.querySelectorAll('.db-tab-pill');
                if (tabs[1]) tabs[1].click();
                setTimeout(() => {
                  resolve({
                    activeTabAfterClick: document.querySelector('.db-tab-pill.active')?.textContent.trim(),
                    canvasesNow: document.querySelectorAll('canvas').length,
                    chartsNow: document.querySelectorAll('.db-chart-canvas').length
                  });
                }, 150);
              }))`,
              awaitPromise: true,
              returnByValue: true
            }}));


          };
          ws.onmessage = (event) => {
            const msg = JSON.parse(event.data);
            if (msg.method === 'Runtime.consoleAPICalled') {
              console.log('Console', msg.params.type, ':', msg.params.args.map(a => a.value || a.description));
            }
            if (msg.method === 'Runtime.exceptionThrown') {
              console.error('EXCEPTION:', msg.params.exceptionDetails.text, msg.params.exceptionDetails.exception?.description);
            }
            if (msg.id === 3) {
              console.log('EVAL RESULT:', msg.result.result.value);
              setTimeout(() => {
                ws.close();
                edge.kill();
                process.exit(0);
              }, 500);
            }
          };
        } else {
          edge.kill();
        }
      } catch (e) {
        console.error('Parse error:', e);
        edge.kill();
      }
    });
  }).on('error', (e) => {
    console.error('HTTP error:', e);
    edge.kill();
  });
}

testHeadless();
