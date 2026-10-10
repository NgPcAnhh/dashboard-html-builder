/**
 * terminal.js - Quản lý Terminal Console & Điều hướng Console Log
 */

const terminal = document.getElementById('terminal');

function printToTerminal(msg, type = '') {
  if (!terminal) return;
  const line = document.createElement('div');
  if (type) line.className = type;
  line.textContent = (typeof msg === 'object') ? JSON.stringify(msg, null, 2) : msg;
  terminal.appendChild(line);
  terminal.scrollTop = terminal.scrollHeight;
}

// Chuyển hướng console.log, warn, error ra Terminal UI
const nativeLog = console.log;
const nativeErr = console.error;
const nativeWarn = console.warn;

console.log = (...args) => {
  nativeLog(...args);
  args.forEach(a => printToTerminal(a));
};

console.error = (...args) => {
  nativeErr(...args);
  args.forEach(a => printToTerminal(a, 'log-err'));
};

console.warn = (...args) => {
  nativeWarn(...args);
  args.forEach(a => printToTerminal(a, 'log-warn'));
};

const btnClearTerm = document.getElementById('btnClearTerm');
if (btnClearTerm) {
  btnClearTerm.onclick = () => {
    if (terminal) terminal.innerHTML = '> Đã xóa màn hình console.';
  };
}
