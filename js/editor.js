/**
 * editor.js - Quản lý Trình soạn thảo Mã nguồn, Cú pháp Đổi màu & Indent Guides
 * Hỗ trợ:
 * - Syntax Highlighting toàn diện (VS Code Dark+ & Light)
 * - Đường dóng thụt lề (Indentation Guides) & Dải màu thụt dòng (Indent Rainbow)
 * - Tự động thụt lề (Auto Indent) khi nhấn Enter, ngoặc block { }
 * - Hỗ trợ Tab / Shift + Tab thụt dòng nhiều dòng
 * - Tự động đóng cặp ngoặc { }, [ ], ( ), " ", ' ', ` `
 * - Nút & Phím tắt Format Code (Shift + Alt + F)
 * - Quản lý nhiều tab file script JS
 */

const codeEditor = document.getElementById('codeEditor');
const codeHighlight = document.getElementById('codeHighlight');
const lineNumbers = document.getElementById('lineNumbers');
const tabsBar = document.getElementById('tabsBar');
const btnAddTab = document.getElementById('btnAddTab');
const btnFormatCode = document.getElementById('btnFormatCode');

// ========================================================
// 1. ENGINE SYNTAX HIGHLIGHTING & INDENT GUIDES (ZERO DEPENDENCY)
// ========================================================
function escapeHtml(str) {
  return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/**
 * Tạo thẻ HTML cho đường dóng và dải màu thụt dòng (Indent Rainbow Guides)
 * Mỗi 2 khoảng trắng (hoặc 1 tab) tương ứng 1 cấp độ thụt dòng
 */
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
      // Khoảng trắng lẻ đơn lẻ (1 space)
      html += `<span class="indent-guide ${lvlClass}">${ch}</span>`;
      i++;
    }
  }

  return html;
}

// Regex nhận diện các token JavaScript chuẩn xác
const JS_TOKEN_REGEX = /(\/\*[\s\S]*?\*\/|\/\/[^\n]*)|("(?:\\.|[^"\n\\])*"|'(?:\\.|[^'\n\\])*'|`(?:\\.|[^`\\])*`|`(?:\\.|[^`\\])*$)|(\b(?:const|let|var|function|return|if|else|for|while|do|switch|case|default|break|continue|try|catch|finally|throw|new|typeof|instanceof|class|extends|import|export|from|async|await|yield|in|of|void|delete|debugger)\b)|(\b(?:true|false|null|undefined|NaN|Infinity)\b)|(\b(?:XLSX|exportExcel|rawRows|files|fileSheets|getFile|getSheet|mergeAllRows|context|dataColumns|sampleRecords|console|window|document|Array|Object|String|Number|Boolean|Math|JSON|Date|RegExp|Map|Set|Promise)\b)|(\b\d+(?:\.\d+)?(?:e[+-]?\d+)?\b)|(\b[a-zA-Z_$][a-zA-Z0-9_$]*(?=\s*\())|(\b[a-zA-Z_$][a-zA-Z0-9_$]*(?=\s*:))|([{}()\[\]])|(===|!==|==|!=|<=|>=|=>|&&|\|\||\?\?|\+=|-=|\*=|\/=|[\+\-\*\/%=\?:\!&\|\^~<>])/g;

/**
 * Tô màu cú pháp một dòng code
 */
function highlightCodeLine(codePart, state) {
  // Xử lý comment nhiều dòng tiếp diễn từ dòng trước
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

  // Xử lý template string nhiều dòng tiếp diễn từ dòng trước
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
      if (comment.startsWith('/*') && !comment.endsWith('*/')) {
        state.inBlockComment = true;
      }
      html += `<span class="token-comment">${escapeHtml(comment)}</span>`;
    } else if (str) {
      if (str.startsWith('`') && (str.length === 1 || !str.endsWith('`') || str.endsWith('\\`'))) {
        state.inTemplateStr = true;
      }
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

/**
 * Hàm phân tích toàn bộ code, chèn Indent Guides và tô màu theo dòng
 */
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

/**
 * Tự động căn chỉnh và thụt lề code chuẩn 2 khoảng trắng (Format / Beautify JS)
 */
function formatJSCode(code) {
  if (!code || typeof code !== 'string') return code;

  const lines = code.split('\n');
  let indentLevel = 0;
  const indentStr = '  '; // 2 spaces chuẩn
  const resultLines = [];
  let inMultiComment = false;

  for (let i = 0; i < lines.length; i++) {
    let line = lines[i].trim();

    // Giữ dòng trống tối đa 1 dòng liên tiếp
    if (!line) {
      if (resultLines.length > 0 && resultLines[resultLines.length - 1] !== '') {
        resultLines.push('');
      }
      continue;
    }

    // Comment khối /* ... */
    if (line.startsWith('/*')) inMultiComment = true;
    if (inMultiComment) {
      resultLines.push(indentStr.repeat(indentLevel) + line);
      if (line.includes('*/')) inMultiComment = false;
      continue;
    }

    // Đếm số ngoặc đóng đầu dòng để lùi lề ngay dòng này
    let leadingClosers = 0;
    for (let c of line) {
      if (c === '}' || c === ']' || c === ')') {
        leadingClosers++;
      } else if (c === ' ' || c === '\t') {
        continue;
      } else {
        break;
      }
    }

    let currentIndent = Math.max(0, indentLevel - leadingClosers);

    // Tính thay đổi indent cho các dòng sau
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
      if (ch === '/' && line[j + 1] === '/') {
        break;
      }
      if (ch === '{' || ch === '[' || ch === '(') {
        delta++;
      } else if (ch === '}' || ch === ']' || ch === ')') {
        delta--;
      }
    }

    resultLines.push(indentStr.repeat(currentIndent) + line);
    indentLevel = Math.max(0, indentLevel + delta);
  }

  return resultLines.join('\n');
}

function formatCurrentCode() {
  if (!codeEditor) return;
  const currentCode = codeEditor.value;
  const formatted = formatJSCode(currentCode);
  if (formatted !== currentCode) {
    codeEditor.value = formatted;
    updateEditor();
  }
}

/**
 * Cập nhật syntax highlighting và số dòng
 */
function updateEditor() {
  if (!codeEditor || !codeHighlight || !lineNumbers) return;
  const code = codeEditor.value;
  codeHighlight.innerHTML = highlightJS(code) + (code.endsWith('\n') ? ' ' : '');

  // Cập nhật số dòng
  const linesCount = code.split("\n").length;
  const linesArr = [];
  for (let i = 1; i <= linesCount; i++) {
    linesArr.push(i);
  }
  lineNumbers.innerText = linesArr.join("\n");

  // Lưu lại vào script hiện tại trong bộ nhớ
  const current = scripts.find(s => s.id === activeScriptId);
  if (current) {
    current.code = code;
  }
}

// ========================================================
// 2. TƯƠNG TÁC BÀN PHÍM THÔNG MINH TRONG EDITOR
// ========================================================
if (codeEditor) {
  // Đồng bộ cuộn giữa Line Numbers, Syntax Highlight và Textarea
  codeEditor.addEventListener('scroll', () => {
    if (codeHighlight && codeHighlight.parentElement) {
      codeHighlight.parentElement.scrollTop = codeEditor.scrollTop;
      codeHighlight.parentElement.scrollLeft = codeEditor.scrollLeft;
    }
    if (lineNumbers) {
      lineNumbers.scrollTop = codeEditor.scrollTop;
    }
  });

  codeEditor.addEventListener('input', updateEditor);

  codeEditor.addEventListener('keydown', function (e) {
    const start = this.selectionStart;
    const end = this.selectionEnd;

    // 1. Phím tắt Format Code (Shift + Alt + F hoặc Ctrl + Shift + F)
    if ((e.shiftKey && e.altKey && (e.key === 'F' || e.key === 'f')) ||
        (e.ctrlKey && e.shiftKey && (e.key === 'F' || e.key === 'f'))) {
      e.preventDefault();
      formatCurrentCode();
      return;
    }

    // 2. Xử lý Shift + Tab (Lùi lề / Unindent)
    if (e.key === 'Tab' && e.shiftKey) {
      e.preventDefault();
      const startLineIndex = this.value.lastIndexOf('\n', start - 1) + 1;
      let endLineIndex = this.value.indexOf('\n', end);
      if (endLineIndex === -1) endLineIndex = this.value.length;

      const lines = this.value.substring(startLineIndex, endLineIndex).split('\n');
      let unindentedFirstLine = 0;
      let totalRemoved = 0;

      const unindented = lines.map((line, idx) => {
        let removeCount = 0;
        if (line.startsWith('  ')) removeCount = 2;
        else if (line.startsWith(' ') || line.startsWith('\t')) removeCount = 1;
        if (idx === 0) unindentedFirstLine = removeCount;
        totalRemoved += removeCount;
        return line.substring(removeCount);
      }).join('\n');

      this.value = this.value.substring(0, startLineIndex) + unindented + this.value.substring(endLineIndex);
      this.selectionStart = Math.max(startLineIndex, start - unindentedFirstLine);
      this.selectionEnd = Math.max(startLineIndex, end - totalRemoved);
      updateEditor();
      return;
    }

    // 3. Xử lý Tab (Thụt lề 2 spaces / Indent đa dòng)
    if (e.key === 'Tab' && !e.shiftKey) {
      e.preventDefault();
      if (start !== end && this.value.substring(start, end).includes('\n')) {
        const startLineIndex = this.value.lastIndexOf('\n', start - 1) + 1;
        let endLineIndex = this.value.indexOf('\n', end);
        if (endLineIndex === -1) endLineIndex = this.value.length;

        const lines = this.value.substring(startLineIndex, endLineIndex).split('\n');
        const indented = lines.map(line => '  ' + line).join('\n');
        this.value = this.value.substring(0, startLineIndex) + indented + this.value.substring(endLineIndex);
        this.selectionStart = start + 2;
        this.selectionEnd = startLineIndex + indented.length;
        updateEditor();
      } else {
        this.value = this.value.substring(0, start) + "  " + this.value.substring(end);
        this.selectionStart = this.selectionEnd = start + 2;
        updateEditor();
      }
      return;
    }

    // 4. Xử lý Enter (Tự động thụt lề chuẩn dòng trước - Auto Indent)
    if (e.key === 'Enter') {
      const lineStart = this.value.lastIndexOf('\n', start - 1) + 1;
      const currentLineUntilCursor = this.value.substring(lineStart, start);
      const indentMatch = currentLineUntilCursor.match(/^[ \t]+/);
      let indent = indentMatch ? indentMatch[0] : '';
      const trimmed = currentLineUntilCursor.trim();
      const nextChar = this.value.charAt(start);

      // Nếu dòng trước kết thúc bằng { hoặc [ hoặc (
      if (trimmed.endsWith('{') || trimmed.endsWith('[') || trimmed.endsWith('(')) {
        // Tự động mở rộng khối nếu ký tự tiếp theo là ngoặc đóng tương ứng
        if ((trimmed.endsWith('{') && nextChar === '}') ||
            (trimmed.endsWith('[') && nextChar === ']') ||
            (trimmed.endsWith('(') && nextChar === ')')) {
          e.preventDefault();
          const insertText = '\n' + indent + '  \n' + indent;
          this.value = this.value.substring(0, start) + insertText + this.value.substring(start);
          this.selectionStart = this.selectionEnd = start + indent.length + 3;
          updateEditor();
          return;
        } else {
          indent += '  ';
        }
      }

      e.preventDefault();
      this.value = this.value.substring(0, start) + '\n' + indent + this.value.substring(start);
      this.selectionStart = this.selectionEnd = start + 1 + indent.length;
      updateEditor();
      return;
    }

    // 5. Tự động đóng cặp ngoặc & dấu nháy { }, [ ], ( ), " ", ' ', ` `
    const autoPairs = { '{': '}', '(': ')', '[': ']', '"': '"', "'": "'", '`': '`' };
    if (autoPairs[e.key]) {
      const closer = autoPairs[e.key];
      const nextChar = this.value.charAt(start);

      // Nếu người dùng gõ ký tự đóng và con trỏ đang nằm ngay trước nó
      if (nextChar === e.key && ["'", '"', '`', '}', ')', ']'].includes(e.key)) {
        e.preventDefault();
        this.selectionStart = this.selectionEnd = start + 1;
        return;
      }

      // Đóng cặp nếu đang bôi đen text
      if (start !== end) {
        e.preventDefault();
        const selected = this.value.substring(start, end);
        this.value = this.value.substring(0, start) + e.key + selected + closer + this.value.substring(end);
        this.selectionStart = start + 1;
        this.selectionEnd = end + 1;
        updateEditor();
        return;
      } else if (!nextChar || /[\s\)\}\]\;\,\.]/.test(nextChar)) {
        e.preventDefault();
        this.value = this.value.substring(0, start) + e.key + closer + this.value.substring(start);
        this.selectionStart = this.selectionEnd = start + 1;
        updateEditor();
        return;
      }
    }
  });
}

if (btnFormatCode) {
  btnFormatCode.addEventListener('click', () => {
    formatCurrentCode();
  });
}

// ========================================================
// 3. QUẢN LÝ NHIỀU TAB SCRIPT JS
// ========================================================
function renderTabs() {
  if (!tabsBar) return;
  tabsBar.innerHTML = '';
  scripts.forEach(script => {
    const tab = document.createElement('div');
    tab.className = `tab-item ${script.id === activeScriptId ? 'active' : ''}`;
    tab.title = `Nhấp đúp để đổi tên file (${script.name})`;

    const nameSpan = document.createElement('span');
    nameSpan.className = 'tab-name';
    nameSpan.textContent = script.name;

    // Đổi tên khi nhấp đúp vào tab
    tab.ondblclick = (e) => {
      e.stopPropagation();
      const newName = prompt("Nhập tên mới cho file script:", script.name);
      if (newName && newName.trim()) {
        script.name = newName.trim().endsWith('.js') ? newName.trim() : newName.trim() + '.js';
        renderTabs();
      }
    };

    const closeBtn = document.createElement('span');
    closeBtn.className = 'tab-close';
    closeBtn.innerHTML = '&times;';
    closeBtn.title = 'Đóng tab này';
    closeBtn.onclick = (e) => {
      e.stopPropagation();
      closeTab(script.id);
    };

    tab.appendChild(nameSpan);
    if (scripts.length > 1) {
      tab.appendChild(closeBtn);
    }

    tab.onclick = () => {
      switchTab(script.id);
    };

    tabsBar.appendChild(tab);
  });
}

function switchTab(scriptId) {
  activeScriptId = scriptId;
  const current = scripts.find(s => s.id === activeScriptId);
  if (current && codeEditor) {
    codeEditor.value = current.code;
    updateEditor();
    codeEditor.scrollTop = 0;
    codeEditor.selectionStart = 0;
    codeEditor.selectionEnd = 0;
  }
  renderTabs();
}

function addTab() {
  const newId = Date.now();
  const scriptCount = scripts.length + 1;
  const newScript = {
    id: newId,
    name: `step_${scriptCount}_process.js`,
    code: `// Script xử lý tiếp theo
console.log("Đang chạy: step_${scriptCount}_process.js");
console.log("Số dòng hiện tại trong rawRows:", rawRows.length);
`
  };
  scripts.push(newScript);
  renderTabs();
  switchTab(newId);
}

function closeTab(scriptId) {
  if (scripts.length <= 1) return;
  const idx = scripts.findIndex(s => s.id === scriptId);
  if (idx !== -1) {
    scripts.splice(idx, 1);
    if (activeScriptId === scriptId) {
      activeScriptId = scripts[Math.max(0, idx - 1)].id;
    }
    renderTabs();
    switchTab(activeScriptId);
  }
}

if (btnAddTab) {
  btnAddTab.onclick = addTab;
}
