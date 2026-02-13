/* global Office, Excel */

// Auto-detect dev vs production environment
const API_BASE = (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1')
  ? `https://${window.location.host}`
  : 'https://cellobot-production.up.railway.app';

// Conversation history for multi-turn context
let conversationHistory = [];

// Pinned context ranges (persists per chat session)
// Each entry: { id, sheet, range, label, mode: 'full'|'sample', sampleRows: 50, rowCount, columnCount }
let pinnedRanges = [];
const DEFAULT_SAMPLE_ROWS = 50;

function getContextWindowSize(modelValue) {
  const val = (modelValue || 'claude-sonnet').toLowerCase();
  switch (val) {
    case 'claude-opus':
    case 'claude-sonnet':
    case 'claude':
      return 200000;
    case 'openai-4o':
    case 'openai-4o-mini':
    case 'openai':
      return 128000;
    default:
      return 200000;
  }
}

function estimatePinnedTokensSync() {
  let total = 0;
  for (const p of pinnedRanges) {
    const rows = p.mode === 'sample' ? Math.min(p.rowCount || 0, p.sampleRows || DEFAULT_SAMPLE_ROWS) : (p.rowCount || 0);
    const cols = p.columnCount || 0;
    total += Math.ceil((rows * cols * 5) / 4);
  }
  return total;
}

async function addPinnedRange(sheet, range, label, mode = 'full', rowCount, columnCount) {
  const key = `${sheet}!${range}`;
  if (pinnedRanges.some(p => `${p.sheet}!${p.range}` === key)) return;

  let rows = rowCount;
  let cols = columnCount;
  if (rows == null || cols == null) {
    try {
      const dims = await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItem(sheet);
        const r = ws.getRange(range);
        r.load('rowCount,columnCount');
        await ctx.sync();
        return { rowCount: r.rowCount, columnCount: r.columnCount };
      });
      rows = dims.rowCount;
      cols = dims.columnCount;
    } catch (e) {
      rows = rows || 0;
      cols = cols || 0;
    }
  }

  const entry = {
    id: 'pinned-' + Date.now() + '-' + Math.random().toString(36).substring(7),
    sheet,
    range,
    label: label || `${sheet}!${range}`,
    mode: mode || 'full',
    sampleRows: DEFAULT_SAMPLE_ROWS,
    rowCount: rows,
    columnCount: cols
  };
  pinnedRanges.push(entry);
  updateContextUI();
}

function removePinnedRange(id) {
  pinnedRanges = pinnedRanges.filter(p => p.id !== id);
  updateContextUI();
}

function clearAllPinned() {
  pinnedRanges = [];
  updateContextUI();
}

function togglePinnedRangeMode(id) {
  const p = pinnedRanges.find(x => x.id === id);
  if (p) {
    p.mode = p.mode === 'full' ? 'sample' : 'full';
    updateContextUI();
  }
}

async function getPinnedContext() {
  if (pinnedRanges.length === 0) return [];

  return Excel.run(async (ctx) => {
    const results = [];
    for (const p of pinnedRanges) {
      try {
        const sheet = ctx.workbook.worksheets.getItem(p.sheet);
        const range = sheet.getRange(p.range);
        range.load('values,formulas,rowCount,columnCount');
        await ctx.sync();

        let values = range.values || [];
        let formulas = range.formulas || [];

        if (p.mode === 'sample' && range.rowCount > (p.sampleRows || DEFAULT_SAMPLE_ROWS)) {
          const limit = p.sampleRows || DEFAULT_SAMPLE_ROWS;
          values = values.slice(0, limit);
          formulas = formulas.slice(0, limit);
        }

        const headers = values.length > 0 ? values[0].map(v => (v != null ? String(v).trim() : '')) : [];
        results.push({
          sheet: p.sheet,
          range: p.range,
          label: p.label,
          headers,
          values,
          formulas
        });
      } catch (e) {
        results.push({ sheet: p.sheet, range: p.range, error: e.message });
      }
    }
    return results;
  });
}

function updateContextUI() {
  const badge = document.getElementById('context-badge');
  if (badge) {
    badge.textContent = pinnedRanges.length;
    badge.classList.toggle('hidden', pinnedRanges.length === 0);
  }
  renderPinnedChips();
  updateTokenMeter();
}

function toggleContextPanel() {
  const panel = document.getElementById('context-panel');
  if (!panel) return;
  const isHidden = panel.classList.toggle('hidden');
  if (!isHidden) loadSheetList();
}

function closeContextPanel() {
  const panel = document.getElementById('context-panel');
  if (panel) panel.classList.add('hidden');
}

async function loadSheetList() {
  const container = document.getElementById('sheet-list');
  if (!container) return;

  try {
    const sheetData = await Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      sheets.load('items/name');
      await ctx.sync();

      const result = [];
      for (const ws of sheets.items) {
        let usedRangeAddr = null;
        let rows = 0;
        let cols = 0;
        let tables = [];

        try {
          const usedRange = ws.getUsedRange();
          usedRange.load('address,rowCount,columnCount');
          await ctx.sync();
          usedRangeAddr = usedRange.address.replace(/^.*!/, '');
          rows = usedRange.rowCount;
          cols = usedRange.columnCount;
        } catch (e) {
          /* empty sheet */
        }

        try {
          const tbls = ws.tables;
          tbls.load('items/name,items/range/address');
          await ctx.sync();
          for (const t of tbls.items) {
            const addr = t.range.address.replace(/^.*!/, '');
            tables.push({ name: t.name, range: addr });
          }
        } catch (e) {
          /* no tables */
        }

        result.push({
          name: ws.name,
          usedRange: usedRangeAddr,
          rows,
          cols,
          tables
        });
      }
      return result;
    });

    container.innerHTML = '';
    for (const s of sheetData) {
      const row = document.createElement('div');
      row.className = 'sheet-row';
      const summary = s.usedRange ? `${s.usedRange} — ${s.rows} rows × ${s.cols} cols` : 'Empty';
      row.innerHTML = `
        <div class="sheet-row-header" data-sheet="${escapeHtml(s.name)}">
          <span class="sheet-row-name">${escapeHtml(s.name)}</span>
          <span class="sheet-row-summary">${escapeHtml(summary)}</span>
        </div>
        <div class="sheet-row-body hidden">
          ${s.tables.map(t => `<button type="button" class="sheet-table-chip" data-sheet="${escapeHtml(s.name)}" data-range="${escapeHtml(t.range)}" data-label="${escapeHtml(t.name)}">${escapeHtml(t.name)} (${escapeHtml(t.range)})</button>`).join('')}
          <div class="sheet-custom-range">
            <input type="text" placeholder="e.g. A1:D50" data-sheet="${escapeHtml(s.name)}" />
            <button type="button" class="sheet-custom-add" data-sheet="${escapeHtml(s.name)}">Add</button>
          </div>
        </div>
      `;
      container.appendChild(row);
    }
  } catch (err) {
    console.warn('Could not load sheet list:', err);
    container.innerHTML = '<p style="color: var(--text-muted); font-size: 13px;">Could not load sheets. Make sure Excel is ready.</p>';
  }
}

function renderPinnedChips() {
  const container = document.getElementById('pinned-chips');
  if (!container) return;

  container.innerHTML = '';
  for (const p of pinnedRanges) {
    const chip = document.createElement('div');
    chip.className = 'pinned-chip';
    const modeLabel = p.mode === 'sample' ? `Sample (${p.sampleRows})` : 'Full';
    chip.innerHTML = `
      <span class="pinned-chip-label" data-id="${p.id}">${escapeHtml(p.label)} (${modeLabel})</span>
      <button type="button" class="pinned-chip-remove" data-id="${p.id}" aria-label="Remove">×</button>
    `;
    container.appendChild(chip);
  }
}

function updateTokenMeter() {
  const fill = document.getElementById('token-meter-fill');
  const label = document.getElementById('token-meter-label');
  const modelSelect = document.getElementById('model-select');
  if (!fill || !label) return;

  const tokens = estimatePinnedTokensSync();
  const maxTokens = getContextWindowSize(modelSelect ? modelSelect.value : null);
  const pct = maxTokens > 0 ? (tokens / maxTokens) * 100 : 0;

  fill.style.width = Math.min(pct, 100) + '%';
  fill.classList.remove('safe', 'warning', 'danger');
  if (pct < 50) fill.classList.add('safe');
  else if (pct < 80) fill.classList.add('warning');
  else fill.classList.add('danger');

  const maxK = Math.round(maxTokens / 1000);
  label.textContent = `~${tokens.toLocaleString()} tokens (${pct.toFixed(1)}% of ${maxK}k)`;
}

async function captureCurrentSelection() {
  try {
    const result = await Excel.run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      range.load('address,rowCount,columnCount');
      sheet.load('name');
      await ctx.sync();
      return {
        sheet: sheet.name,
        range: range.address.replace(/^.*!/, ''),
        rowCount: range.rowCount,
        columnCount: range.columnCount
      };
    });
    await addPinnedRange(result.sheet, result.range, `${result.sheet}!${result.range}`, 'full', result.rowCount, result.columnCount);
  } catch (err) {
    console.warn('Could not capture selection:', err);
    addMessage('Could not capture selection. Make sure a range is selected.', 'assistant', { className: 'tool-status' });
  }
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById('send-btn').addEventListener('click', onSend);
    document.getElementById('new-chat-btn').addEventListener('click', onNewChat);
    document.getElementById('undo-all-btn').addEventListener('click', undoAllChanges);
    document.getElementById('export-chat-btn').addEventListener('click', exportChatHistory);
    document.getElementById('context-btn').addEventListener('click', toggleContextPanel);
    document.getElementById('context-panel-close').addEventListener('click', closeContextPanel);
    document.getElementById('capture-selection-btn').addEventListener('click', captureCurrentSelection);

    const modelSelect = document.getElementById('model-select');
    if (modelSelect) modelSelect.addEventListener('change', updateTokenMeter);

    // Event delegation for sheet list (expand, add table, add custom range)
    const sheetList = document.getElementById('sheet-list');
    if (sheetList) {
      sheetList.addEventListener('click', (e) => {
        if (e.target.classList.contains('sheet-row-header')) {
          const body = e.target.nextElementSibling;
          if (body) body.classList.toggle('hidden');
        } else if (e.target.classList.contains('sheet-table-chip')) {
          const sheet = e.target.getAttribute('data-sheet');
          const range = e.target.getAttribute('data-range');
          const label = e.target.getAttribute('data-label');
          if (sheet && range) addPinnedRange(sheet, range, label || `${sheet}!${range}`);
        } else if (e.target.classList.contains('sheet-custom-add')) {
          const sheet = e.target.getAttribute('data-sheet');
          const input = e.target.previousElementSibling;
          const range = input && input.value ? input.value.trim() : '';
          if (sheet && range) {
            addPinnedRange(sheet, range);
            if (input) input.value = '';
          }
        }
      });
    }

    // Event delegation for pinned chips (remove, toggle mode)
    const pinnedChips = document.getElementById('pinned-chips');
    if (pinnedChips) {
      pinnedChips.addEventListener('click', (e) => {
        const id = e.target.getAttribute('data-id');
        if (!id) return;
        if (e.target.classList.contains('pinned-chip-remove')) {
          removePinnedRange(id);
        } else if (e.target.classList.contains('pinned-chip-label')) {
          togglePinnedRangeMode(id);
        }
      });
    }

    // Event delegation for clickable cell citations
    document.getElementById('message-thread').addEventListener('click', (e) => {
      if (e.target.classList.contains('cell-citation')) {
        e.preventDefault();
        const ref = e.target.getAttribute('data-ref');
        if (ref) navigateToCell(ref);
      }
    });
    const input = document.getElementById('message-input');
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        onSend();
      }
    });
    input.addEventListener('input', autoResizeTextarea);

    // Check backend health on load
    checkBackendHealth();

    // Suggestion chips
    document.querySelectorAll('.chip').forEach((chip) => {
      chip.addEventListener('click', () => {
        const prompt = chip.getAttribute('data-prompt');
        if (prompt) {
          input.value = prompt;
          autoResizeTextarea({ target: input });
          onSend();
        }
      });
    });
  }
});

function autoResizeTextarea(e) {
  const ta = e && e.target ? e.target : document.getElementById('message-input');
  if (!ta) return;
  ta.style.height = 'auto';
  ta.style.height = Math.min(ta.scrollHeight, 120) + 'px';
}

function exportChatHistory() {
  if (conversationHistory.length === 0) {
    addMessage('No conversation to export.', 'assistant', { className: 'tool-status' });
    return;
  }

  let markdown = '# CelloBot Chat Export\n';
  markdown += `Date: ${new Date().toISOString()}\n\n---\n\n`;

  for (const msg of conversationHistory) {
    const role = msg.role === 'user' ? 'You' : 'CelloBot';
    markdown += `**${role}:**\n${msg.content}\n\n`;
  }

  // Create and trigger download
  const blob = new Blob([markdown], { type: 'text/markdown' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `cellobot-chat-${new Date().toISOString().slice(0, 10)}.md`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

async function checkBackendHealth() {
  try {
    const res = await fetch(`${API_BASE}/health`, { method: 'GET' });
    if (!res.ok) throw new Error('Backend returned ' + res.status);
  } catch (err) {
    console.warn('Backend health check failed:', err);
    const banner = document.createElement('div');
    banner.className = 'message assistant';
    banner.style.color = '#e57373';
    banner.textContent = 'Unable to connect to CelloBot backend. Please check your connection or try again later.';
    document.getElementById('message-thread').appendChild(banner);
  }
}

function onNewChat() {
  conversationHistory = [];
  pinnedRanges = [];
  document.getElementById('message-thread').innerHTML = '';
  clearChangeLog();
  updateContextUI();
  showWelcomeScreen();
}

// Rough token estimator: ~4 characters per token for English text
function estimateTokens(messages) {
  let total = 0;
  for (const msg of messages) {
    if (typeof msg.content === 'string') {
      total += Math.ceil(msg.content.length / 4);
    }
  }
  return total;
}

const MAX_CONVERSATION_TOKENS = 80000;
const COMPACTION_KEEP_RECENT = 6; // Keep last 6 messages (3 turns)

async function maybeCompactHistory() {
  if (conversationHistory.length <= COMPACTION_KEEP_RECENT) return;

  const tokenCount = estimateTokens(conversationHistory);
  if (tokenCount < MAX_CONVERSATION_TOKENS) return;

  // Summarize the older messages
  const olderMessages = conversationHistory.slice(0, -COMPACTION_KEEP_RECENT);
  const recentMessages = conversationHistory.slice(-COMPACTION_KEEP_RECENT);

  // Build a summary using the backend
  try {
    const modelSelect = document.getElementById('model-select');
    const res = await fetch(`${API_BASE}/chat`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        messages: [
          ...olderMessages,
          { role: 'user', content: 'Please summarize our conversation so far in 300 words, focusing on what changes were made to the workbook and key decisions. This summary will be used to maintain context in a new conversation.' }
        ],
        model: modelSelect.value,
        context: {}
      })
    });

    const data = await res.json();
    const summary = data.text || data.response || '';

    if (summary) {
      // Replace conversation history with summary + recent messages
      conversationHistory = [
        { role: 'user', content: '[Previous conversation summary]: ' + summary },
        { role: 'assistant', content: 'I have the context from our previous conversation. How can I help you next?' },
        ...recentMessages
      ];
      addMessage('(Conversation compacted to save context space)', 'assistant', { className: 'tool-status' });
    }
  } catch (e) {
    console.warn('Auto-compaction failed:', e);
  }
}

function hideWelcomeScreen() {
  const welcome = document.getElementById('welcome-screen');
  if (welcome) welcome.classList.add('hidden');
}

function showWelcomeScreen() {
  const welcome = document.getElementById('welcome-screen');
  if (welcome) welcome.classList.remove('hidden');
}

function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

function parseMarkdown(text) {
  const escaped = escapeHtml(text);
  const codeBlockPlaceholders = [];

  // Extract code blocks first and replace with placeholders
  let result = escaped.replace(/```([\s\S]*?)```/g, (_, code) => {
    const trimmed = code.trim();
    const idx = codeBlockPlaceholders.length;
    codeBlockPlaceholders.push('<pre><code>' + trimmed + '</code></pre>');
    return '\x00CODEBLOCK' + idx + '\x00';
  });

  // Inline code (`...`)
  result = result.replace(/`([^`]+)`/g, '<code>$1</code>');

  // Bold (**text**)
  result = result.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');

  // Line breaks
  result = result.replace(/\n/g, '<br>');

  // Cell citations: [[Sheet1!A1]] or [[A1]] become clickable
  result = result.replace(/\[\[([^\]]+)\]\]/g, function(_, ref) {
    return '<a class="cell-citation" href="#" data-ref="' + ref + '">' + ref + '</a>';
  });

  // Restore code blocks
  codeBlockPlaceholders.forEach((html, idx) => {
    result = result.replace('\x00CODEBLOCK' + idx + '\x00', html);
  });

  return result;
}

function addLoadingMessage(id) {
  const div = document.createElement('div');
  div.className = 'message assistant loading';
  div.id = id;
  const loader = document.createElement('span');
  loader.className = 'loader-dots';
  loader.innerHTML = '<span></span><span></span><span></span>';
  div.appendChild(loader);
  document.getElementById('message-thread').appendChild(div);
  div.scrollIntoView({ behavior: 'smooth' });
}

async function getContext() {
  return Excel.run(async (context) => {
    const workbook = context.workbook;
    const sheet = workbook.worksheets.getActiveWorksheet();
    const selectedRange = workbook.getSelectedRange();

    // Load active sheet info
    sheet.load('name');
    selectedRange.load('address,formulas,values,rowIndex,columnIndex,rowCount,columnCount');

    // Load all sheet names
    const sheets = workbook.worksheets;
    sheets.load('items/name');

    await context.sync();

    const activeSheetName = sheet.name;
    const address = selectedRange.address.replace(/^.*!/, '');

    // --- Selected cell formula ---
    const formulas = selectedRange.formulas;
    let selectedFormula = null;
    let selectedValue = null;

    if (formulas && formulas.length > 0 && formulas[0].length > 0) {
      const cellFormula = formulas[0][0];
      if (typeof cellFormula === 'string' && cellFormula.startsWith('=')) {
        selectedFormula = cellFormula;
      }
    }
    if (selectedRange.values && selectedRange.values.length > 0) {
      selectedValue = selectedRange.values[0][0];
    }

    // --- Sheet list with dimensions ---
    const sheetSummaries = [];
    for (const ws of sheets.items) {
      sheetSummaries.push(ws.name);
    }

    // --- Headers from used range (not hardcoded A-J) ---
    let headers = [];
    let usedRangeAddress = null;
    try {
      const usedRange = sheet.getUsedRange();
      usedRange.load('address,rowCount,columnCount');
      await context.sync();
      usedRangeAddress = usedRange.address.replace(/^.*!/, '');

      // Read first row as headers
      const headerRow = usedRange.getRow(0);
      headerRow.load('values');
      await context.sync();
      if (headerRow.values && headerRow.values[0]) {
        headers = headerRow.values[0].map(v => (v != null ? String(v).trim() : '')).filter(Boolean);
      }
    } catch (e) {
      // Sheet may be empty — no used range
    }

    // --- Nearby formulas (expanded range, no cap) ---
    const nearbyFormulas = [];
    try {
      const expanded = selectedRange.getOffsetRange(-2, -2).getBoundingRect(selectedRange.getOffsetRange(2, 2));
      expanded.load('formulas,address');
      await context.sync();
      if (expanded.formulas) {
        expanded.formulas.forEach((row, r) => {
          (row || []).forEach((cell, c) => {
            if (typeof cell === 'string' && cell.startsWith('=')) {
              const cellAddr = getCellAddress(expanded.address, r, c);
              if (cellAddr !== address) {
                nearbyFormulas.push({ address: cellAddr, formula: cell });
              }
            }
          });
        });
      }
    } catch (e) {
      // Ignore if range is out of bounds
    }

    // --- Detect errors in selected range and nearby ---
    const errors = [];
    try {
      const scanRange = selectedRange.getOffsetRange(-2, -2).getBoundingRect(selectedRange.getOffsetRange(2, 2));
      scanRange.load('values,address');
      await context.sync();
      if (scanRange.values) {
        const errorTypes = ['#REF!', '#VALUE!', '#N/A', '#DIV/0!', '#NAME?', '#NULL!', '#NUM!'];
        scanRange.values.forEach((row, r) => {
          (row || []).forEach((cell, c) => {
            if (typeof cell === 'string' && errorTypes.includes(cell)) {
              const cellAddr = getCellAddress(scanRange.address, r, c);
              errors.push({ address: cellAddr, error: cell });
            }
          });
        });
      }
    } catch (e) {
      // Ignore scan errors
    }

    // --- Tables on active sheet ---
    const tableNames = [];
    try {
      const tables = sheet.tables;
      tables.load('items/name,items/columns/items/name');
      await context.sync();
      for (const table of tables.items) {
        const colNames = table.columns.items.map(col => col.name);
        tableNames.push({ name: table.name, columns: colNames });
      }
    } catch (e) {
      // Tables API may not be available
    }

    // --- Named ranges ---
    const namedRanges = [];
    try {
      const names = workbook.names;
      names.load('items/name,items/value');
      await context.sync();
      for (const n of names.items) {
        namedRanges.push({ name: n.name, value: n.value });
      }
    } catch (e) {
      // Named ranges may not be accessible
    }

    return {
      activeSheet: activeSheetName,
      sheets: sheetSummaries,
      selectedAddress: address,
      selectedFormula,
      selectedValue,
      usedRange: usedRangeAddress,
      headers,
      nearbyFormulas,
      errors,
      tables: tableNames,
      namedRanges
    };
  });
}

function getCellAddress(regionAddress, row, col) {
  const addr = regionAddress.replace(/^.*!/, '');
  const match = addr.match(/\$?([A-Z]+)\$?(\d+)/);
  if (!match) return addr;
  const startCol = match[1];
  const startRow = parseInt(match[2], 10);
  const c = colLettersToIndex(startCol) + col;
  const colLetters = indexToColLetters(c);
  return colLetters + (startRow + row);
}

function colLettersToIndex(letters) {
  let idx = 0;
  for (let i = 0; i < letters.length; i++) {
    idx = idx * 26 + (letters.charCodeAt(i) - 64);
  }
  return idx - 1;
}

function indexToColLetters(idx) {
  let result = '';
  idx++;
  while (idx > 0) {
    idx--;
    result = String.fromCharCode(65 + (idx % 26)) + result;
    idx = Math.floor(idx / 26);
  }
  return result;
}

function addMessage(content, role, options = {}) {
  hideWelcomeScreen();

  const div = document.createElement('div');
  div.className = `message ${role}`;
  if (options.className) div.classList.add(options.className);
  if (options.id) div.id = options.id;

  if (role === 'assistant' && !options.loading && !options.className) {
    div.innerHTML = parseMarkdown(content);
  } else {
    div.textContent = content;
  }

  if (options.formula) {
    const btn = document.createElement('button');
    btn.className = 'insert-formula';
    btn.textContent = 'Insert into cell';
    btn.onclick = () => insertFormula(options.formula);
    div.appendChild(document.createElement('br'));
    div.appendChild(btn);
  }

  document.getElementById('message-thread').appendChild(div);
  div.scrollIntoView({ behavior: 'smooth' });
}

function navigateToCell(ref) {
  Excel.run(async (context) => {
    // Handle references like "Sheet1!A1" or just "A1"
    let range;
    if (ref.includes('!')) {
      const parts = ref.split('!');
      const sheetName = parts[0].replace(/'/g, '');
      const address = parts[1];
      const sheet = context.workbook.worksheets.getItem(sheetName);
      sheet.activate();
      range = sheet.getRange(address);
    } else {
      range = context.workbook.worksheets.getActiveWorksheet().getRange(ref);
    }
    range.select();
    await context.sync();
  }).catch(err => {
    console.warn('Could not navigate to cell:', ref, err);
  });
}

function insertFormula(formula) {
  const cleanFormula = formula.trim().startsWith('=') ? formula.trim() : '=' + formula.trim();
  Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.formulas = [[cleanFormula]];
    await context.sync();

    // Read back the result to check for errors
    range.load('values,address');
    await context.sync();

    const result = range.values[0][0];
    const address = range.address.replace(/^.*!/, '');
    const errorTypes = ['#REF!', '#VALUE!', '#N/A', '#DIV/0!', '#NAME?', '#NULL!', '#NUM!'];

    if (typeof result === 'string' && errorTypes.includes(result)) {
      addMessage(`Formula inserted into ${address} but produced error: ${result}`, 'assistant', { className: 'tool-status' });
    } else {
      addMessage(`Formula inserted into ${address}`, 'assistant', { className: 'tool-status' });
    }
  }).catch(err => {
    addMessage('Failed to insert: ' + err.message, 'assistant');
  });
}

async function onSend() {
  const input = document.getElementById('message-input');
  const message = input.value.trim();
  if (!message) return;

  const sendBtn = document.getElementById('send-btn');
  const modelSelect = document.getElementById('model-select');

  sendBtn.disabled = true;
  input.value = '';
  autoResizeTextarea({ target: input });
  addMessage(message, 'user');

  // Add user message to conversation history
  conversationHistory.push({ role: 'user', content: message });

  // Auto-compact if conversation is getting too long
  await maybeCompactHistory();

  const loadingId = 'loading-' + Date.now();
  addLoadingMessage(loadingId);

  let context = {};
  try {
    context = await getContext();
  } catch (err) {
    console.warn('Could not get context:', err);
  }

  if (pinnedRanges.length > 0) {
    try {
      context.pinnedRanges = await getPinnedContext();
    } catch (err) {
      console.warn('Could not read pinned context:', err);
    }
  }

  try {
    await processChat(conversationHistory, modelSelect.value, context, loadingId);
  } catch (err) {
    const loadingEl = document.getElementById(loadingId);
    if (loadingEl) loadingEl.remove();
    addMessage('Error: ' + (err.message || 'Network error'), 'assistant');
  } finally {
    sendBtn.disabled = false;
  }
}

/**
 * Process a chat request using SSE streaming with tool-use loop support.
 * Streams text incrementally, handles tool calls, and recurses for multi-turn.
 */
async function processChat(messages, modelValue, context, loadingId, pendingMessages, toolResults) {
  const body = {};

  if (pendingMessages && toolResults) {
    body.pendingMessages = pendingMessages;
    body.toolResults = toolResults;
    body.model = modelValue;
    body.context = context;
  } else {
    body.messages = messages;
    body.model = modelValue;
    body.context = context;
  }

  return new Promise((resolve, reject) => {
    // Remove loading indicator — we'll show streamed text instead
    const loadingEl = document.getElementById(loadingId);
    if (loadingEl) loadingEl.remove();

    // Create a message element for streaming text into
    let streamDiv = null;
    let fullText = '';

    fetch(`${API_BASE}/stream`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    }).then(response => {
      if (!response.ok) {
        return response.text().then(text => {
          addMessage('Request failed: ' + text, 'assistant');
          resolve();
        });
      }

      const reader = response.body.getReader();
      const decoder = new TextDecoder();
      let buffer = '';

      function processChunk() {
        reader.read().then(({ done, value }) => {
          if (done) {
            // Stream ended — finalize if we have accumulated text
            if (fullText && streamDiv) {
              streamDiv.innerHTML = parseMarkdown(fullText);
              const isFormula = fullText.trim().startsWith('=');
              if (isFormula) {
                const btn = document.createElement('button');
                btn.className = 'insert-formula';
                btn.textContent = 'Insert into cell';
                btn.onclick = () => insertFormula(fullText.trim());
                streamDiv.appendChild(document.createElement('br'));
                streamDiv.appendChild(btn);
              }
            }
            resolve();
            return;
          }

          buffer += decoder.decode(value, { stream: true });
          const lines = buffer.split('\n');
          buffer = lines.pop(); // Keep incomplete line in buffer

          for (const line of lines) {
            if (line.startsWith('event: ')) {
              var currentEvent = line.substring(7).trim();
            } else if (line.startsWith('data: ') && currentEvent) {
              try {
                const data = JSON.parse(line.substring(6));
                handleSSEEvent(currentEvent, data);
              } catch (e) {
                // Ignore parse errors
              }
              currentEvent = null;
            }
          }

          processChunk();
        }).catch(err => {
          addMessage('Stream error: ' + err.message, 'assistant');
          resolve();
        });
      }

      function handleSSEEvent(event, data) {
        switch (event) {
          case 'text_delta':
            if (!streamDiv) {
              hideWelcomeScreen();
              streamDiv = document.createElement('div');
              streamDiv.className = 'message assistant';
              document.getElementById('message-thread').appendChild(streamDiv);
            }
            fullText += data.text;
            // Update with raw text during streaming, parse markdown on completion
            streamDiv.textContent = fullText;
            streamDiv.scrollIntoView({ behavior: 'smooth' });
            break;

          case 'tool_calls':
            // LLM wants to execute tools
            if (data.text && !streamDiv) {
              addMessage(data.text, 'assistant');
            }

            // Execute tools then continue the loop
            (async () => {
              const results = [];
              for (const toolCall of data.toolCalls) {
                try {
                  addMessage(`Executing: ${toolCall.name}...`, 'assistant', { className: 'tool-status' });
                  const output = await executeTool(toolCall.name, toolCall.input);
                  results.push({ id: toolCall.id, output });
                } catch (err) {
                  results.push({ id: toolCall.id, output: { success: false, error: err.message } });
                }
              }

              // Recurse: send results back for next LLM turn
              const newLoadingId = 'loading-' + Date.now();
              addLoadingMessage(newLoadingId);
              await processChat(messages, modelValue, context, newLoadingId, data.messages, results);
              resolve();
            })();
            return; // Don't continue reading — new processChat call handles it

          case 'done':
            if (data.text) {
              fullText = data.text;
              conversationHistory.push({ role: 'assistant', content: fullText });

              if (streamDiv) {
                streamDiv.innerHTML = parseMarkdown(fullText);
                const isFormula = fullText.trim().startsWith('=');
                if (isFormula) {
                  const btn = document.createElement('button');
                  btn.className = 'insert-formula';
                  btn.textContent = 'Insert into cell';
                  btn.onclick = () => insertFormula(fullText.trim());
                  streamDiv.appendChild(document.createElement('br'));
                  streamDiv.appendChild(btn);
                }
              } else {
                const isFormula = fullText.trim().startsWith('=');
                addMessage(fullText, 'assistant', isFormula ? { formula: fullText.trim() } : {});
              }
            }
            resolve();
            break;

          case 'error':
            addMessage('Error: ' + (data.error || 'Unknown error'), 'assistant');
            resolve();
            break;
        }
      }

      processChunk();
    }).catch(err => {
      addMessage('Error: ' + (err.message || 'Network error'), 'assistant');
      resolve();
    });
  });
}
