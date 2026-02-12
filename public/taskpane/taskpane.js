/* global Office, Excel */

const API_BASE = 'https://cellobot-production.up.railway.app';

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById('send-btn').addEventListener('click', onSend);
    document.getElementById('message-input').addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        onSend();
      }
    });
  }
});

async function getContext() {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const selectedRange = context.workbook.getSelectedRange();

    selectedRange.load('address', 'formulas', 'values', 'rowIndex', 'columnIndex');
    await context.sync();

    const address = selectedRange.address.replace(/^.*!/, '');
    const formulas = selectedRange.formulas;
    let selectedFormula = null;
    const nearbyFormulas = [];
    let headers = [];

    if (formulas && formulas.length > 0 && formulas[0].length > 0) {
      const cellValue = formulas[0][0];
      if (typeof cellValue === 'string' && cellValue.startsWith('=')) {
        selectedFormula = cellValue;
      }
    }

    const rowIndex = selectedRange.rowIndex;
    if (rowIndex > 0) {
      const headerAddr = 'A' + rowIndex + ':J' + rowIndex;
      const headerRow = sheet.getRange(headerAddr);
      headerRow.load('values');
      await context.sync();
      if (headerRow.values && headerRow.values[0]) {
        headers = headerRow.values[0].map(v => (v != null ? String(v).trim() : '')).filter(Boolean);
      }
    }

    try {
      const expanded = selectedRange.getOffsetRange(-1, -1).getBoundingRect(selectedRange.getOffsetRange(1, 1));
      expanded.load('formulas', 'address');
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

    return {
      selectedAddress: address,
      selectedFormula,
      nearbyFormulas: nearbyFormulas.slice(0, 5),
      headers
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
  const div = document.createElement('div');
  div.className = `message ${role}`;
  if (options.id) div.id = options.id;
  const text = document.createTextNode(content);
  div.appendChild(text);
  if (options.formula) {
    const btn = document.createElement('button');
    btn.className = 'insert-formula';
    btn.textContent = 'Insert into cell';
    btn.onclick = () => insertFormula(options.formula);
    div.appendChild(document.createElement('br'));
    div.appendChild(btn);
  }
  document.getElementById('chat-messages').appendChild(div);
  div.scrollIntoView({ behavior: 'smooth' });
}

function insertFormula(formula) {
  const cleanFormula = formula.trim().startsWith('=') ? formula.trim() : '=' + formula.trim();
  Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.formulas = [[cleanFormula]];
    await context.sync();
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
  addMessage(message, 'user');

  const loadingId = 'loading-' + Date.now();
  addMessage('Thinking...', 'assistant', { id: loadingId });

  let context = {};
  try {
    context = await getContext();
  } catch (err) {
    console.warn('Could not get context:', err);
  }

  try {
    const res = await fetch(`${API_BASE}/chat`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        message,
        model: modelSelect.value,
        context
      })
    });

    const loadingEl = document.getElementById(loadingId);
    if (loadingEl) loadingEl.remove();

    const data = await res.json();
    if (!res.ok) {
      addMessage(data.error || 'Request failed', 'assistant');
      return;
    }

    const response = data.response;
    const isFormula = response.trim().startsWith('=');
    addMessage(response, 'assistant', isFormula ? { formula: response.trim() } : {});
  } catch (err) {
    const loadingEl = document.getElementById(loadingId);
    if (loadingEl) loadingEl.remove();
    addMessage('Error: ' + (err.message || 'Network error'), 'assistant');
  } finally {
    sendBtn.disabled = false;
  }
}
