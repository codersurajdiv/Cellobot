/* global Excel */

/**
 * Change Tracker: Records before/after state of cells modified by CelloBot.
 * Provides undo functionality and change highlighting.
 */

// Stores all changes made in the current session
const changeLog = [];

// Highlight color for modified cells
const HIGHLIGHT_COLOR = '#2D2B55'; // Subtle purple tint matching the theme

/**
 * Record a change before it's made.
 * Captures the current state of the target range.
 * @param {string} sheet - Sheet name
 * @param {string} range - Range address
 * @returns {Promise<Object>} beforeState object to pass to recordAfter
 */
async function recordBefore(sheet, range) {
  try {
    return await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem(sheet);
      const r = ws.getRange(range);
      r.load('values,formulas,numberFormat,address');
      await ctx.sync();

      return {
        sheet,
        range: r.address.replace(/^.*!/, ''),
        values: r.values,
        formulas: r.formulas,
        numberFormat: r.numberFormat,
        timestamp: new Date().toISOString()
      };
    });
  } catch (e) {
    return { sheet, range, values: null, formulas: null, timestamp: new Date().toISOString() };
  }
}

/**
 * Record the state after a change is made and add to the change log.
 * @param {Object} beforeState - From recordBefore
 * @param {string} toolName - Name of the tool that made the change
 * @param {Object} toolInput - Input parameters of the tool
 */
async function recordAfter(beforeState, toolName, toolInput) {
  try {
    const afterState = await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem(beforeState.sheet);
      const r = ws.getRange(beforeState.range);
      r.load('values,formulas,numberFormat');
      await ctx.sync();
      return {
        values: r.values,
        formulas: r.formulas,
        numberFormat: r.numberFormat
      };
    });

    const entry = {
      id: 'change-' + Date.now() + '-' + Math.random().toString(36).substring(7),
      toolName,
      toolInput,
      sheet: beforeState.sheet,
      range: beforeState.range,
      before: {
        values: beforeState.values,
        formulas: beforeState.formulas,
        numberFormat: beforeState.numberFormat
      },
      after: {
        values: afterState.values,
        formulas: afterState.formulas,
        numberFormat: afterState.numberFormat
      },
      timestamp: beforeState.timestamp
    };

    changeLog.push(entry);
    highlightChangedCells(beforeState.sheet, beforeState.range);
    updateChangeUI();
    return entry;
  } catch (e) {
    console.warn('Failed to record after state:', e);
  }
}

/**
 * Apply a subtle highlight to modified cells.
 */
async function highlightChangedCells(sheet, range) {
  try {
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem(sheet);
      const r = ws.getRange(range);
      r.format.fill.color = HIGHLIGHT_COLOR;
      await ctx.sync();
    });
  } catch (e) {
    // Silently ignore highlight failures
  }
}

/**
 * Undo all changes made by CelloBot in reverse order.
 * @returns {number} The number of changes reverted.
 */
async function undoAllChanges() {
  if (changeLog.length === 0) return 0;

  const count = changeLog.length;

  // Process in reverse order
  const changes = [...changeLog].reverse();

  for (const entry of changes) {
    try {
      await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItem(entry.sheet);
        const r = ws.getRange(entry.range);

        // Restore original values/formulas
        if (entry.before.formulas) {
          r.formulas = entry.before.formulas;
        } else if (entry.before.values) {
          r.values = entry.before.values;
        }

        // Restore number format
        if (entry.before.numberFormat) {
          r.numberFormat = entry.before.numberFormat;
        }

        // Clear highlight
        r.format.fill.clear();

        await ctx.sync();
      });
    } catch (e) {
      console.warn('Failed to undo change:', entry.id, e);
    }
  }

  // Clear the log
  changeLog.length = 0;
  updateChangeUI();
  return count;
}

/**
 * Get the current change log.
 */
function getChangeLog() {
  return [...changeLog];
}

/**
 * Clear the change log (e.g. on new chat).
 */
function clearChangeLog() {
  changeLog.length = 0;
  updateChangeUI();
}

/**
 * Write the change log to a "CelloBot Log" sheet in the workbook.
 */
async function writeSessionLog() {
  if (changeLog.length === 0) return;

  try {
    await Excel.run(async (ctx) => {
      let logSheet;
      try {
        logSheet = ctx.workbook.worksheets.getItem('CelloBot Log');
      } catch (e) {
        // Create the log sheet if it doesn't exist
        logSheet = ctx.workbook.worksheets.add('CelloBot Log');
        // Write headers
        const headers = logSheet.getRange('A1:F1');
        headers.values = [['Timestamp', 'Action', 'Sheet', 'Range', 'Before', 'After']];
        headers.format.font.bold = true;
        await ctx.sync();
      }

      // Find the next empty row
      const usedRange = logSheet.getUsedRange();
      usedRange.load('rowCount');
      await ctx.sync();
      const nextRow = usedRange.rowCount + 1;

      // Write each log entry
      const rows = changeLog.map(entry => [
        entry.timestamp,
        entry.toolName,
        entry.sheet,
        entry.range,
        entry.before.values ? JSON.stringify(entry.before.values).substring(0, 200) : '',
        entry.after.values ? JSON.stringify(entry.after.values).substring(0, 200) : ''
      ]);

      if (rows.length > 0) {
        const targetRange = logSheet.getRange(`A${nextRow}:F${nextRow + rows.length - 1}`);
        targetRange.values = rows;
        await ctx.sync();
      }
    });
  } catch (e) {
    console.warn('Failed to write session log:', e);
  }
}

/**
 * Update the change tracking UI in the task pane.
 */
function updateChangeUI() {
  const container = document.getElementById('change-tracker');
  if (!container) return;

  if (changeLog.length === 0) {
    container.classList.add('hidden');
    return;
  }

  container.classList.remove('hidden');
  const countEl = container.querySelector('.change-count');
  if (countEl) {
    countEl.textContent = `${changeLog.length} change${changeLog.length === 1 ? '' : 's'} made`;
  }
}
