/* global Excel */

/**
 * Client-side tool executor.
 * Maps tool names from the LLM to Excel JavaScript API calls.
 */

// Tools that modify the workbook and should be tracked
const TRACKED_TOOLS = ['write_cells', 'write_cells_force', 'format_cells', 'sort_range', 'add_conditional_format', 'set_data_validation'];

async function executeTool(toolName, toolInput) {
  // Record before state for tracked tools
  let beforeState = null;
  if (TRACKED_TOOLS.includes(toolName) && toolInput.sheet && toolInput.range) {
    try {
      beforeState = await recordBefore(toolInput.sheet, toolInput.range);
    } catch (e) {
      // Continue even if recording fails
    }
  }

  let result;
  switch (toolName) {
    case 'write_cells':
      result = await writeCells(toolInput);
      break;
    case 'write_cells_force':
      result = await writeCellsForce(toolInput);
      break;
    case 'read_range':
      return await readRange(toolInput);
    case 'get_workbook_info':
      return await getWorkbookInfo();
    case 'format_cells':
      result = await formatCells(toolInput);
      break;
    case 'create_chart':
      return await createChart(toolInput);
    case 'sort_range':
      result = await sortRange(toolInput);
      break;
    case 'filter_data':
      return await filterData(toolInput);
    case 'set_data_validation':
      result = await setDataValidation(toolInput);
      break;
    case 'add_conditional_format':
      result = await addConditionalFormat(toolInput);
      break;
    case 'set_column_width':
      return await setColumnWidth(toolInput);
    case 'set_row_height':
      return await setRowHeight(toolInput);
    case 'toggle_gridlines':
      return await toggleGridlines(toolInput);
    case 'set_print_area':
      return await setPrintArea(toolInput);
    case 'add_worksheet':
      return await addWorksheet(toolInput);
    case 'trace_formula':
      return await traceFormula(toolInput);
    case 'find_errors':
      return await findErrors(toolInput);
    case 'edit_chart':
      return await editChart(toolInput);
    case 'create_pivot_table':
      return await createPivotTable(toolInput);
    case 'refresh_pivot_table':
      return await refreshPivotTable(toolInput);
    default:
      return { success: false, error: `Unknown tool: ${toolName}` };
  }

  // Record after state for tracked tools
  if (beforeState && result && result.success) {
    try {
      await recordAfter(beforeState, toolName, toolInput);
    } catch (e) {
      // Continue even if recording fails
    }
  }

  return result;
}

// ==================== Tool Handlers ====================

async function writeCells(input) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(input.sheet);
    const range = sheet.getRange(input.range);

    // Check for existing content (overwrite protection)
    range.load('values');
    await ctx.sync();

    let nonEmptyCount = 0;
    if (range.values) {
      for (const row of range.values) {
        for (const cell of row) {
          if (cell !== '' && cell !== null && cell !== undefined) {
            nonEmptyCount++;
          }
        }
      }
    }

    if (nonEmptyCount > 0) {
      return {
        pending: true,
        nonEmptyCount,
        sheet: input.sheet,
        range: input.range
      };
    }

    if (input.formulas) {
      range.formulas = input.formulas;
    } else if (input.values) {
      range.values = input.values;
    } else {
      return { success: false, error: 'Either values or formulas must be provided' };
    }

    await ctx.sync();
    return { success: true, message: `Wrote to ${input.sheet}!${input.range}` };
  });
}

async function writeCellsForce(input) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(input.sheet);
    const range = sheet.getRange(input.range);

    if (input.formulas) {
      range.formulas = input.formulas;
    } else if (input.values) {
      range.values = input.values;
    } else {
      return { success: false, error: 'Either values or formulas must be provided' };
    }

    await ctx.sync();
    return { success: true, message: `Wrote to ${input.sheet}!${input.range}` };
  });
}

async function readRange(input) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(input.sheet);
    const range = sheet.getRange(input.range);
    range.load('values,formulas,address,rowCount,columnCount,numberFormat');
    await ctx.sync();

    return {
      success: true,
      address: range.address,
      values: range.values,
      formulas: range.formulas,
      rowCount: range.rowCount,
      columnCount: range.columnCount
    };
  });
}

async function getWorkbookInfo() {
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load('items/name');
    await ctx.sync();

    const sheetDetails = [];
    for (const ws of sheets.items) {
      try {
        const usedRange = ws.getUsedRange();
        usedRange.load('address,rowCount,columnCount');
        await ctx.sync();
        sheetDetails.push({
          name: ws.name,
          usedRange: usedRange.address.replace(/^.*!/, ''),
          rows: usedRange.rowCount,
          columns: usedRange.columnCount
        });
      } catch (e) {
        sheetDetails.push({ name: ws.name, usedRange: null, rows: 0, columns: 0 });
      }
    }

    // Named ranges
    const namedRanges = [];
    try {
      const names = ctx.workbook.names;
      names.load('items/name,items/value');
      await ctx.sync();
      for (const n of names.items) {
        namedRanges.push({ name: n.name, value: n.value });
      }
    } catch (e) {
      // Named ranges may not be accessible
    }

    return {
      success: true,
      sheets: sheetDetails,
      namedRanges
    };
  });
}

async function formatCells(input) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(input.sheet);
    const range = sheet.getRange(input.range);

    if (input.numberFormat) {
      range.numberFormat = [[input.numberFormat]];
    }
    if (input.bold !== undefined) {
      range.format.font.bold = input.bold;
    }
    if (input.italic !== undefined) {
      range.format.font.italic = input.italic;
    }
    if (input.fontSize) {
      range.format.font.size = input.fontSize;
    }
    if (input.fontColor) {
      range.format.font.color = input.fontColor;
    }
    if (input.fillColor) {
      range.format.fill.color = input.fillColor;
    }
    if (input.horizontalAlignment) {
      range.format.horizontalAlignment = input.horizontalAlignment;
    }
    if (input.verticalAlignment) {
      range.format.verticalAlignment = input.verticalAlignment;
    }
    if (input.wrapText !== undefined) {
      range.format.wrapText = input.wrapText;
    }
    if (input.borderStyle && input.borderStyle !== 'None') {
      const borders = range.format.borders;
      const sides = ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight'];
      for (const side of sides) {
        const border = borders.getItem(side);
        border.style = input.borderStyle;
        if (input.borderColor) {
          border.color = input.borderColor;
        }
      }
    }
    if (input.merge === true) {
      range.merge(false);
    } else if (input.merge === false) {
      range.unmerge();
    }

    await ctx.sync();
    return { success: true, message: `Formatted ${input.sheet}!${input.range}` };
  });
}

async function createChart(input) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(input.sheet);
    const dataRange = sheet.getRange(input.dataRange);
    const seriesBy = input.seriesBy || 'Auto';

    const chart = sheet.charts.add(input.chartType, dataRange, seriesBy);

    if (input.title) {
      chart.title.text = input.title;
      chart.title.visible = true;
    }

    if (input.position) {
      const anchorCell = sheet.getRange(input.position);
      chart.setPosition(anchorCell);
    }

    await ctx.sync();
    return { success: true, message: `Created ${input.chartType} chart on ${input.sheet}` };
  });
}

async function sortRange(input) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(input.sheet);
    const range = sheet.getRange(input.range);

    const sortFields = input.sortBy.map(criteria => ({
      key: criteria.column,
      ascending: criteria.ascending
    }));

    range.sort.apply(sortFields, input.hasHeaders !== false);
    await ctx.sync();
    return { success: true, message: `Sorted ${input.sheet}!${input.range}` };
  });
}

async function filterData(input) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(input.sheet);
    const range = sheet.getRange(input.range);

    range.autoFilter.apply(range, input.column, {
      criterion1: input.values[0],
      filterOn: 'Values',
      values: input.values
    });

    await ctx.sync();
    return { success: true, message: `Filtered column ${input.column} in ${input.sheet}!${input.range}` };
  });
}

async function setDataValidation(input) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(input.sheet);
    const range = sheet.getRange(input.range);

    let rule = {};
    switch (input.type) {
      case 'list':
        rule = { list: { inCellDropDown: true, source: input.listSource } };
        break;
      case 'wholeNumber':
        rule = { wholeNumber: { operator: input.operator, formula1: input.formula1, formula2: input.formula2 } };
        break;
      case 'decimal':
        rule = { decimal: { operator: input.operator, formula1: input.formula1, formula2: input.formula2 } };
        break;
      case 'date':
        rule = { date: { operator: input.operator, formula1: input.formula1, formula2: input.formula2 } };
        break;
      case 'textLength':
        rule = { textLength: { operator: input.operator, formula1: input.formula1, formula2: input.formula2 } };
        break;
      case 'custom':
        rule = { custom: { formula: input.customFormula } };
        break;
    }

    range.dataValidation.rule = rule;

    if (input.errorMessage) {
      range.dataValidation.errorAlert = {
        showAlert: true,
        title: 'Invalid Input',
        message: input.errorMessage,
        style: 'Stop'
      };
    }

    await ctx.sync();
    return { success: true, message: `Set ${input.type} validation on ${input.sheet}!${input.range}` };
  });
}

async function addConditionalFormat(input) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(input.sheet);
    const range = sheet.getRange(input.range);

    switch (input.ruleType) {
      case 'cellValue': {
        const cf = range.conditionalFormats.add('CellValue');
        cf.cellValue.rule = {
          formula1: input.formula1,
          formula2: input.formula2 || undefined,
          operator: input.operator
        };
        if (input.fontColor) cf.cellValue.format.font.color = input.fontColor;
        if (input.fillColor) cf.cellValue.format.fill.color = input.fillColor;
        break;
      }
      case 'colorScale': {
        range.conditionalFormats.add('ColorScale');
        break;
      }
      case 'dataBar': {
        range.conditionalFormats.add('DataBar');
        break;
      }
      case 'iconSet': {
        range.conditionalFormats.add('IconSet');
        break;
      }
    }

    await ctx.sync();
    return { success: true, message: `Applied ${input.ruleType} conditional format to ${input.sheet}!${input.range}` };
  });
}

async function setColumnWidth(input) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(input.sheet);
    const range = sheet.getRange(input.range);

    if (input.autoFit) {
      range.format.autofitColumns();
    } else if (input.width) {
      range.format.columnWidth = input.width;
    }

    await ctx.sync();
    return { success: true, message: `Set column width for ${input.sheet}!${input.range}` };
  });
}

async function setRowHeight(input) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(input.sheet);
    const range = sheet.getRange(input.range);

    if (input.autoFit) {
      range.format.autofitRows();
    } else if (input.height) {
      range.format.rowHeight = input.height;
    }

    await ctx.sync();
    return { success: true, message: `Set row height for ${input.sheet}!${input.range}` };
  });
}

async function toggleGridlines(input) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(input.sheet);
    sheet.showGridlines = input.show;
    await ctx.sync();
    return { success: true, message: `Gridlines ${input.show ? 'shown' : 'hidden'} on ${input.sheet}` };
  });
}

async function setPrintArea(input) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(input.sheet);
    sheet.pageLayout.printArea = sheet.getRange(input.range);
    await ctx.sync();
    return { success: true, message: `Print area set to ${input.sheet}!${input.range}` };
  });
}

async function addWorksheet(input) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.add(input.name);
    sheet.activate();
    await ctx.sync();
    return { success: true, message: `Created worksheet "${input.name}"` };
  });
}

async function traceFormula(input) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(input.sheet);
    const cell = sheet.getRange(input.cell);
    cell.load('formulas,values,address');
    await ctx.sync();

    const formula = cell.formulas[0][0];
    const value = cell.values[0][0];

    const result = {
      success: true,
      cell: input.cell,
      formula: formula,
      value: value,
      isError: typeof value === 'string' && value.startsWith('#'),
      precedents: []
    };

    // Parse formula to extract cell references
    if (typeof formula === 'string' && formula.startsWith('=')) {
      const refPattern = /(?:'?([^'!]+)'?!)?(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)/g;
      let match;
      const refs = new Set();

      while ((match = refPattern.exec(formula)) !== null) {
        const sheetRef = match[1] || input.sheet;
        const cellRef = match[2];
        refs.add(`${sheetRef}!${cellRef}`);
      }

      // Read each precedent cell
      for (const ref of refs) {
        try {
          const parts = ref.split('!');
          const refSheet = ctx.workbook.worksheets.getItem(parts[0]);
          const refRange = refSheet.getRange(parts[1]);
          refRange.load('values,formulas');
          await ctx.sync();

          const refValue = refRange.values[0][0];
          const refFormula = refRange.formulas[0][0];
          result.precedents.push({
            address: ref,
            formula: refFormula,
            value: refValue,
            isError: typeof refValue === 'string' && refValue.startsWith('#')
          });
        } catch (e) {
          result.precedents.push({ address: ref, error: 'Could not read cell' });
        }
      }
    }

    return result;
  });
}

async function findErrors(input) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(input.sheet);

    let range;
    if (input.range) {
      range = sheet.getRange(input.range);
    } else {
      range = sheet.getUsedRange();
    }

    range.load('values,formulas,address,rowCount,columnCount');
    await ctx.sync();

    const errorTypes = ['#REF!', '#VALUE!', '#N/A', '#DIV/0!', '#NAME?', '#NULL!', '#NUM!'];
    const errors = [];

    // Parse the starting address
    const baseAddr = range.address.replace(/^.*!/, '');
    const startMatch = baseAddr.match(/\$?([A-Z]+)\$?(\d+)/);
    if (!startMatch) {
      return { success: true, errors: [], message: 'Could not parse range address' };
    }

    const startColStr = startMatch[1];
    const startRow = parseInt(startMatch[2], 10);

    // Convert column letters to index
    let startCol = 0;
    for (let i = 0; i < startColStr.length; i++) {
      startCol = startCol * 26 + (startColStr.charCodeAt(i) - 64);
    }
    startCol--; // 0-based

    for (let r = 0; r < range.values.length; r++) {
      for (let c = 0; c < range.values[r].length; c++) {
        const val = range.values[r][c];
        if (typeof val === 'string' && errorTypes.includes(val)) {
          // Convert back to cell address
          let colIdx = startCol + c + 1;
          let colStr = '';
          while (colIdx > 0) {
            colIdx--;
            colStr = String.fromCharCode(65 + (colIdx % 26)) + colStr;
            colIdx = Math.floor(colIdx / 26);
          }
          const cellAddr = colStr + (startRow + r);

          errors.push({
            address: cellAddr,
            error: val,
            formula: range.formulas[r][c]
          });
        }
      }
    }

    return {
      success: true,
      errors,
      totalScanned: range.rowCount * range.columnCount,
      message: errors.length === 0
        ? 'No errors found'
        : `Found ${errors.length} error${errors.length === 1 ? '' : 's'}`
    };
  });
}

async function editChart(input) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(input.sheet);
    let chart;

    if (input.chartName) {
      chart = sheet.charts.getItem(input.chartName);
    } else if (input.chartIndex !== undefined) {
      chart = sheet.charts.getItemAt(input.chartIndex);
    } else {
      // Default to first chart
      chart = sheet.charts.getItemAt(0);
    }

    if (input.title !== undefined) {
      chart.title.text = input.title;
      chart.title.visible = true;
    }

    if (input.showLegend !== undefined) {
      chart.legend.visible = input.showLegend;
    }

    if (input.legendPosition) {
      chart.legend.position = input.legendPosition;
    }

    if (input.valueAxisTitle) {
      chart.axes.valueAxis.title.text = input.valueAxisTitle;
      chart.axes.valueAxis.title.visible = true;
    }

    if (input.categoryAxisTitle) {
      chart.axes.categoryAxis.title.text = input.categoryAxisTitle;
      chart.axes.categoryAxis.title.visible = true;
    }

    if (input.dataRange) {
      chart.setData(sheet.getRange(input.dataRange));
    }

    await ctx.sync();
    return { success: true, message: `Edited chart on ${input.sheet}` };
  });
}

async function createPivotTable(input) {
  return Excel.run(async (ctx) => {
    const sourceSheet = ctx.workbook.worksheets.getItem(input.sourceSheet);
    const sourceRange = sourceSheet.getRange(input.sourceRange);
    const destSheet = ctx.workbook.worksheets.getItem(input.destinationSheet);
    const destCell = destSheet.getRange(input.destinationCell);

    const pivotTable = ctx.workbook.pivotTables.add(input.name, sourceRange, destCell);
    await ctx.sync();

    // Add row fields
    if (input.rows) {
      for (const fieldName of input.rows) {
        const field = pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem(fieldName));
        await ctx.sync();
      }
    }

    // Add column fields
    if (input.columns) {
      for (const fieldName of input.columns) {
        pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem(fieldName));
        await ctx.sync();
      }
    }

    // Add value fields
    if (input.values) {
      for (const valConfig of input.values) {
        const dataField = pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem(valConfig.field));
        await ctx.sync();
        if (valConfig.summarizeBy) {
          dataField.summarizeBy = valConfig.summarizeBy;
          await ctx.sync();
        }
      }
    }

    // Add filter fields
    if (input.filters) {
      for (const fieldName of input.filters) {
        pivotTable.filterHierarchies.add(pivotTable.hierarchies.getItem(fieldName));
        await ctx.sync();
      }
    }

    return { success: true, message: `Created pivot table "${input.name}" on ${input.destinationSheet}` };
  });
}

async function refreshPivotTable(input) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(input.sheet);
    const pivotTable = sheet.pivotTables.getItem(input.name);
    pivotTable.refresh();
    await ctx.sync();
    return { success: true, message: `Refreshed pivot table "${input.name}"` };
  });
}
