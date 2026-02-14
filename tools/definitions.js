/**
 * Tool definitions for LLM function calling.
 * These schemas tell the LLM what tools are available and how to call them.
 * Tools are executed client-side against the Excel JS API.
 */

// Anthropic tool format
const anthropicTools = [
  {
    name: 'write_cells',
    description: 'Write values or formulas to a range of cells in the workbook. Use this to insert data, formulas, or update existing cells.',
    input_schema: {
      type: 'object',
      properties: {
        sheet: {
          type: 'string',
          description: 'The worksheet name to write to (e.g. "Sheet1")'
        },
        range: {
          type: 'string',
          description: 'The cell range address (e.g. "A1", "A1:B10", "C3:C3")'
        },
        values: {
          type: 'array',
          items: { type: 'array', items: { type: 'string' } },
          description: '2D array of values to write. Each inner array is a row. Use this for plain values.'
        },
        formulas: {
          type: 'array',
          items: { type: 'array', items: { type: 'string' } },
          description: '2D array of formulas to write. Each inner array is a row. Formulas must start with "=". If provided, takes precedence over values.'
        }
      },
      required: ['sheet', 'range']
    }
  },
  {
    name: 'read_range',
    description: 'Read the values and formulas from a specific range of cells. Use this to inspect cell contents before making changes or to gather additional context.',
    input_schema: {
      type: 'object',
      properties: {
        sheet: {
          type: 'string',
          description: 'The worksheet name to read from'
        },
        range: {
          type: 'string',
          description: 'The cell range address to read (e.g. "A1:D20")'
        }
      },
      required: ['sheet', 'range']
    }
  },
  {
    name: 'get_workbook_info',
    description: 'Get detailed information about the workbook structure: all sheet names, used ranges, tables, and named ranges. Use this when you need to understand the overall workbook layout.',
    input_schema: {
      type: 'object',
      properties: {}
    }
  },
  {
    name: 'format_cells',
    description: 'Apply formatting to a range of cells (number format, font, fill, borders, alignment).',
    input_schema: {
      type: 'object',
      properties: {
        sheet: {
          type: 'string',
          description: 'The worksheet name'
        },
        range: {
          type: 'string',
          description: 'The cell range address to format'
        },
        numberFormat: {
          type: 'string',
          description: 'Number format string (e.g. "#,##0.00", "0%", "yyyy-mm-dd", "$#,##0.00")'
        },
        bold: {
          type: 'boolean',
          description: 'Set font to bold'
        },
        italic: {
          type: 'boolean',
          description: 'Set font to italic'
        },
        fontSize: {
          type: 'number',
          description: 'Font size in points'
        },
        fontColor: {
          type: 'string',
          description: 'Font color as hex (e.g. "#FF0000" for red)'
        },
        fillColor: {
          type: 'string',
          description: 'Cell background fill color as hex (e.g. "#FFFF00" for yellow)'
        },
        horizontalAlignment: {
          type: 'string',
          enum: ['General', 'Left', 'Center', 'Right', 'Fill', 'Justify', 'CenterAcrossSelection', 'Distributed'],
          description: 'Horizontal text alignment'
        },
        verticalAlignment: {
          type: 'string',
          enum: ['Top', 'Center', 'Bottom', 'Justify', 'Distributed'],
          description: 'Vertical text alignment'
        },
        wrapText: {
          type: 'boolean',
          description: 'Whether to wrap text in cells'
        },
        borderStyle: {
          type: 'string',
          enum: ['None', 'Thin', 'Medium', 'Thick', 'Double'],
          description: 'Border style to apply around the range'
        },
        borderColor: {
          type: 'string',
          description: 'Border color as hex'
        },
        merge: {
          type: 'boolean',
          description: 'Whether to merge the cells in the range'
        }
      },
      required: ['sheet', 'range']
    }
  },
  {
    name: 'create_chart',
    description: 'Create a new chart from a data range.',
    input_schema: {
      type: 'object',
      properties: {
        sheet: {
          type: 'string',
          description: 'The worksheet name where the chart will be created'
        },
        dataRange: {
          type: 'string',
          description: 'The data range for the chart (e.g. "A1:D10")'
        },
        chartType: {
          type: 'string',
          enum: ['ColumnClustered', 'ColumnStacked', 'BarClustered', 'BarStacked', 'Line', 'LineMarkers', 'Pie', 'Area', 'AreaStacked', 'XYScatter', 'XYScatterLines', 'Radar', 'Doughnut', 'Bubble'],
          description: 'The type of chart to create'
        },
        title: {
          type: 'string',
          description: 'Chart title text'
        },
        seriesBy: {
          type: 'string',
          enum: ['Auto', 'Columns', 'Rows'],
          description: 'Whether data series are in rows or columns. Default: Auto'
        },
        position: {
          type: 'string',
          description: 'Cell address where the chart top-left corner should be placed (e.g. "F1")'
        }
      },
      required: ['sheet', 'dataRange', 'chartType']
    }
  },
  {
    name: 'sort_range',
    description: 'Sort data in a range or table by one or more columns.',
    input_schema: {
      type: 'object',
      properties: {
        sheet: {
          type: 'string',
          description: 'The worksheet name'
        },
        range: {
          type: 'string',
          description: 'The range to sort (e.g. "A1:D20")'
        },
        sortBy: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              column: { type: 'integer', description: 'Zero-based column index to sort by' },
              ascending: { type: 'boolean', description: 'Sort ascending (true) or descending (false)' }
            },
            required: ['column', 'ascending']
          },
          description: 'Array of sort criteria'
        },
        hasHeaders: {
          type: 'boolean',
          description: 'Whether the first row contains headers. Default: true'
        }
      },
      required: ['sheet', 'range', 'sortBy']
    }
  },
  {
    name: 'filter_data',
    description: 'Apply auto-filter to a range or table column.',
    input_schema: {
      type: 'object',
      properties: {
        sheet: {
          type: 'string',
          description: 'The worksheet name'
        },
        range: {
          type: 'string',
          description: 'The data range to filter (e.g. "A1:D20")'
        },
        column: {
          type: 'integer',
          description: 'Zero-based column index to filter on'
        },
        values: {
          type: 'array',
          items: { type: 'string' },
          description: 'Array of values to show (hide all others)'
        }
      },
      required: ['sheet', 'range', 'column', 'values']
    }
  },
  {
    name: 'set_data_validation',
    description: 'Set data validation rules on a range (dropdowns, numeric constraints, etc.).',
    input_schema: {
      type: 'object',
      properties: {
        sheet: {
          type: 'string',
          description: 'The worksheet name'
        },
        range: {
          type: 'string',
          description: 'The cell range to apply validation to'
        },
        type: {
          type: 'string',
          enum: ['list', 'wholeNumber', 'decimal', 'date', 'textLength', 'custom'],
          description: 'The type of validation to apply'
        },
        listSource: {
          type: 'string',
          description: 'For list validation: comma-separated values (e.g. "Active,Pending,Closed")'
        },
        operator: {
          type: 'string',
          enum: ['Between', 'NotBetween', 'EqualTo', 'NotEqualTo', 'GreaterThan', 'LessThan', 'GreaterThanOrEqualTo', 'LessThanOrEqualTo'],
          description: 'Comparison operator for numeric/date/textLength validation'
        },
        formula1: {
          type: 'string',
          description: 'First formula/value for the validation rule'
        },
        formula2: {
          type: 'string',
          description: 'Second formula/value (for Between/NotBetween operators)'
        },
        customFormula: {
          type: 'string',
          description: 'For custom validation: a formula that evaluates to TRUE/FALSE'
        },
        errorMessage: {
          type: 'string',
          description: 'Error message shown when validation fails'
        }
      },
      required: ['sheet', 'range', 'type']
    }
  },
  {
    name: 'add_conditional_format',
    description: 'Apply conditional formatting rules to a range.',
    input_schema: {
      type: 'object',
      properties: {
        sheet: {
          type: 'string',
          description: 'The worksheet name'
        },
        range: {
          type: 'string',
          description: 'The cell range to format'
        },
        ruleType: {
          type: 'string',
          enum: ['cellValue', 'colorScale', 'dataBar', 'iconSet'],
          description: 'The type of conditional format'
        },
        operator: {
          type: 'string',
          enum: ['GreaterThan', 'LessThan', 'Between', 'EqualTo', 'NotEqualTo', 'GreaterThanOrEqual', 'LessThanOrEqual'],
          description: 'For cellValue rules: the comparison operator'
        },
        formula1: {
          type: 'string',
          description: 'First formula/value for the condition'
        },
        formula2: {
          type: 'string',
          description: 'Second formula/value (for Between operator)'
        },
        fontColor: {
          type: 'string',
          description: 'Font color when condition is met (hex)'
        },
        fillColor: {
          type: 'string',
          description: 'Fill color when condition is met (hex)'
        }
      },
      required: ['sheet', 'range', 'ruleType']
    }
  },
  {
    name: 'set_column_width',
    description: 'Set the width of one or more columns.',
    input_schema: {
      type: 'object',
      properties: {
        sheet: {
          type: 'string',
          description: 'The worksheet name'
        },
        range: {
          type: 'string',
          description: 'Column range (e.g. "A:C" or "B:B")'
        },
        width: {
          type: 'number',
          description: 'Column width in points'
        },
        autoFit: {
          type: 'boolean',
          description: 'Auto-fit column width to content. If true, width is ignored.'
        }
      },
      required: ['sheet', 'range']
    }
  },
  {
    name: 'set_row_height',
    description: 'Set the height of one or more rows.',
    input_schema: {
      type: 'object',
      properties: {
        sheet: {
          type: 'string',
          description: 'The worksheet name'
        },
        range: {
          type: 'string',
          description: 'Row range (e.g. "1:1" or "1:5")'
        },
        height: {
          type: 'number',
          description: 'Row height in points'
        },
        autoFit: {
          type: 'boolean',
          description: 'Auto-fit row height to content. If true, height is ignored.'
        }
      },
      required: ['sheet', 'range']
    }
  },
  {
    name: 'toggle_gridlines',
    description: 'Show or hide gridlines on a worksheet.',
    input_schema: {
      type: 'object',
      properties: {
        sheet: {
          type: 'string',
          description: 'The worksheet name'
        },
        show: {
          type: 'boolean',
          description: 'Whether to show (true) or hide (false) gridlines'
        }
      },
      required: ['sheet', 'show']
    }
  },
  {
    name: 'set_print_area',
    description: 'Set the print area for a worksheet.',
    input_schema: {
      type: 'object',
      properties: {
        sheet: {
          type: 'string',
          description: 'The worksheet name'
        },
        range: {
          type: 'string',
          description: 'The range to set as print area (e.g. "A1:F20")'
        }
      },
      required: ['sheet', 'range']
    }
  },
  {
    name: 'add_worksheet',
    description: 'Add a new worksheet to the workbook.',
    input_schema: {
      type: 'object',
      properties: {
        name: {
          type: 'string',
          description: 'Name for the new worksheet'
        }
      },
      required: ['name']
    }
  },
  {
    name: 'trace_formula',
    description: 'Trace a formula\'s dependencies. Returns the formula, its direct precedent cells (cells it references), and whether any cells in the chain have errors. Use this to debug formula errors like #REF!, #VALUE!, #N/A, etc.',
    input_schema: {
      type: 'object',
      properties: {
        sheet: {
          type: 'string',
          description: 'The worksheet name'
        },
        cell: {
          type: 'string',
          description: 'The cell address to trace (e.g. "B5")'
        }
      },
      required: ['sheet', 'cell']
    }
  },
  {
    name: 'find_errors',
    description: 'Scan a range for cells containing errors (#REF!, #VALUE!, #N/A, #DIV/0!, #NAME?, #NULL!, #NUM!). Returns addresses and error types of all error cells found.',
    input_schema: {
      type: 'object',
      properties: {
        sheet: {
          type: 'string',
          description: 'The worksheet name'
        },
        range: {
          type: 'string',
          description: 'The range to scan (e.g. "A1:Z100"). If omitted, scans the used range.'
        }
      },
      required: ['sheet']
    }
  },
  {
    name: 'edit_chart',
    description: 'Modify an existing chart\'s properties (title, axes, legend, data range).',
    input_schema: {
      type: 'object',
      properties: {
        sheet: {
          type: 'string',
          description: 'The worksheet containing the chart'
        },
        chartName: {
          type: 'string',
          description: 'The name of the chart to edit (use get_workbook_info to find chart names)'
        },
        chartIndex: {
          type: 'integer',
          description: 'Zero-based index of the chart on the sheet (alternative to chartName)'
        },
        title: {
          type: 'string',
          description: 'New chart title'
        },
        showLegend: {
          type: 'boolean',
          description: 'Show or hide the legend'
        },
        legendPosition: {
          type: 'string',
          enum: ['Top', 'Bottom', 'Left', 'Right', 'Invalid'],
          description: 'Position of the legend'
        },
        valueAxisTitle: {
          type: 'string',
          description: 'Title for the value (Y) axis'
        },
        categoryAxisTitle: {
          type: 'string',
          description: 'Title for the category (X) axis'
        },
        dataRange: {
          type: 'string',
          description: 'New data range for the chart (e.g. "A1:D10")'
        }
      },
      required: ['sheet']
    }
  },
  {
    name: 'create_pivot_table',
    description: 'Create a pivot table from a data range.',
    input_schema: {
      type: 'object',
      properties: {
        sourceSheet: {
          type: 'string',
          description: 'Sheet containing the source data'
        },
        sourceRange: {
          type: 'string',
          description: 'The data range for the pivot table (e.g. "A1:E100")'
        },
        destinationSheet: {
          type: 'string',
          description: 'Sheet where the pivot table will be placed'
        },
        destinationCell: {
          type: 'string',
          description: 'Top-left cell for the pivot table (e.g. "A1")'
        },
        name: {
          type: 'string',
          description: 'Name for the pivot table'
        },
        rows: {
          type: 'array',
          items: { type: 'string' },
          description: 'Column names to use as row fields'
        },
        columns: {
          type: 'array',
          items: { type: 'string' },
          description: 'Column names to use as column fields'
        },
        values: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              field: { type: 'string', description: 'Column name for the value field' },
              summarizeBy: {
                type: 'string',
                enum: ['Sum', 'Count', 'Average', 'Max', 'Min', 'Product', 'CountNumbers'],
                description: 'Aggregation function. Default: Sum'
              }
            },
            required: ['field']
          },
          description: 'Value fields with aggregation settings'
        },
        filters: {
          type: 'array',
          items: { type: 'string' },
          description: 'Column names to use as filter fields'
        }
      },
      required: ['sourceSheet', 'sourceRange', 'destinationSheet', 'destinationCell', 'name']
    }
  },
  {
    name: 'refresh_pivot_table',
    description: 'Refresh an existing pivot table to reflect updated source data.',
    input_schema: {
      type: 'object',
      properties: {
        sheet: {
          type: 'string',
          description: 'Sheet containing the pivot table'
        },
        name: {
          type: 'string',
          description: 'Name of the pivot table to refresh'
        }
      },
      required: ['sheet', 'name']
    }
  }
];

// Convert to OpenAI tool format
const openaiTools = anthropicTools.map(tool => ({
  type: 'function',
  function: {
    name: tool.name,
    description: tool.description,
    parameters: tool.input_schema
  }
}));

module.exports = { anthropicTools, openaiTools };
