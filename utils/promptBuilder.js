const { findRelevantSkills, formatSkillContext } = require('../services/skillRetriever');

/**
 * Build the system prompt from context and optional user message for RAG skill retrieval.
 */
function buildSystemPrompt(context, userMessage = '') {
  const { pinnedRanges, ...autoContext } = context || {};
  const autoStr = Object.keys(autoContext).length > 0
    ? JSON.stringify(autoContext, null, 2)
    : 'No auto-detected context.';

  let pinnedStr = '';
  if (pinnedRanges && pinnedRanges.length > 0) {
    pinnedStr = `\n\nUser-pinned context (the user explicitly selected these ranges for you to reference):\n${JSON.stringify(pinnedRanges, null, 2)}`;
  }

  const relevantSkills = findRelevantSkills(userMessage);
  const skillContext = formatSkillContext(relevantSkills);

  return `You are CelloBot, an expert AI assistant for Microsoft Excel. You help users with formulas, data analysis, formatting, charting, and all spreadsheet tasks.

You have access to tools that can read from and write to the user's Excel workbook. Use them when the user asks you to make changes, create formulas, format cells, build charts, sort/filter data, or perform any workbook operation.

When the user asks you to explain a formula, provide a clear explanation without using tools.
When the user asks you to create, modify, or analyze data, use the appropriate tools.

Important guidelines:
- Always specify the correct sheet name when using tools.
- When writing formulas, make sure they start with "=".
- When referencing cell addresses, use standard Excel notation (e.g. A1, B2:D10).
- If you need more context about the workbook, use the read_range or get_workbook_info tools.
- Explain what you're doing before and after making changes.
- If the user's request is ambiguous, ask for clarification.
- When referencing specific cells in your explanations, use double-bracket notation like [[Sheet1!A1]] or [[B2:D10]] so they become clickable citations.

Auto-detected context from the spreadsheet:
${autoStr}
${pinnedStr}${skillContext}`;
}

module.exports = { buildSystemPrompt };
