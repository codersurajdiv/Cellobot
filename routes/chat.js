const express = require('express');
const router = express.Router();
const { runToolLoop } = require('../services/toolLoop');
const { resolveModel } = require('../utils/modelResolver');

/**
 * Build the system prompt from context.
 */
function buildSystemPrompt(context) {
  const contextStr = context ? JSON.stringify(context, null, 2) : 'No context provided.';

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

Context from the spreadsheet:
${contextStr}`;
}

/**
 * POST /chat
 *
 * Supports two modes:
 * 1. Tool-use mode (default): LLM can call tools, frontend executes them
 * 2. Continue mode: frontend sends tool results, LLM continues
 *
 * Request body:
 * - messages: Array of conversation messages
 * - message: (legacy) Single message string
 * - model: 'claude' or 'openai'
 * - context: Workbook context object
 * - toolResults: (optional) Array of tool execution results from frontend
 * - pendingMessages: (optional) Updated messages array including assistant tool calls
 */
router.post('/', async (req, res) => {
  try {
    const { message, messages, model, context, toolResults, pendingMessages } = req.body;

    const { provider, modelId } = resolveModel(model);

    // --- Continue mode: tool results coming back from frontend ---
    if (toolResults && pendingMessages) {
      const systemPrompt = buildSystemPrompt(context);

      // Append tool results to the message history
      let updatedMessages = [...pendingMessages];

      if (provider === 'openai') {
        // OpenAI format: each tool result is a separate message
        for (const result of toolResults) {
          updatedMessages.push({
            role: 'tool',
            tool_call_id: result.id,
            content: JSON.stringify(result.output)
          });
        }
      } else {
        // Anthropic format: tool results in a single user message
        updatedMessages.push({
          role: 'user',
          content: toolResults.map(result => ({
            type: 'tool_result',
            tool_use_id: result.id,
            content: JSON.stringify(result.output)
          }))
        });
      }

      const loopResult = await runToolLoop({
        systemPrompt,
        messages: updatedMessages,
        provider,
        modelId
      });

      return res.json(loopResult);
    }

    // --- Initial request mode ---
    let conversationMessages;
    if (messages && Array.isArray(messages) && messages.length > 0) {
      conversationMessages = messages;
    } else if (message && typeof message === 'string') {
      conversationMessages = [{ role: 'user', content: message }];
    } else {
      return res.status(400).json({ error: 'Message or messages array is required' });
    }

    const systemPrompt = buildSystemPrompt(context);

    const result = await runToolLoop({
      systemPrompt,
      messages: conversationMessages,
      provider,
      modelId
    });

    res.json(result);
  } catch (err) {
    console.error('Chat error:', err);
    res.status(500).json({
      error: err.message || 'Failed to get AI response'
    });
  }
});

module.exports = router;
