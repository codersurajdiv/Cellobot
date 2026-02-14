const express = require('express');
const router = express.Router();
const { runToolLoop } = require('../services/toolLoop');
const { resolveModel } = require('../utils/modelResolver');
const { buildSystemPrompt } = require('../utils/promptBuilder');

function getLastUserMessage(messages) {
  if (!messages || !Array.isArray(messages)) return '';
  const lastUser = messages.filter((m) => m.role === 'user').pop();
  if (!lastUser) return '';
  if (typeof lastUser.content === 'string') return lastUser.content;
  if (Array.isArray(lastUser.content)) {
    const textBlock = lastUser.content.find((c) => c.type === 'text');
    return textBlock?.text || '';
  }
  return '';
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
      const lastUserMsg = getLastUserMessage(pendingMessages);
      const systemPrompt = buildSystemPrompt(context, lastUserMsg);

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

    const lastUserMsg = getLastUserMessage(conversationMessages);
    const systemPrompt = buildSystemPrompt(context, lastUserMsg);

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
