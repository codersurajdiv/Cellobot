const express = require('express');
const router = express.Router();
const Anthropic = require('@anthropic-ai/sdk');
const OpenAI = require('openai');
const { anthropicTools, openaiTools } = require('../tools/definitions');
const { resolveModel } = require('../utils/modelResolver');

/**
 * Build the system prompt from context (shared with chat.js).
 */
function buildSystemPrompt(context) {
  const { pinnedRanges, ...autoContext } = context || {};
  const autoStr = Object.keys(autoContext).length > 0
    ? JSON.stringify(autoContext, null, 2)
    : 'No auto-detected context.';

  let pinnedStr = '';
  if (pinnedRanges && pinnedRanges.length > 0) {
    pinnedStr = `\n\nUser-pinned context (the user explicitly selected these ranges for you to reference):\n${JSON.stringify(pinnedRanges, null, 2)}`;
  }

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
${pinnedStr}`;
}

/**
 * POST /stream
 *
 * Server-Sent Events endpoint for streaming LLM responses.
 * Sends events:
 *   - text_delta: { text: "chunk" }
 *   - tool_calls: { toolCalls: [...], messages: [...] }
 *   - done: { text: "full response" }
 *   - error: { error: "message" }
 */
router.post('/', async (req, res) => {
  const { messages, model, context, toolResults, pendingMessages } = req.body;

  // Set up SSE headers
  res.writeHead(200, {
    'Content-Type': 'text/event-stream',
    'Cache-Control': 'no-cache',
    'Connection': 'keep-alive',
    'X-Accel-Buffering': 'no'
  });

  const sendEvent = (event, data) => {
    res.write(`event: ${event}\ndata: ${JSON.stringify(data)}\n\n`);
  };

  try {
    const { provider, modelId } = resolveModel(model);
    const systemPrompt = buildSystemPrompt(context);

    // Build final messages array
    let finalMessages;
    if (toolResults && pendingMessages) {
      finalMessages = [...pendingMessages];
      if (provider === 'openai') {
        for (const result of toolResults) {
          finalMessages.push({
            role: 'tool',
            tool_call_id: result.id,
            content: JSON.stringify(result.output)
          });
        }
      } else {
        finalMessages.push({
          role: 'user',
          content: toolResults.map(result => ({
            type: 'tool_result',
            tool_use_id: result.id,
            content: JSON.stringify(result.output)
          }))
        });
      }
    } else if (messages && Array.isArray(messages)) {
      finalMessages = messages;
    } else {
      sendEvent('error', { error: 'Messages array is required' });
      res.end();
      return;
    }

    if (provider === 'openai') {
      await streamOpenAI(systemPrompt, finalMessages, modelId, sendEvent, res);
    } else {
      await streamClaude(systemPrompt, finalMessages, modelId, sendEvent, res);
    }
  } catch (err) {
    console.error('Stream error:', err);
    sendEvent('error', { error: err.message || 'Streaming failed' });
    res.end();
  }
});

async function streamClaude(systemPrompt, messages, modelId, sendEvent, res) {
  if (!process.env.ANTHROPIC_API_KEY) {
    throw new Error('ANTHROPIC_API_KEY is not configured');
  }

  const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

  const stream = anthropic.messages.stream({
    model: modelId,
    max_tokens: 4096,
    system: systemPrompt,
    messages: messages,
    tools: anthropicTools
  });

  let fullText = '';
  let toolCalls = [];
  let rawContent = [];

  stream.on('text', (text) => {
    fullText += text;
    sendEvent('text_delta', { text });
  });

  stream.on('contentBlock', (block) => {
    rawContent.push(block);
    if (block.type === 'tool_use') {
      toolCalls.push({
        id: block.id,
        name: block.name,
        input: block.input
      });
    }
  });

  const finalMessage = await stream.finalMessage();

  if (finalMessage.stop_reason === 'tool_use' && toolCalls.length > 0) {
    // LLM wants tools — send structured event, not streamed text
    sendEvent('tool_calls', {
      toolCalls,
      text: fullText,
      messages: [...messages, { role: 'assistant', content: finalMessage.content }]
    });
  } else {
    sendEvent('done', { text: fullText });
  }

  res.end();
}

async function streamOpenAI(systemPrompt, messages, modelId, sendEvent, res) {
  if (!process.env.OPENAI_API_KEY) {
    throw new Error('OPENAI_API_KEY is not configured');
  }

  const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

  const stream = await openai.chat.completions.create({
    model: modelId,
    messages: [
      { role: 'system', content: systemPrompt },
      ...messages
    ],
    max_tokens: 4096,
    tools: openaiTools,
    tool_choice: 'auto',
    stream: true
  });

  let fullText = '';
  let toolCallsMap = {};
  let hasToolCalls = false;

  for await (const chunk of stream) {
    const delta = chunk.choices[0]?.delta;
    if (!delta) continue;

    // Text content
    if (delta.content) {
      fullText += delta.content;
      sendEvent('text_delta', { text: delta.content });
    }

    // Tool calls (accumulated across chunks)
    if (delta.tool_calls) {
      hasToolCalls = true;
      for (const tc of delta.tool_calls) {
        const idx = tc.index;
        if (!toolCallsMap[idx]) {
          toolCallsMap[idx] = { id: tc.id || '', name: '', arguments: '' };
        }
        if (tc.id) toolCallsMap[idx].id = tc.id;
        if (tc.function?.name) toolCallsMap[idx].name += tc.function.name;
        if (tc.function?.arguments) toolCallsMap[idx].arguments += tc.function.arguments;
      }
    }
  }

  if (hasToolCalls) {
    const toolCalls = Object.values(toolCallsMap).map(tc => ({
      id: tc.id,
      name: tc.name,
      input: JSON.parse(tc.arguments)
    }));

    // Build the raw assistant message for continuation
    const rawMessage = {
      role: 'assistant',
      content: fullText || null,
      tool_calls: Object.values(toolCallsMap).map(tc => ({
        id: tc.id,
        type: 'function',
        function: { name: tc.name, arguments: tc.arguments }
      }))
    };

    sendEvent('tool_calls', {
      toolCalls,
      text: fullText,
      messages: [...messages, rawMessage]
    });
  } else {
    sendEvent('done', { text: fullText });
  }

  res.end();
}

module.exports = router;
