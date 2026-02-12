/**
 * Tool loop service: manages multi-turn tool-use conversations.
 */

const claudeService = require('./claude');
const openaiService = require('./openai');

const MAX_TOOL_ROUNDS = 10;

/**
 * Start or continue a tool-use conversation.
 * @param {Object} params
 * @param {string} params.systemPrompt
 * @param {Array} params.messages
 * @param {string} params.provider - 'claude' or 'openai'
 * @param {string} params.modelId - Specific model ID
 * @returns {Object} { type: 'text'|'tool_calls', content, toolCalls, messages }
 */
async function runToolLoop(params) {
  const { systemPrompt, messages, provider, modelId } = params;

  if (provider === 'openai') {
    return await runOpenAIToolLoop(systemPrompt, messages, modelId);
  } else {
    return await runClaudeToolLoop(systemPrompt, messages, modelId);
  }
}

async function runClaudeToolLoop(systemPrompt, messages, modelId) {
  const result = await claudeService.chatWithTools(systemPrompt, messages, modelId);

  if (result.stopReason === 'tool_use') {
    return {
      type: 'tool_calls',
      toolCalls: result.toolCalls,
      text: result.text || '',
      messages: [
        ...messages,
        { role: 'assistant', content: result.rawContent }
      ]
    };
  }

  return {
    type: 'text',
    text: result.text,
    messages: messages
  };
}

async function runOpenAIToolLoop(systemPrompt, messages, modelId) {
  const result = await openaiService.chatWithTools(systemPrompt, messages, modelId);

  if (result.toolCalls && result.toolCalls.length > 0) {
    return {
      type: 'tool_calls',
      toolCalls: result.toolCalls,
      text: result.text || '',
      messages: [
        ...messages,
        result.rawMessage
      ]
    };
  }

  return {
    type: 'text',
    text: result.text,
    messages: messages
  };
}

module.exports = { runToolLoop, MAX_TOOL_ROUNDS };
