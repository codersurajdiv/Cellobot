const OpenAI = require('openai');
const { openaiTools } = require('../tools/definitions');

const DEFAULT_MODEL = 'gpt-4o-mini';

function getClient() {
  if (!process.env.OPENAI_API_KEY) {
    throw new Error('OPENAI_API_KEY is not configured');
  }
  return new OpenAI({ apiKey: process.env.OPENAI_API_KEY });
}

/**
 * Simple chat without tools.
 */
async function chat(systemPrompt, messages, modelId) {
  const openai = getClient();

  const completion = await openai.chat.completions.create({
    model: modelId || DEFAULT_MODEL,
    messages: [
      { role: 'system', content: systemPrompt },
      ...messages
    ],
    max_tokens: 4096
  });

  const content = completion.choices[0]?.message?.content;
  return content ? content.trim() : '';
}

/**
 * Chat with tool-use support.
 */
async function chatWithTools(systemPrompt, messages, modelId) {
  const openai = getClient();

  const completion = await openai.chat.completions.create({
    model: modelId || DEFAULT_MODEL,
    messages: [
      { role: 'system', content: systemPrompt },
      ...messages
    ],
    max_tokens: 4096,
    tools: openaiTools,
    tool_choice: 'auto'
  });

  const message = completion.choices[0]?.message;
  const text = message?.content || '';
  const rawToolCalls = message?.tool_calls || [];

  const toolCalls = rawToolCalls.map(tc => ({
    id: tc.id,
    name: tc.function.name,
    input: JSON.parse(tc.function.arguments)
  }));

  return {
    text: text.trim(),
    toolCalls,
    rawMessage: message
  };
}

module.exports = { chat, chatWithTools, DEFAULT_MODEL };
