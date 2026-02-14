const Anthropic = require('@anthropic-ai/sdk');
const { anthropicTools } = require('../tools/definitions');

const DEFAULT_MODEL = 'claude-sonnet-4-5-20250929';

function getClient() {
  if (!process.env.ANTHROPIC_API_KEY) {
    throw new Error('ANTHROPIC_API_KEY is not configured');
  }
  return new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
}

/**
 * Simple chat without tools.
 */
async function chat(systemPrompt, messages, modelId) {
  const anthropic = getClient();

  const response = await anthropic.messages.create({
    model: modelId || DEFAULT_MODEL,
    max_tokens: 4096,
    system: systemPrompt,
    messages: messages
  });

  const textBlock = response.content.find(block => block.type === 'text');
  return textBlock ? textBlock.text.trim() : '';
}

/**
 * Chat with tool-use support.
 */
async function chatWithTools(systemPrompt, messages, modelId) {
  const anthropic = getClient();

  const response = await anthropic.messages.create({
    model: modelId || DEFAULT_MODEL,
    max_tokens: 4096,
    system: systemPrompt,
    messages: messages,
    tools: anthropicTools
  });

  const textBlocks = response.content.filter(b => b.type === 'text');
  const toolBlocks = response.content.filter(b => b.type === 'tool_use');

  const text = textBlocks.map(b => b.text).join('\n').trim();
  const toolCalls = toolBlocks.map(b => ({
    id: b.id,
    name: b.name,
    input: b.input
  }));

  return {
    text,
    toolCalls,
    stopReason: response.stop_reason,
    rawContent: response.content
  };
}

module.exports = { chat, chatWithTools, DEFAULT_MODEL };
