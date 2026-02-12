const Anthropic = require('@anthropic-ai/sdk');

async function chat(systemPrompt, userMessage) {
  if (!process.env.ANTHROPIC_API_KEY) {
    throw new Error('ANTHROPIC_API_KEY is not configured');
  }

  const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

  const message = await anthropic.messages.create({
    model: 'claude-3-5-sonnet-20241022',
    max_tokens: 1024,
    system: systemPrompt,
    messages: [
      { role: 'user', content: userMessage }
    ]
  });

  const textBlock = message.content.find(block => block.type === 'text');
  return textBlock ? textBlock.text.trim() : '';
}

module.exports = { chat };
