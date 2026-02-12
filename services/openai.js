const OpenAI = require('openai');

async function chat(systemPrompt, userMessage) {
  if (!process.env.OPENAI_API_KEY) {
    throw new Error('OPENAI_API_KEY is not configured');
  }

  const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

  const completion = await openai.chat.completions.create({
    model: 'gpt-4o-mini',
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: userMessage }
    ],
    max_tokens: 1024
  });

  const content = completion.choices[0]?.message?.content;
  return content ? content.trim() : '';
}

module.exports = { chat };
