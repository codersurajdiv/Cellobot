const express = require('express');
const router = express.Router();
const claudeService = require('../services/claude');
const openaiService = require('../services/openai');

router.post('/', async (req, res) => {
  try {
    const { message, model, context } = req.body;

    if (!message || typeof message !== 'string') {
      return res.status(400).json({ error: 'Message is required' });
    }

    const isExplanation = message.trim().startsWith('=');
    const contextStr = context ? JSON.stringify(context, null, 2) : 'No context provided.';

    let systemPrompt;
    if (isExplanation) {
      systemPrompt = `You are an Excel tutor. Explain this Excel formula clearly for a beginner. Be concise and helpful.

Context from the spreadsheet:
${contextStr}

Return only the explanation, no preamble.`;
    } else {
      systemPrompt = `You are an expert Excel formula generator. Return only a valid Excel formula. No explanation, no markdown, no code blocks. Start with =.

Context from the spreadsheet:
${contextStr}

If the user asks for something that cannot be done as a single formula, return a formula that best approximates their request.`;
    }

    const modelChoice = (model || 'claude').toLowerCase();
    let response;

    if (modelChoice === 'openai') {
      response = await openaiService.chat(systemPrompt, message);
    } else {
      response = await claudeService.chat(systemPrompt, message);
    }

    res.json({ response });
  } catch (err) {
    console.error('Chat error:', err);
    res.status(500).json({
      error: err.message || 'Failed to get AI response'
    });
  }
});

module.exports = router;
