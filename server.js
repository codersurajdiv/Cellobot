require('dotenv').config();
const express = require('express');
const cors = require('cors');
const path = require('path');
const chatRouter = require('./routes/chat');
const streamRouter = require('./routes/stream');

const app = express();
const PORT = process.env.PORT || 3000;

// Updated CORS for production - allow Office.com domains
app.use(cors({
  origin: [
    'https://localhost:3000',
    'https://localhost:3001',
    'https://127.0.0.1:3000',
    'https://127.0.0.1:3001',
    /^https:\/\/.*\.officeapps\.live\.com$/,
    /^https:\/\/.*\.office\.com$/,
    /^https:\/\/.*\.microsoft\.com$/,
    'null'
  ],
  credentials: true
}));

// Increase JSON body limit for large workbook contexts
app.use(express.json({ limit: '5mb' }));

// Simple in-memory rate limiter (no extra dependency)
const rateLimitMap = new Map();
const RATE_LIMIT_WINDOW = 60 * 1000; // 1 minute
const RATE_LIMIT_MAX = 30; // 30 requests per minute

function rateLimit(req, res, next) {
  const ip = req.ip || req.connection.remoteAddress || 'unknown';
  const now = Date.now();

  if (!rateLimitMap.has(ip)) {
    rateLimitMap.set(ip, { count: 1, resetAt: now + RATE_LIMIT_WINDOW });
    return next();
  }

  const entry = rateLimitMap.get(ip);
  if (now > entry.resetAt) {
    entry.count = 1;
    entry.resetAt = now + RATE_LIMIT_WINDOW;
    return next();
  }

  entry.count++;
  if (entry.count > RATE_LIMIT_MAX) {
    return res.status(429).json({ error: 'Too many requests. Please wait a moment and try again.' });
  }

  next();
}

// Clean up stale rate limit entries every 5 minutes
setInterval(() => {
  const now = Date.now();
  for (const [key, val] of rateLimitMap) {
    if (now > val.resetAt) rateLimitMap.delete(key);
  }
}, 5 * 60 * 1000);

// Apply rate limiting to AI endpoints
app.use('/chat', rateLimit, chatRouter);
app.use('/stream', rateLimit, streamRouter);

// Serve static files - check multiple locations for dev/prod compatibility
const addinPath = path.join(__dirname, '../add-in/src');
const publicPath = path.join(__dirname, 'public');
const fs = require('fs');

if (fs.existsSync(publicPath)) {
  app.use(express.static(publicPath));
} else if (fs.existsSync(addinPath)) {
  app.use(express.static(addinPath));
}

app.get('/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// Minimal PNG icons for Office add-in
const minimalPng = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==', 'base64');
app.get('/taskpane/assets/icon-16.png', (req, res) => {
  res.type('image/png').send(minimalPng);
});
app.get('/taskpane/assets/icon-32.png', (req, res) => {
  res.type('image/png').send(minimalPng);
});
app.get('/taskpane/assets/icon-64.png', (req, res) => {
  res.type('image/png').send(minimalPng);
});
app.get('/taskpane/assets/icon-80.png', (req, res) => {
  res.type('image/png').send(minimalPng);
});

// Global error handler (error boundary)
app.use((err, req, res, _next) => {
  console.error('Unhandled error:', err);
  res.status(500).json({
    error: 'An unexpected error occurred. Please try again.',
    details: process.env.NODE_ENV === 'development' ? err.message : undefined
  });
});

// Use HTTP in production - Railway handles SSL termination
app.listen(PORT, '0.0.0.0', () => {
  console.log(`CelloBot backend running on port ${PORT}`);
  console.log(`Environment: ${process.env.NODE_ENV || 'development'}`);
});
