require('dotenv').config();
const express = require('express');
const cors = require('cors');
const path = require('path');
const chatRouter = require('./routes/chat');

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
app.use(express.json());

app.use('/chat', chatRouter);

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

// Use HTTP in production - Railway handles SSL termination
app.listen(PORT, '0.0.0.0', () => {
  console.log(`CelloBot backend running on port ${PORT}`);
  console.log(`Environment: ${process.env.NODE_ENV || 'development'}`);
});
