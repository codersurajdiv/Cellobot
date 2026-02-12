require('dotenv').config();
const express = require('express');
const cors = require('cors');
const path = require('path');
const https = require('https');
const chatRouter = require('./routes/chat');
const devCerts = require('office-addin-dev-certs');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors({
  origin: ['https://localhost:3000', 'https://localhost:3001', 'https://127.0.0.1:3000', 'https://127.0.0.1:3001', 'null'],
  credentials: true
}));
app.use(express.json());

app.use('/chat', chatRouter);
app.use(express.static(path.join(__dirname, '../add-in/src')));

app.get('/health', (req, res) => {
  res.json({ status: 'ok' });
});

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

async function start() {
  const options = await devCerts.getHttpsServerOptions();
  const server = https.createServer(options, app);
  server.listen(PORT, () => {
    console.log(`CelloBot backend running on https://localhost:${PORT}`);
  });
}

start().catch(err => {
  console.error('Failed to start server:', err);
  process.exit(1);
});
