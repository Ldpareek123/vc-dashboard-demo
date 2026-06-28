// Simple Node.js server to serve static HTML dashboard
const express = require('express');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 8080;

// Serve static files from the root directory
app.use(express.static(path.join(__dirname)));

// Serve index.html for all routes (SPA fallback)
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

// Start the server
app.listen(PORT, () => {
  console.log(`✓ SSPI Dashboard V4 is running on port ${PORT}`);
  console.log(`✓ Access it at: http://localhost:${PORT}`);
  console.log(`✓ Serving static files from: ${__dirname}`);
});
