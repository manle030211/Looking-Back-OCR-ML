const express = require('express');
const path = require('path');
const fs = require('fs');

const app = express();
const port = parseInt(process.env.PORT || '3000', 10);
const distDir = path.join(__dirname, 'dist');

// Serve static assets from the Angular build directory
app.use(express.static(distDir));

// Fallback to index.html for Angular SPA routing
app.use((req, res) => {
  const indexPath = path.join(distDir, 'index.html');
  if (fs.existsSync(indexPath)) {
    res.sendFile(indexPath);
  } else {
    res.status(404).send('Application build not found. Please build the project.');
  }
});

app.listen(port, '0.0.0.0', () => {
  console.log(`Application server running on http://0.0.0.0:${port}`);
});
