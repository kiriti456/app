const express = require('express');
const path = require('path');

const app = express();
const port = process.env.PORT || 8080;

// Serve the Spring Boot JAR file
app.use(express.static(path.join(__dirname, '../target')));

// Handle health check
app.get('/health', (req, res) => {
  res.status(200).send('OK');
});

app.listen(port, () => {
  console.log(`Server started on port ${port}`);
});
