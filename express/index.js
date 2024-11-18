// server.js
require('dotenv').config();
const fs = require("fs");
const express = require('express');
const { Anthropic } = require('@anthropic-ai/sdk');
const app = express();
// const https = require("https");
const cors = require("cors");

const PORT = 5002;
app.use(cors({ origin: "*" }));

  
app.use(express.json());
// Load SSL certificate and key files
// const key = fs.readFileSync("./ssl/key.pem");
// const cert = fs.readFileSync("./ssl/cert.pem");

// Initialize the Anthropic client


const anthropicClient = new Anthropic({ apiKey: process.env.Anthropic});
app.get('/', (req, res) => {
    res.send('Hello World');
  });   
app.post('/api/anthropic', async (req, res) => {
  const { message } = req.body;

  try {
    // Use the Anthropic SDK to send the message
    const response = await anthropicClient.messages.create({
        model: "claude-3-5-sonnet-20241022", // replace with your desired model
        max_tokens: 1024,
      messages: [{role: "user", content: message}],
    });
    console.log(response);
    res.json(response);
  } catch (error) {
    console.error('Error:', error);
    res.status(500).json({ error: 'Failed to fetch data from Anthropic API' });
  }
});

// const httpsServer = https.createServer({ key, cert }, app);


app.listen(PORT, () => {
    console.log("HTTPS server running on http://localhost:5000");
  });