// server.js
require('dotenv').config();
const fs = require("fs");
const express = require('express');
const { Anthropic } = require('@anthropic-ai/sdk');
const app = express();
const path = require("path");
// const https = require("https");
const cors = require("cors");
const multer = require("multer");
const xlsx = require("xlsx");

const PORT = 5002;
app.use(cors({ origin: "*" }));
const upload = multer({ dest: "uploads/" });

  
app.use(express.json());
// Load SSL certificate and key files
// const key = fs.readFileSync("./ssl/key.pem");
// const cert = fs.readFileSync("./ssl/cert.pem");

// Initialize the Anthropic client


const anthropicClient = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY});
app.get('/', (req, res) => {
    res.send('Hello World');
  });   

  app.get("/file/:filename", (req, res) => {
    const fileName = req.params.filename; // Get the filename from the URL
    const filePath = path.join(__dirname, fileName); // Build the full file path

  // Read the file and convert it to Base64
  fs.readFile(filePath, (err, data) => {
    if (err) {
        console.error("Error reading file:", err);
        return res.status(500).send("Failed to read the file.");
    }

    // Convert the file content to Base64
    const base64Data = data.toString("base64");

    // Send the Base64 string as a JSON response
    res.json({
        fileName: fileName,
        base64: base64Data,
    });
});
});



app.post('/api/anthropic/upload', upload.single("file"), async (req, res) => {
  try {
      // Check if the file is uploaded
      if (!req.file) {
          return res.status(400).json({ error: "No file uploaded" });
      }

      const filePath = req.file.path;
      const fileName = req.file.originalname;

      // Read the Excel file
      const workbook = xlsx.readFile(filePath);
      const sheetName = workbook.SheetNames[0]; // Get the first sheet
      const sheetData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

      // Convert the Excel data to a JSON string
      const fileContent = JSON.stringify(sheetData, null, 2);

      // Prepare the message for Anthropic API
      const messageContent = `The user uploaded an Excel file (${fileName}) with the following data: \n${fileContent}`;

      // Send the message to Anthropic
      const response = await anthropicClient.messages.create({
          model: "claude-3-5-sonnet-20241022", // Replace with your desired model
          max_tokens: 1024,
          messages: [{ role: "user", content: messageContent }],
      });

      // Cleanup: Remove the uploaded file after processing
      fs.unlinkSync(filePath);

      console.log(response);
      res.json(response);
  } catch (error) {
      console.error("Error processing file upload:", error);
      res.status(500).json({ error: "Failed to process the uploaded file" });
  }
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