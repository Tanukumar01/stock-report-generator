require('dotenv').config();
const express = require('express');
const { google } = require('googleapis');
const yahooFinance = require('yahoo-finance2').default;
const fs = require('fs');

const app = express();
const PORT = 3000;

// Google Sheets setup
const auth = new google.auth.GoogleAuth({
  credentials: require('./service-account.json'),
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});
const sheets = google.sheets({ version: 'v4', auth });

const SHEET_ID = process.env.GOOGLE_SHEET_ID;

// Helper: Analyze stock
function analyzeStock(current, avg10) {
  const risk = ((Math.abs(current - avg10) / avg10) * 100).toFixed(2);
  const good = current > avg10 ? 'Yes' : 'No';
  const bad = current < avg10 ? 'Yes' : 'No';
  const sentiment = current > avg10 ? 'Positive' : 'Negative';
  const hold = good === 'Yes' ? 'Hold' : 'Sell';
  return { good, bad, risk, sentiment, hold };
}

// Main endpoint
app.get('/analyze', async (req, res) => {
  try {
    // 1. Read input stocks (E5:K9)
    const readRange = 'Sheet1!E5:K9';
    const { data } = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: readRange,
    });
    const stocks = data.values.filter(row => row[0] && row[1]); // Stock Name, Ticker

    // 2. Prepare output
    const output = [];

    for (const row of stocks) {
      const [stockName, ticker, , , , , avg10] = row;
      // Fetch current price from Yahoo Finance
      let currentPrice = null;
      try {
        const quote = await yahooFinance.quote(ticker);
        currentPrice = quote.regularMarketPrice;
      } catch (e) {
        currentPrice = row[3]; // fallback to sheet value
      }
      // Analyze
      const { good, bad, risk, sentiment, hold } = analyzeStock(Number(currentPrice), Number(avg10));
      output.push([
        stockName,
        currentPrice,
        good,
        bad,
        risk,
        sentiment,
        hold,
      ]);
    }

    // 3. Write to Output Report (E15:K...)
    const writeRange = `Sheet1!E15:K${15 + output.length - 1}`;
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: writeRange,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: output },
    });

    res.json({ message: 'Output report updated!', output });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});