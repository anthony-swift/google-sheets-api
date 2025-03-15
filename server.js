const express = require('express');
const { google } = require('googleapis');
const dotenv = require('dotenv');
const cors = require('cors');
const bodyParser = require('body-parser');

dotenv.config();
const app = express();
app.use(cors());
app.use(bodyParser.json());

const PORT = process.env.PORT || 3000;

// Load Google Service Account Credentials from Environment Variables
const credentials = JSON.parse(process.env.GOOGLE_APPLICATION_CREDENTIALS_JSON);

const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

const sheets = google.sheets({ version: 'v4', auth });

// Middleware to check for Vercel Protection Bypass
app.use((req, res, next) => {
    const bypassSecret = process.env.VERCEL_AUTOMATION_BYPASS_SECRET;
    const requestSecret = req.headers["x-vercel-protection-bypass"];

    if (!bypassSecret || requestSecret !== bypassSecret) {
        return res.status(403).json({ error: "Unauthorized access. Provide correct bypass secret." });
    }
    next();
});


// ✅ Test API Route
app.get('/', (req, res) => {
    res.send('✅ Google Sheets API is running on Vercel!');
});

// ✅ Read Data from Sheets
app.post('/read', async (req, res) => {
    try {
        console.log("🟢 Received read request:", req.body);
        const { spreadsheetId, range } = req.body;
        if (!spreadsheetId || !range) return res.status(400).json({ error: "Missing parameters" });

        const response = await sheets.spreadsheets.values.get({ spreadsheetId, range });
        res.json(response.data.values || []);
    } catch (error) {
        console.error("❌ Error reading data:", error);
        res.status(500).json({ error: error.message });
    }
});

// ✅ Write Data to Sheets
app.post('/write', async (req, res) => {
    try {
        console.log("🟢 Received write request:", req.body);
        const { spreadsheetId, range, values } = req.body;
        if (!spreadsheetId || !range || !values) return res.status(400).json({ error: "Missing parameters" });

        await sheets.spreadsheets.values.update({
            spreadsheetId,
            range,
            valueInputOption: 'RAW',
            requestBody: { values }
        });

        res.json({ message: "✅ Data written successfully" });
    } catch (error) {
        console.error("❌ Error writing data:", error);
        res.status(500).json({ error: error.message });
    }
});

// ✅ Format Cells (Bold, Colors, etc.)
app.post('/format', async (req, res) => {
    try {
        console.log("🟢 Received format request:", req.body);
        const { spreadsheetId, formatRequests } = req.body;
        if (!spreadsheetId || !formatRequests) return res.status(400).json({ error: "Missing parameters" });

        await sheets.spreadsheets.batchUpdate({
            spreadsheetId,
            requestBody: { requests: formatRequests }
        });

        res.json({ message: "✅ Cells formatted successfully" });
    } catch (error) {
        console.error("❌ Error formatting cells:", error);
        res.status(500).json({ error: error.message });
    }
});

// ✅ Create a New Sheet
app.post('/createSheet', async (req, res) => {
    try {
        console.log("🟢 Received createSheet request:", req.body);
        const { spreadsheetId, title } = req.body;
        if (!spreadsheetId || !title) return res.status(400).json({ error: "Missing parameters" });

        await sheets.spreadsheets.batchUpdate({
            spreadsheetId,
            requestBody: {
                requests: [{ addSheet: { properties: { title } } }]
            }
        });

        res.json({ message: `✅ Sheet '${title}' created successfully` });
    } catch (error) {
        console.error("❌ Error creating sheet:", error);
        res.status(500).json({ error: error.message });
    }
});

// ✅ Rename a Sheet
app.post('/renameSheet', async (req, res) => {
    try {
        console.log("🟢 Received renameSheet request:", req.body);
        const { spreadsheetId, sheetId, newTitle } = req.body;
        if (!spreadsheetId || !sheetId || !newTitle) return res.status(400).json({ error: "Missing parameters" });

        await sheets.spreadsheets.batchUpdate({
            spreadsheetId,
            requestBody: {
                requests: [{ updateSheetProperties: { properties: { sheetId, title: newTitle }, fields: "title" } }]
            }
        });

        res.json({ message: `✅ Sheet renamed to '${newTitle}' successfully` });
    } catch (error) {
        console.error("❌ Error renaming sheet:", error);
        res.status(500).json({ error: error.message });
    }
});

// ✅ Get List of Sheets
app.get('/getSheets', async (req, res) => {
    try {
        console.log("🟢 Received getSheets request:", req.query);
        const { spreadsheetId } = req.query;
        if (!spreadsheetId) return res.status(400).json({ error: "Missing spreadsheetId" });

        const response = await sheets.spreadsheets.get({ spreadsheetId });
        const sheetsList = response.data.sheets.map(sheet => ({
            title: sheet.properties.title,
            sheetId: sheet.properties.sheetId
        }));

        res.json(sheetsList);
    } catch (error) {
        console.error("❌ Error getting sheets:", error);
        res.status(500).json({ error: error.message });
    }
});

// ✅ Add Chart
app.post('/addChart', async (req, res) => {
    try {
        console.log("🟢 Received addChart request:", req.body);
        const { spreadsheetId, sheetId, chartSpec } = req.body;
        if (!spreadsheetId || sheetId === undefined || !chartSpec) return res.status(400).json({ error: "Missing parameters" });

        await sheets.spreadsheets.batchUpdate({
            spreadsheetId,
            requestBody: { requests: [{ addChart: { chart: chartSpec } }] }
        });

        res.json({ message: "✅ Chart added successfully" });
    } catch (error) {
        console.error("❌ Error adding chart:", error);
        res.status(500).json({ error: error.message });
    }
});

// ✅ Add Pivot Table
app.post('/addPivotTable', async (req, res) => {
    try {
        console.log("🟢 Received addPivotTable request:", req.body);
        const { spreadsheetId, sheetId, pivotTableSpec } = req.body;
        if (!spreadsheetId || sheetId === undefined || !pivotTableSpec) return res.status(400).json({ error: "Missing parameters" });

        await sheets.spreadsheets.batchUpdate({
            spreadsheetId,
            requestBody: { requests: [{ updateCells: { start: { sheetId }, rows: [{ values: [{ pivotTable: pivotTableSpec }] }], fields: "pivotTable" } }] }
        });

        res.json({ message: "✅ Pivot table added successfully" });
    } catch (error) {
        console.error("❌ Error adding pivot table:", error);
        res.status(500).json({ error: error.message });
    }
});

// ✅ Get Current Date & Time
app.get('/getDateTime', (req, res) => {
    const now = new Date();
    res.json({ utc: now.toISOString(), local: now.toLocaleString() });
});

// ✅ Start Server
app.listen(PORT, () => console.log(`🚀 Server running on http://localhost:${PORT}`));

// ✅ Export for Vercel Deployment
module.exports = app;
