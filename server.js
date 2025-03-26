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



// ✅ Test API Route
app.get('/', (req, res) => {
    res.send('✅ Google Sheets API is running on Vercel!');
});

// ✅ Refresh the Directory Sheet
app.post('/refreshDirectory', async (req, res) => {
  try {
    const { spreadsheetId } = req.body;

    if (!spreadsheetId) {
      return res.status(400).json({ error: 'Missing spreadsheetId' });
    }

    // 1️⃣ Get all sheets
    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
    const sheetList = spreadsheet.data.sheets;

    // 2️⃣ Check if 'Directory' exists
    let directorySheet = sheetList.find(s => s.properties.title === 'Directory');

    // 3️⃣ Create Directory sheet if it doesn't exist
    if (!directorySheet) {
      const createRes = await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [
            {
              addSheet: {
                properties: {
                  title: 'Directory',
                }
              }
            }
          ]
        }
      });

      directorySheet = createRes.data.replies[0].addSheet;
    }

    const sheetId = directorySheet.properties.sheetId;

    // 4️⃣ Build data to write
    const rows = [
      ['Title', 'Sheet ID', 'Description', 'Last Edited'], // header
      ...sheetList
        .filter(sheet => sheet.properties.title !== 'Directory')
        .map(sheet => [
          sheet.properties.title,
          sheet.properties.sheetId,
          '', // leave description blank for now
          ''  // leave last edited blank for now
        ])
    ];

    // 5️⃣ Write to 'Directory' sheet
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: 'Directory!A1:D' + rows.length,
      valueInputOption: 'RAW',
      requestBody: {
        values: rows
      }
    });

    res.json({ message: 'Directory refreshed successfully', count: rows.length - 1 });
  } catch (error) {
    console.error('❌ Error refreshing directory:', error);
    res.status(500).json({ error: error.message });
  }
});

// 🔄 Add or update a single row in the Directory sheet
async function refreshDirectoryEntry(spreadsheetId, title, sheetId) {
  const now = new Date().toLocaleString();

  // Fetch current Directory contents
  const getRes = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: 'Directory!A2:D1000'
  });

  const values = getRes.data.values || [];

  // Check if sheet already exists in the directory
  const index = values.findIndex(row => row[0] === title);

  if (index !== -1) {
    // 📝 Update existing row
    values[index][1] = sheetId;
    values[index][3] = now;
  } else {
    // ➕ Add new row
    values.push([title, sheetId, '', now]);
  }

  // Write back
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: 'Directory!A2:D1000',
    valueInputOption: 'RAW',
    requestBody: {
      values
    }
  });

  console.log(`📘 Directory updated for: ${title}`);
}


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

        // Remove leading apostrophe from cells if the second character is '=' or a digit
        const cleanedValues = values.map(row => row.map(cell => {
            if (typeof cell === 'string' && cell.startsWith("'") && 
                (cell[1] === '=' || !isNaN(parseInt(cell[1])))) {
                return cell.slice(1);
            }
            return cell;
        }));

        await sheets.spreadsheets.values.update({
            spreadsheetId,
            range,
            valueInputOption: 'USER_ENTERED',
            requestBody: { values: cleanedValues }
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

// ✅ Create a New Sheet + Update Directory
app.post('/createSheet', async (req, res) => {
  try {
    console.log("🟢 Received createSheet request:", req.body);
    const { spreadsheetId, title } = req.body;
    if (!spreadsheetId || !title) return res.status(400).json({ error: "Missing parameters" });

    const createRes = await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{ addSheet: { properties: { title } } }]
      }
    });

    const sheetId = createRes.data.replies[0].addSheet.properties.sheetId;

    // ✅ Update the Directory sheet with this new sheet
    await refreshDirectoryEntry(spreadsheetId, title, sheetId);

    res.json({ message: `Sheet '${title}' created successfully and Directory updated` });
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
        if (!spreadsheetId || sheetId === undefined || !chartSpec) {
            return res.status(400).json({ error: "Missing parameters" });
        }

        // Define a default chart configuration
        const defaultChartSpec = {
            spec: {
                title: "Chart",
                basicChart: {
                    chartType: "LINE",
                    legendPosition: "BOTTOM_LEGEND",
                    axis: [
                        { position: "BOTTOM_AXIS", title: "Date" },
                        { position: "LEFT_AXIS", title: "Value" }
                    ],
                    domains: [
                        {
                            domain: {
                                sourceRange: {
                                    sources: [
                                        {
                                            sheetId: sheetId,
                                            startRowIndex: 1,
                                            endRowIndex: 8,
                                            startColumnIndex: 0,
                                            endColumnIndex: 1
                                        }
                                    ]
                                }
                            }
                        }
                    ],
                    series: [
                        {
                            series: {
                                sourceRange: {
                                    sources: [
                                        {
                                            sheetId: sheetId,
                                            startRowIndex: 1,
                                            endRowIndex: 8,
                                            startColumnIndex: 2,
                                            endColumnIndex: 3
                                        }
                                    ]
                                }
                            },
                            targetAxis: "LEFT_AXIS"
                        },
                        {
                            series: {
                                sourceRange: {
                                    sources: [
                                        {
                                            sheetId: sheetId,
                                            startRowIndex: 1,
                                            endRowIndex: 8,
                                            startColumnIndex: 4,
                                            endColumnIndex: 5
                                        }
                                    ]
                                }
                            },
                            targetAxis: "LEFT_AXIS"
                        }
                    ]
                }
            },
            position: {
                overlayPosition: {
                    anchorCell: {
                        sheetId: sheetId,
                        rowIndex: 10,
                        columnIndex: 0
                    },
                    offsetXPixels: 0,
                    offsetYPixels: 0,
                    widthPixels: 600,
                    heightPixels: 371
                }
            }
        };

        // Merge the incoming chartSpec with defaults (shallow merge)
        // For deeper structures like 'spec.basicChart', we merge those objects too.
        const completeChartSpec = {
            ...defaultChartSpec,
            ...chartSpec,
            spec: {
                ...defaultChartSpec.spec,
                ...(chartSpec.spec || {}),
                basicChart: {
                    ...defaultChartSpec.spec.basicChart,
                    ...((chartSpec.spec && chartSpec.spec.basicChart) || {})
                }
            },
            position: {
                ...defaultChartSpec.position,
                ...chartSpec.position
            }
        };

        await sheets.spreadsheets.batchUpdate({
            spreadsheetId,
            requestBody: { requests: [{ addChart: { chart: completeChartSpec } }] }
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


