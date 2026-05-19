import express from 'express';
import cors from 'cors';
import axios from 'axios';
import dotenv from 'dotenv';
import * as XLSX from 'xlsx';

dotenv.config();

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));

// Helper function for sleep
const sleep = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

app.post('/api/m365/upload-excel', async (req, res) => {
  try {
    const { filename, excelBase64 } = req.body;
    if (!filename || !excelBase64) {
      return res.status(400).json({ success: false, message: 'Missing file data' });
    }

    const tenantId = process.env.M365_TENANT_ID;
    const clientId = process.env.M365_CLIENT_ID;
    const clientSecret = process.env.M365_CLIENT_SECRET;
    const adminEmail = process.env.M365_ADMIN_EMAIL;

    // 1. Get Microsoft Graph Access Token
    const tokenResponse = await axios.post(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      new URLSearchParams({ 
        client_id: clientId!, 
        scope: 'https://graph.microsoft.com/.default', 
        client_secret: clientSecret!, 
        grant_type: 'client_credentials' 
      }).toString(),
      { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    );
    const accessToken = tokenResponse.data.access_token;

    const fileUrl = `https://graph.microsoft.com/v1.0/users/${adminEmail}/drive/root:/OmDedy_Projects/${filename}`;
    const contentUrl = `${fileUrl}:/content`;
    
    let existingCommentsMap: { [key: string]: { fachrul: string, barra: string } } = {};

    // STEP 1: READ EXISTING EXCEL (If it exists)
    try {
      const downloadResponse = await axios.get(contentUrl, {
        headers: { 'Authorization': `Bearer ${accessToken}` },
        responseType: 'arraybuffer'
      });

      const workbook = XLSX.read(downloadResponse.data, { type: 'buffer' });
      const sheetName = 'Timeline & Breakdown';
      
      if (workbook.SheetNames.includes(sheetName)) {
        const worksheet = workbook.Sheets[sheetName];
        const rows: any[] = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        // Assuming Column A is Task Title/ID, Column L is Fachrul Feedback (index 11), Column M is Barra Feedback (index 12)
        // Adjust the header row index if necessary (usually index 0 or 1)
        rows.forEach((row, index) => {
          if (index > 0 && row[0]) { // Generic skip header
            const taskTitle = String(row[0]).trim();
            existingCommentsMap[taskTitle] = {
              fachrul: row[11] || "-",
              barra: row[12] || "-"
            };
          }
        });
      }
    } catch (err: any) {
      if (err.response && err.response.status === 404) {
        console.log(`[INFO] File ${filename} does not exist yet. Creating a new one without merging comments.`);
        existingCommentsMap = {}; // Safe fallback
      } else {
        console.error("[ERROR] Failed to fetch existing Excel file:", err.message);
        throw err; 
      }
    }

    // STEP 2: MERGE LOGIC WITH NEW PAYLOAD
    const incomingBuffer = Buffer.from(excelBase64, 'base64');
    const newWorkbook = XLSX.read(incomingBuffer, { type: 'buffer' });
    const targetSheetName = 'Timeline & Breakdown';

    if (newWorkbook.SheetNames.includes(targetSheetName)) {
      const worksheet = newWorkbook.Sheets[targetSheetName];
      const rows: any[] = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      rows.forEach((row, index) => {
        if (index > 0 && row[0]) {
          const taskTitle = String(row[0]).trim();
          if (existingCommentsMap[taskTitle]) {
            row[11] = existingCommentsMap[taskTitle].fachrul;
            row[12] = existingCommentsMap[taskTitle].barra;
          } else {
            row[11] = row[11] || "-";
            row[12] = row[12] || "-";
          }
        }
      });

      // Update the worksheet with merged data
      const updatedWorksheet = XLSX.utils.aoa_to_sheet(rows);
      newWorkbook.Sheets[targetSheetName] = updatedWorksheet;
    }

    const finalBuffer = XLSX.write(newWorkbook, { type: 'buffer', bookType: 'xlsx' });

    // STEP 3: WRITE (OVERWRITE WITH RETRY)
    let uploadSuccess = false;
    let uploadResponseData: any = null;
    let attempts = 0;

    while (attempts < 3 && !uploadSuccess) {
      try {
        const uploadResponse = await axios.put(contentUrl, finalBuffer, {
          headers: { 
            'Authorization': `Bearer ${accessToken}`, 
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
          }
        });
        uploadResponseData = uploadResponse.data;
        uploadSuccess = true;
      } catch (err: any) {
        if (err.response && err.response.status === 423) {
          attempts++;
          if (attempts < 3) {
            console.log(`File is locked (423). Retry attempt ${attempts}/3...`);
            await sleep(2000);
          }
        } else {
          throw err;
        }
      }
    }

    if (!uploadSuccess) {
      return res.status(423).json({ 
        success: false, 
        message: "File sedang dibuka oleh user lain. Mohon tutup Excel Online atau tunggu 1 minute lalu coba lagi." 
      });
    }

    // 4. Generate or Retrieve the Permanent Sharing Link
    const itemId = uploadResponseData.id;
    const linkUrl = `https://graph.microsoft.com/v1.0/users/${adminEmail}/drive/items/${itemId}/createLink`;
    const linkResponse = await axios.post(
      linkUrl, 
      { type: 'edit', scope: 'anonymous' }, 
      { headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' } }
    );

    res.json({ success: true, embedUrl: linkResponse.data.link.webUrl });
  } catch (error: any) {
    console.error('M365 API Error:', error.response?.data || error.message);
    res.status(500).json({ success: false, message: 'Integration Failed' });
  }
});

// Root endpoint for status check
app.get('/', (req, res) => {
  res.send('Om Dedy Backend API is running. Use POST /api/m365/upload-excel for uploads.');
});

const PORT = Number(process.env.PORT) || 3000;
if (process.env.NODE_ENV !== 'production') {
  app.listen(PORT, '0.0.0.0', () => console.log(`Om Dedy Backend API running on port ${PORT}`));
}

export default app; // CRITICAL FOR VERCEL SERVERLESS

