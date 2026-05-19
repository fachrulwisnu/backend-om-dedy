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

app.get('/api/m365/get-comments', async (req, res) => {
  try {
    const { projectName } = req.query;
    if (!projectName) {
      return res.status(400).json({ success: false, message: 'Missing projectName' });
    }

    const tenantId = process.env.M365_TENANT_ID;
    const clientId = process.env.M365_CLIENT_ID;
    const clientSecret = process.env.M365_CLIENT_SECRET;
    const adminEmail = process.env.M365_ADMIN_EMAIL;

    // 1. Get Access Token
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

    const filename = `OM_DEDY_Timeline_${String(projectName).replace(/[^a-zA-Z0-9]/g, '_')}.xlsx`;
    const contentUrl = `https://graph.microsoft.com/v1.0/users/${adminEmail}/drive/root:/OmDedy_Projects/${filename}:/content`;

    let comments: { [key: string]: { fachrul: string, barra: string } } = {};

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
        
        rows.forEach((row, index) => {
          if (index > 0 && row[0]) {
            const taskTitle = String(row[0]).trim();
            comments[taskTitle] = {
              fachrul: row[11] || "-",
              barra: row[12] || "-"
            };
          }
        });
      }
    } catch (err: any) {
      if (err.response && err.response.status === 404) {
        return res.json({ success: true, comments: {} });
      }
      throw err;
    }

    res.json({ success: true, comments });
  } catch (error: any) {
    console.error('Fetch Comments Error:', error.response?.data || error.message);
    res.status(500).json({ success: false, message: 'Failed to fetch comments' });
  }
});

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

    const contentUrl = `https://graph.microsoft.com/v1.0/users/${adminEmail}/drive/root:/OmDedy_Projects/${filename}:/content`;
    
    // Convert Base64 to Buffer
    const fileBuffer = Buffer.from(excelBase64, 'base64');

    // STEP 3: WRITE (OVERWRITE WITH RETRY)
    let uploadSuccess = false;
    let uploadResponseData: any = null;
    let attempts = 0;

    while (attempts < 3 && !uploadSuccess) {
      try {
        const uploadResponse = await axios.put(contentUrl, fileBuffer, {
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

