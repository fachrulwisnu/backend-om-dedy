import express from 'express';
import cors from 'cors';
import axios from 'axios';
import dotenv from 'dotenv';

dotenv.config();

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));

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

    let itemId;
    const checkUrl = `https://graph.microsoft.com/v1.0/users/${adminEmail}/drive/root:/OmDedy_Projects/${filename}`;

    try {
      // 2. SMART CHECK: See if the file already exists to bypass locking issues
      const checkResponse = await axios.get(checkUrl, {
        headers: { 'Authorization': `Bearer ${accessToken}` }
      });
      itemId = checkResponse.data.id; // File exists! Reuse this ID without overwriting
    } catch (err: any) {
      // 3. IF FILE DOES NOT EXIST (404), UPLOAD IT FOR THE FIRST TIME
      if (err.response && err.response.status === 404) {
        const fileBuffer = Buffer.from(excelBase64, 'base64');
        const uploadUrl = `${checkUrl}:/content`;
        const uploadResponse = await axios.put(uploadUrl, fileBuffer, {
          headers: { 
            'Authorization': `Bearer ${accessToken}`, 
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
          }
        });
        itemId = uploadResponse.data.id;
      } else {
        throw err; // Rethrow any other unexpected network errors
      }
    }

    // 4. Generate or Retrieve the Permanent Sharing Link (Safely handles open files)
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

