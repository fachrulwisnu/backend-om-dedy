import express from 'express';
import cors from 'cors';
import axios from 'axios';
import dotenv from 'dotenv';
import * as XLSX from 'xlsx';

dotenv.config();

// Placeholder for the database client.
// The user should replace this with their actual DB client (Prisma, Supabase, etc.)
// For now, it is defined to satisfy TypeScript and provide a structure.
const db: any = (global as any).db || {
  task: { updateMany: async () => ({ count: 1 }) },
  historyLog: { create: async () => ({}) }
};

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

app.post('/api/m365/sync-feedback', async (req, res) => {
  try {
    const { filename, projectName } = req.body;
    if (!filename || !projectName) {
      return res.status(400).json({ success: false, message: 'Missing parameters' });
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

    // 2. Download the Excel file from OneDrive (the updated one)
    const downloadUrl = `https://graph.microsoft.com/v1.0/users/${adminEmail}/drive/root:/OmDedy_Projects/${filename}:/content`;
    const downloadResponse = await axios.get(downloadUrl, {
      headers: { 'Authorization': `Bearer ${accessToken}` },
      responseType: 'arraybuffer'
    });

    // 3. Parse the Excel file to extract updatedTasks
    const workbook = XLSX.read(downloadResponse.data, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    // This assumes Excel columns: taskName, fachrulFeedback, barraFeedback
    const updatedTasks: any[] = XLSX.utils.sheet_to_json(sheet);

    // 4. DATABASE UPDATE AND HISTORY LOGGING (Lead Backend Architect Request)
    for (const task of updatedTasks) {
      try {
        // 1. Update the main task table
        const updatedRecord = await db.task.updateMany({
          where: { 
            task_title: task.taskName,
            projectName: projectName 
          },
          data: {
            fachrul_feedback: task.fachrulFeedback,
            barra_feedback: task.barraFeedback
          }
        });

        // 2. Insert into History Logs as "System (M365 Sync)"
        // Check if update was successful (depending on DB client response format)
        if (updatedRecord && (updatedRecord.count > 0 || updatedRecord.length > 0)) {
          await db.historyLog.create({
            data: {
              action: `Updated Feedback via Microsoft 365 Sync`,
              entityName: task.taskName,
              picName: "System (M365 Sync)", // <--- CRITICAL REQUIREMENT
              createdAt: new Date()
            }
          });
        }
      } catch (dbErr) {
        console.error(`Failed to update DB for task ${task.taskName}:`, dbErr);
        // Continue to next task even if one fails
      }
    }

    res.json({ 
      success: true, 
      message: 'Microsoft 365 data sync and database updates completed successfully.',
      taskCount: updatedTasks.length 
    });
  } catch (error: any) {
    console.error('M365 Sync Error:', error.response?.data || error.message);
    res.status(500).json({ success: false, message: 'Sync Integration Failed' });
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

