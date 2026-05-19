import express from 'express';
import cors from 'cors';
import axios from 'axios';
import dotenv from 'dotenv';
import * as XLSX from 'xlsx';
import { createClient } from '@supabase/supabase-js';

dotenv.config();

// Initialize Supabase Client
const supabaseUrl = process.env.SUPABASE_URL || '';
const supabaseKey = process.env.SUPABASE_SERVICE_ROLE_KEY || '';
const supabase = createClient(supabaseUrl, supabaseKey);

// Placeholder for the database client.
// The user should replace this with their actual DB client (Prisma, Supabase, etc.)
// For now, these are mapped to Supabase calls.
const db: any = {
  task: { 
    updateMany: async ({ where, data }: any) => {
      const { data: result, error } = await supabase
        .from('tasks')
        .update(data)
        .match(where)
        .select();
      if (error) throw error;
      return { count: Array.isArray(result) ? result.length : 0 };
    }
  },
  historyLog: {
    create: async ({ data }: any) => {
      const { error } = await supabase
        .from('history_logs')
        .insert(data);
      if (error) throw error;
    }
  },
  project: {
    findFirst: async ({ where }: any) => {
      const { data, error } = await supabase
        .from('projects')
        .select('*')
        .eq('id', where.id)
        .maybeSingle();
      if (error) {
        console.error("[DB ERROR] findFirst failed:", error);
        return null;
      }
      return data;
    },
    update: async ({ where, data }: any) => {
      const { data: result, error } = await supabase
        .from('projects')
        .update(data)
        .eq('id', where.id)
        .select()
        .single();
      if (error) throw error;
      return result;
    }
  }
};

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));

app.post('/api/m365/upload-excel', async (req, res) => {
  try {
    const { projectId, projectName, excelBase64 } = req.body;
    
    // Log incoming payload as requested
    console.log("[EXPORT PAYLOAD] Received:", { projectName, projectId });

    if (!projectId || !projectName || !excelBase64) {
      return res.status(400).json({ success: false, message: 'Missing projectId, projectName or Excel data' });
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

    // ONE-WAY EXPORT: Uses static filename
    const cleanProjectName = projectName.replace(/[^a-zA-Z0-9]/g, '_');
    const filename = `OM_DEDY_Timeline_${cleanProjectName}.xlsx`;
    
    console.log(`[EXPORT] Prepared for one-way push: ${filename}`);

    const fileBuffer = Buffer.from(excelBase64, 'base64');
    const uploadUrl = `https://graph.microsoft.com/v1.0/users/${adminEmail}/drive/root:/OmDedy_Projects/${filename}:/content`;
    
    // Robust overwrite with retry to handle M365 file locking
    const uploadWithRetry = async (url: string, data: Buffer, token: string, retries = 3) => {
      for (let i = 0; i < retries; i++) {
        try {
          return await axios.put(url, data, {
            headers: { 
              'Authorization': `Bearer ${token}`, 
              'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
            }
          });
        } catch (err: any) {
          if (err.response && err.response.status === 423 && i < retries - 1) {
            console.log(`[EXPORT] File locked (423), retrying in 2s... (Attempt ${i + 1}/${retries})`);
            await new Promise(resolve => setTimeout(resolve, 2000));
            continue;
          }
          throw err;
        }
      }
    };

    let uploadResponse;
    try {
      uploadResponse = await uploadWithRetry(uploadUrl, fileBuffer, accessToken);
    } catch (uploadErr: any) {
      if (uploadErr.response && uploadErr.response.status === 423) {
        return res.status(423).json({ 
          success: false, 
          message: "File sedang digunakan oleh user lain. Mohon tunggu 1 menit lalu coba lagi." 
        });
      }
      throw uploadErr;
    }

    const itemId = uploadResponse?.data.id;

    // Generate or retrieve the permanent sharing link
    const linkUrl = `https://graph.microsoft.com/v1.0/users/${adminEmail}/drive/items/${itemId}/createLink`;
    const linkResponse = await axios.post(
      linkUrl, 
      { type: 'edit', scope: 'anonymous' }, 
      { headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' } }
    );

    const sharingUrl = linkResponse.data.link.webUrl;

    // Update database with latest export metadata
    try {
      const { error } = await supabase
        .from('projects')
        .update({ 
          current_excel_filename: filename, 
          current_sharing_url: sharingUrl 
        })
        .eq('id', projectId);

      if (error) {
        console.error("[SUPABASE UPDATE ERROR]:", error);
      } else {
        console.log("[SUPABASE UPDATE SUCCESS] DB updated for project:", projectName);
      }
    } catch (dbCrash) {
      console.error("[SUPABASE CRASH]:", dbCrash);
    }

    res.json({ success: true, embedUrl: sharingUrl });
  } catch (error: any) {
    console.error('M365 API Error:', error.response?.data || error.message);
    res.status(500).json({ success: false, message: 'Integration Failed' });
  }
});

app.post('/api/m365/sync-feedback', async (req, res) => {
  try {
    const { projectId, projectName } = req.body;
    if (!projectId || !projectName) {
      return res.status(400).json({ success: false, message: 'Project ID and Name are required' });
    }

    // 1. Fetch the latest project data from the database using projectId
    const project = await db.project.findFirst({ where: { id: projectId } });
    
    if (!project || !project.current_excel_filename) {
      return res.status(404).json({ success: false, message: 'File Excel tidak ditemukan. Silakan ekspor terlebih dahulu.' });
    }

    const filename = project.current_excel_filename;
    console.log(`[MANUAL SYNC] Initiated for projectId: ${projectId}, using filename from DB: ${filename}`);

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

    // 2. Download Excel directly from OneDrive
    const downloadUrl = `https://graph.microsoft.com/v1.0/users/${adminEmail}/drive/root:/OmDedy_Projects/${filename}:/content`;
    const fileResponse = await axios.get(downloadUrl, {
      headers: { 'Authorization': `Bearer ${accessToken}` },
      responseType: 'arraybuffer' 
    });

    // 3. Parse Excel using SheetJS
    const workbook = XLSX.read(fileResponse.data, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const rawData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

    // 4. Extract Updates and Integrate Database Logic
    let updateCount = 0;
    // Row 12 is index 12 in 0-based array (13th row)
    for (let i = 12; i < rawData.length; i++) {
      const row: any = rawData[i];
      if (!row || row.length === 0) continue;

      let taskName = row[2]; // Column C (Index 2)
      const fachrulFeedback = row[11]; // Column L (Index 11)
      const barraFeedback = row[12]; // Column M (Index 12)
      
      if (taskName && !String(taskName).includes('TOTAL') && taskName !== 'Task') {
        taskName = String(taskName).trim(); // Prevent trailing space mismatches

        const fFeedbackVal = (fachrulFeedback && fachrulFeedback !== '-') ? String(fachrulFeedback).trim() : null;
        const bFeedbackVal = (barraFeedback && barraFeedback !== '-') ? String(barraFeedback).trim() : null;

        if (fFeedbackVal !== null || bFeedbackVal !== null) {
          console.log(`[SYNC ATTEMPT] Task: "${taskName}" | F-Feed: "${fFeedbackVal}" | B-Feed: "${bFeedbackVal}"`);
          
          try {
            // 1. Update the main task table
            const updatedRecord = await db.task.updateMany({
              where: { 
                task_title: taskName,
                projectName: projectName 
              },
              data: {
                fachrul_feedback: fFeedbackVal,
                barra_feedback: bFeedbackVal
              }
            });

            // 2. Insert into History Logs as "System (M365 Sync)"
            // Check if update was successful (depending on DB client response format)
            if (updatedRecord && (updatedRecord.count > 0 || updatedRecord.length > 0)) {
              await db.historyLog.create({
                data: {
                  action: `Manual Sync Feedback from M365`,
                  entityName: taskName,
                  picName: "System (M365 Sync)", // <--- CRITICAL REQUIREMENT
                  createdAt: new Date()
                }
              });
              updateCount++;
            }
          } catch (dbErr) {
            console.error(`[DB ERROR] Failed to update task "${taskName}":`, dbErr);
          }
        }
      }
    }
    console.log(`[SYNC FINAL] Successfully processed ${updateCount} rows.`);

    res.status(200).json({ 
      success: true, 
      message: `Successfully synced ${updateCount} tasks from M365.`,
      updateCount
    });

  } catch (error: any) {
    if (error.response && error.response.status === 404) {
      return res.status(404).json({ success: false, message: 'File Excel tidak ditemukan di Microsoft 365. Lakukan export terlebih dahulu.' });
    }
    console.error('[MANUAL SYNC ERROR]:', error.response?.data || error.message);
    res.status(500).json({ success: false, message: 'Gagal melakukan sinkronisasi dengan Microsoft 365.' });
  }
});

// Permanent Redirect Endpoint
app.get('/api/m365/share-link/:projectId', async (req, res) => {
  try {
    const { projectId } = req.params;
    
    // 1. Fetch the latest project data from our database using projectId
    const project = await db.project.findFirst({ where: { id: projectId } });
    
    if (!project || !project.current_sharing_url) {
      return res.status(404).send("File Excel untuk proyek ini belum pernah diexport atau tidak ditemukan.");
    }

    console.log(`[REDIRECT] Forwarding user to latest Excel for projectId: ${projectId}`);
    
    // 2. Perform a HTTP 302 temporary redirect to the actual dynamic Microsoft link
    return res.redirect(project.current_sharing_url);

  } catch (error: any) {
    console.error("[REDIRECT ERROR]:", error.message);
    res.status(500).send("Gagal mengalihkan komponen tautan Microsoft 365.");
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

