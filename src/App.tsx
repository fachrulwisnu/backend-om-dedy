import { useState, useRef, type ChangeEvent } from "react";
import { Upload, FileSpreadsheet, CheckCircle2, AlertCircle, ExternalLink, Loader2 } from "lucide-react";
import { motion, AnimatePresence } from "motion/react";

export default function App() {
  const [file, setFile] = useState<File | null>(null);
  const [projectName, setProjectName] = useState("");
  const [uploading, setUploading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [shareUrl, setShareUrl] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileChange = (e: ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0];
    if (selectedFile) {
      if (
        selectedFile.type === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ||
        selectedFile.name.endsWith(".xlsx")
      ) {
        setFile(selectedFile);
        setError(null);
        setShareUrl(null);
      } else {
        setError("Please select a valid Excel (.xlsx) file.");
        setFile(null);
      }
    }
  };

  const readFileAsBase64 = (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const base64String = (reader.result as string).split(",")[1];
        resolve(base64String);
      };
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  };

  const handleUpload = async () => {
    if (!file) return;

    setUploading(true);
    setError(null);

    try {
      const base64 = await readFileAsBase64(file);
      
      // Clean up the project name from special characters to ensure a safe, static filename
      const cleanProjectName = (projectName || 'Project').replace(/[^a-zA-Z0-9]/g, '_');
      const filename = `OM_DEDY_Timeline_${cleanProjectName}.xlsx`;

      const response = await fetch("/api/m365/upload-excel", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          filename: filename,
          excelBase64: base64,
        }),
      });

      const data = await response.json();

      if (data.success) {
        setShareUrl(data.embedUrl);
      } else {
        setError(data.message || "Failed to upload file.");
      }
    } catch (err: any) {
      setError(err.message || "An unexpected error occurred.");
    } finally {
      setUploading(false);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center p-6 sm:p-12">
      <div className="w-full max-w-xl bg-white rounded-3xl shadow-xl border border-slate-100 overflow-hidden">
        <div className="bg-slate-900 p-8 text-white">
          <h1 className="text-2xl font-semibold tracking-tight flex items-center gap-3">
            <FileSpreadsheet className="text-emerald-400" />
            M365 Excel Bridge
          </h1>
          <p className="text-slate-400 mt-2 text-sm">
            Quickly upload local Excel files to OneDrive and generate secure sharing links for your organization.
          </p>
        </div>

        <div className="p-8">
          {!shareUrl ? (
            <motion.div 
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              className="space-y-6"
            >
              <div className="space-y-2">
                <label htmlFor="projectName" className="text-sm font-semibold text-slate-700 ml-1">
                  Project Name
                </label>
                <input
                  id="projectName"
                  type="text"
                  placeholder="Enter project name (e.g. Q2_Marketing)"
                  value={projectName}
                  onChange={(e) => setProjectName(e.target.value)}
                  className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:outline-none focus:ring-2 focus:ring-slate-900/5 focus:border-slate-900 transition-all text-sm"
                />
              </div>

              <div 
                id="dropzone"
                onClick={() => fileInputRef.current?.click()}
                className={`
                  relative border-2 border-dashed rounded-2xl p-10 
                  flex flex-col items-center justify-center transition-all cursor-pointer
                  ${file ? 'border-emerald-200 bg-emerald-50/30' : 'border-slate-200 hover:border-slate-300 bg-slate-50/50'}
                `}
              >
                <input 
                  type="file" 
                  ref={fileInputRef}
                  onChange={handleFileChange}
                  accept=".xlsx"
                  className="hidden"
                />
                
                {file ? (
                  <>
                    <FileSpreadsheet className="w-12 h-12 text-emerald-500 mb-4" />
                    <span className="text-sm font-medium text-slate-700">{file.name}</span>
                    <span className="text-xs text-slate-500 mt-1">{(file.size / 1024 / 1024).toFixed(2)} MB</span>
                  </>
                ) : (
                  <>
                    <div className="w-12 h-12 rounded-full bg-slate-100 flex items-center justify-center mb-4">
                      <Upload className="w-6 h-6 text-slate-400" />
                    </div>
                    <span className="text-sm font-medium text-slate-700 underline underline-offset-4">Click to select Excel file</span>
                    <span className="text-xs text-slate-400 mt-1">Maximum size: 50MB</span>
                  </>
                )}
              </div>

              <div className="space-y-4">
                <button
                  id="upload-button"
                  disabled={!file || uploading}
                  onClick={handleUpload}
                  className={`
                    w-full py-4 rounded-xl flex items-center justify-center gap-2 font-medium transition-all
                    ${!file || uploading 
                      ? 'bg-slate-100 text-slate-400 cursor-not-allowed' 
                      : 'bg-slate-900 text-white hover:bg-slate-800 shadow-lg active:scale-[0.98]'}
                  `}
                >
                  {uploading ? (
                    <>
                      <Loader2 className="w-5 h-5 animate-spin" />
                      Uploading to Microsoft 365...
                    </>
                  ) : (
                    <>Upload & Generate Link</>
                  )}
                </button>

                <AnimatePresence>
                  {error && (
                    <motion.div 
                      id="error-message"
                      initial={{ opacity: 0, height: 0 }}
                      animate={{ opacity: 1, height: 'auto' }}
                      exit={{ opacity: 0, height: 0 }}
                      className="p-4 bg-red-50 border border-red-100 rounded-xl flex items-start gap-3 text-red-700 text-sm"
                    >
                      <AlertCircle className="w-5 h-5 shrink-0" />
                      <span>{error}</span>
                    </motion.div>
                  )}
                </AnimatePresence>
              </div>
            </motion.div>
          ) : (
            <motion.div 
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              className="py-6 text-center space-y-6"
            >
              <div className="inline-flex items-center justify-center w-20 h-20 rounded-full bg-emerald-50 text-emerald-500 mb-2">
                <CheckCircle2 className="w-10 h-10" />
              </div>
              
              <div>
                <h2 className="text-xl font-bold text-slate-900">Successfully Uploaded!</h2>
                <p className="text-slate-500 mt-2 text-sm px-6">
                  The file is now in your OneDrive. Anyone in your organization with the link can edit it.
                </p>
              </div>

              <div className="p-4 bg-slate-50 border border-slate-100 rounded-2xl flex items-center justify-between gap-4">
                <div className="flex-1 text-left overflow-hidden">
                  <div className="text-[10px] uppercase tracking-wider text-slate-400 font-bold mb-1">Generated Link</div>
                  <div className="text-sm font-mono text-slate-600 truncate">{shareUrl}</div>
                </div>
                <a 
                  href={shareUrl} 
                  target="_blank" 
                  rel="noopener noreferrer"
                  className="flex items-center gap-1.5 px-4 py-2 bg-white border border-slate-200 rounded-xl text-slate-700 text-sm font-medium hover:bg-slate-50 transition-colors shrink-0 shadow-sm"
                >
                  <ExternalLink className="w-4 h-4" />
                  Open
                </a>
              </div>

              <button 
                onClick={() => {
                  setShareUrl(null);
                  setFile(null);
                }}
                className="text-slate-400 text-sm font-medium hover:text-slate-600 transition-colors"
              >
                Upload another file
              </button>
            </motion.div>
          )}
        </div>
      </div>
      
      <div className="mt-8 text-center text-slate-400 text-[11px] uppercase tracking-widest font-semibold flex items-center gap-2">
        <div className="w-1.5 h-1.5 rounded-full bg-emerald-500/50 animate-pulse" />
        Backend Secured with M365 Client Credentials
      </div>
    </div>
  );
}
