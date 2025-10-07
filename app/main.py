from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse
from starlette.background import BackgroundTask
from typing import Optional
from io import BytesIO

# Import bytes-based processor
from excel_processor import process_excel_bytes

app = FastAPI(title="Media Point Excel Processor")

# In-memory buffers and simple timestamps
UPLOAD_BUFFER: Optional[bytes] = None
OUTPUT_BUFFER: Optional[bytes] = None
UPLOADED_AT: Optional[float] = None
PROCESSED_AT: Optional[float] = None


def cleanup_memory():
    global UPLOAD_BUFFER, OUTPUT_BUFFER, UPLOADED_AT, PROCESSED_AT
    UPLOAD_BUFFER = None
    OUTPUT_BUFFER = None
    UPLOADED_AT = None
    PROCESSED_AT = None


@app.get("/", response_class=HTMLResponse)
def read_root():
    return """
<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Media Point Excel Processor</title>
    <style>
      body { font-family: system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif; max-width: 420px; margin: 40px auto; padding: 0 16px; text-align: center; }
      h1 { font-size: 1.6rem; margin-bottom: 24px; }
      form { display: flex; flex-direction: column; gap: 12px; align-items: center; }
      button { padding: 10px 18px; border-radius: 6px; border: 1px solid #111827; background: #111827; color: white; cursor: pointer; }
      button:disabled { opacity: 0.6; cursor: not-allowed; }
      input[type=file] { padding: 8px; }
      .muted { color: #6b7280; font-size: 0.95rem; min-height: 1.25rem; }
      .hidden { display: none; }
      #downloadBtn { margin-top: 20px; }
    </style>
  </head>
  <body>
    <h1>Media Point Excel Processor</h1>
    <form id="uploadForm" enctype="multipart/form-data">
      <input id="fileInput" type="file" name="file" accept=".xlsx" required />
      <button id="uploadBtn" type="submit">Upload Bron.xlsx</button>
      <div id="status" class="muted"></div>
    </form>

    <button id="downloadBtn" class="hidden">Download Modified_Bron.xlsx</button>

    <script>
      const uploadForm = document.getElementById('uploadForm');
      const fileInput = document.getElementById('fileInput');
      const uploadBtn = document.getElementById('uploadBtn');
      const statusBox = document.getElementById('status');
      const downloadBtn = document.getElementById('downloadBtn');

      function toggleDownload(show) {
        if (show) {
          downloadBtn.classList.remove('hidden');
        } else {
          downloadBtn.classList.add('hidden');
        }
      }

      async function refreshState() {
        try {
          const res = await fetch('/status');
          if (!res.ok) return;
          const data = await res.json();
          toggleDownload(data.processed);
        } catch (e) { /* ignore */ }
      }

      uploadForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        if (!fileInput.files.length) return;
        statusBox.textContent = 'Uploading...';
        uploadBtn.disabled = true;
        const formData = new FormData();
        formData.append('file', fileInput.files[0]);
        try {
          const res = await fetch('/upload', { method: 'POST', body: formData });
          if (!res.ok) {
            const msg = await res.text();
            throw new Error(msg || 'Upload failed');
          }
          statusBox.textContent = 'Processing...';
          toggleDownload(false);
          const runRes = await fetch('/run', { method: 'POST' });
          const runData = await runRes.json().catch(() => ({}));
          if (!runRes.ok) {
            const msg = runData && runData.message ? runData.message : 'Run failed';
            throw new Error(msg);
          }
          statusBox.textContent = 'Download is ready.';
          toggleDownload(true);
        } catch (err) {
          statusBox.textContent = `Error: ${err.message}`;
        } finally {
          uploadBtn.disabled = false;
        }
      });

      downloadBtn.addEventListener('click', async () => {
        try {
          const res = await fetch('/status');
          const data = await res.json();
          if (!data.processed) {
            return;
          }
          window.location.href = '/download';
          setTimeout(() => {
            toggleDownload(false);
            statusBox.textContent = '';
          }, 750);
        } catch (_) {
          statusBox.textContent = 'Unable to download right now.';
        }
      });

      document.addEventListener('DOMContentLoaded', refreshState);
    </script>
  </body>
 </html>
        """


@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Only .xlsx files are supported")
    global UPLOAD_BUFFER, OUTPUT_BUFFER, UPLOADED_AT, PROCESSED_AT
    data = await file.read()
    if not data:
        raise HTTPException(status_code=400, detail="Uploaded file is empty")
    UPLOAD_BUFFER = data
    UPLOADED_AT = __import__("time").time()
    # Reset processed state on new upload
    OUTPUT_BUFFER = None
    PROCESSED_AT = None
    return JSONResponse({"message": "Upload successful"})


@app.post("/run")
def run_processing():
    global UPLOAD_BUFFER, OUTPUT_BUFFER, PROCESSED_AT
    if UPLOAD_BUFFER is None:
        raise HTTPException(status_code=400, detail="Please upload Bron.xlsx first.")
    try:
        output_bytes = process_excel_bytes(UPLOAD_BUFFER)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Processing failed: {e}")
    OUTPUT_BUFFER = output_bytes
    PROCESSED_AT = __import__("time").time()
    return JSONResponse(
        status_code=200, content={"message": "File processed successfully."}
    )


@app.get("/download")
def download_output():
    if OUTPUT_BUFFER is None:
        raise HTTPException(
            status_code=404,
            detail="Modified_Bron.xlsx not found. Run the processor first.",
        )
    background = BackgroundTask(cleanup_memory)
    return StreamingResponse(
        BytesIO(OUTPUT_BUFFER),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=Modified_Bron.xlsx"},
        background=background,
    )


@app.get("/status")
def status():
    uploaded = UPLOAD_BUFFER is not None
    processed = (
        OUTPUT_BUFFER is not None
        and uploaded
        and ((PROCESSED_AT or 0) >= (UPLOADED_AT or 0))
    )
    return JSONResponse({"uploaded": uploaded, "processed": processed})
