from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from pathlib import Path
import shutil
from starlette.background import BackgroundTask

# Import the existing lambda entrypoint
from excel_processor import lambda_handler

app = FastAPI(title="Media Point Excel Processor")

DATA_DIR = Path("/data")
INPUT_FILE = DATA_DIR / "Bron.xlsx"
OUTPUT_FILE = DATA_DIR / "Modified_Bron.xlsx"
UPLOADED_FLAG = DATA_DIR / ".uploaded"
PROCESSED_FLAG = DATA_DIR / ".processed"


def cleanup_files():
    for path in (INPUT_FILE, OUTPUT_FILE, UPLOADED_FLAG, PROCESSED_FLAG):
        try:
            if path.exists():
                path.unlink()
        except Exception:
            pass


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

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    with INPUT_FILE.open("wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
    # update workflow flags
    try:
        if PROCESSED_FLAG.exists():
            PROCESSED_FLAG.unlink()
    except Exception:
        pass
    UPLOADED_FLAG.touch()

    return JSONResponse({"message": "Upload successful", "path": str(INPUT_FILE)})


@app.post("/run")
def run_processing():
    if not UPLOADED_FLAG.exists():
        raise HTTPException(status_code=400, detail="Please upload Bron.xlsx first.")

    # Call the existing lambda handler which is adapted to use local files
    result = lambda_handler(event={}, context=None)
    status_code = result.get("statusCode", 500)
    body = result.get("body", "Unknown error")
    if status_code == 200 and OUTPUT_FILE.exists():
        PROCESSED_FLAG.touch()
    return JSONResponse(status_code=status_code, content={"message": body})


@app.get("/download")
def download_output():
    # Only allow download if processed is newer than uploaded and file exists
    processed_ok = (
        OUTPUT_FILE.exists()
        and PROCESSED_FLAG.exists()
        and UPLOADED_FLAG.exists()
        and PROCESSED_FLAG.stat().st_mtime >= UPLOADED_FLAG.stat().st_mtime
    )
    if not processed_ok:
        raise HTTPException(
            status_code=404,
            detail="Modified_Bron.xlsx not found. Run the processor first.",
        )
    return FileResponse(
        path=str(OUTPUT_FILE),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="Modified_Bron.xlsx",
        background=BackgroundTask(cleanup_files),
    )


@app.get("/status")
def status():
    uploaded = UPLOADED_FLAG.exists() and INPUT_FILE.exists()
    processed = (
        OUTPUT_FILE.exists()
        and PROCESSED_FLAG.exists()
        and UPLOADED_FLAG.exists()
        and PROCESSED_FLAG.stat().st_mtime >= UPLOADED_FLAG.stat().st_mtime
    )
    return JSONResponse({"uploaded": uploaded, "processed": processed})
