import os
import sys
import uuid
import shutil
import tempfile
import subprocess
from pathlib import Path
from typing import List, Dict, Any
from fastapi import FastAPI, UploadFile, File, BackgroundTasks, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles

app = FastAPI(
    title="LaTeX to DOCX Converter",
    description="Minimal Web Interface for LaTeX to DOCX Conversion",
    version="1.0.0"
)

# Define directories
BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
TEMP_DIR = Path(tempfile.gettempdir()) / "temp_conversions"

# Ensure directories exist
STATIC_DIR.mkdir(exist_ok=True)
TEMP_DIR.mkdir(exist_ok=True)

# Mount static files (HTML, CSS, JS)
# We will serve the index.html directly from root GET '/' and mount static files for css/js
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

def cleanup_file_or_dir(path: Path):
    """Utility function to clean up files or directories in background tasks."""
    try:
        if path.is_file():
            path.unlink()
        elif path.is_dir():
            shutil.rmtree(path)
    except Exception as e:
        print(f"Error during cleanup of {path}: {e}")

@app.get("/", response_class=HTMLResponse)
def get_index():
    """Serves the main web interface page."""
    index_path = STATIC_DIR / "index.html"
    if not index_path.exists():
        # Fallback if index.html is missing
        return HTMLResponse(
            "<html><body><h1>LaTeX to DOCX Converter</h1><p>Frontend missing. Please create static/index.html</p></body></html>"
        )
    return HTMLResponse(content=index_path.read_text(encoding="utf-8"))

@app.post("/convert")
def convert_latex(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...)
):
    """
    Endpoint to upload a .tex file and run the conversion.
    Returns the logs and a download URL for the generated docx file.
    """
    if not file.filename.endswith(".tex"):
        raise HTTPException(status_code=400, detail="Only .tex files are supported.")

    # Generate unique IDs for the session
    session_id = str(uuid.uuid4())
    session_dir = TEMP_DIR / session_id
    session_dir.mkdir(parents=True, exist_ok=True)

    input_tex_path = session_dir / "input.tex"
    output_docx_path = session_dir / "output.docx"

    try:
        # Save uploaded file
        with open(input_tex_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
    except Exception as e:
        cleanup_file_or_dir(session_dir)
        raise HTTPException(status_code=500, detail=f"Failed to save uploaded file: {str(e)}")

    # Run converter as a subprocess
    script_path = BASE_DIR / "generate_docx.py"
    
    # We pass the absolute paths to the script as arguments
    cmd = [
        sys.executable,
        str(script_path),
        str(input_tex_path),
        str(output_docx_path)
    ]

    try:
        # Run conversion script and capture output
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            cwd=str(BASE_DIR)
        )
    except Exception as e:
        cleanup_file_or_dir(session_dir)
        raise HTTPException(
            status_code=500,
            detail=f"Failed to execute conversion script: {str(e)}"
        )

    # Process logs
    stdout_lines = [line.strip() for line in result.stdout.splitlines() if line.strip()]
    stderr_lines = [line.strip() for line in result.stderr.splitlines() if line.strip()]
    
    success = output_docx_path.exists() and result.returncode == 0

    if success:
        # Setup scheduled cleanup of input file immediately, and output file directory
        # after a reasonable time. Or we can delete it via a download path background task.
        # But if the user never downloads, we should ensure it gets cleaned up.
        # We'll register cleanup for the input file now.
        background_tasks.add_task(cleanup_file_or_dir, input_tex_path)
        
        return {
            "success": True,
            "logs": stdout_lines,
            "download_url": f"/download/{session_id}"
        }
    else:
        # Clean up immediately on failure
        background_tasks.add_task(cleanup_file_or_dir, session_dir)
        error_msg = stderr_lines[-1] if stderr_lines else "Conversion failed for unknown reasons."
        return {
            "success": False,
            "logs": stdout_lines + stderr_lines,
            "error": error_msg
        }

@app.get("/download/{session_id}")
def download_docx(session_id: str, background_tasks: BackgroundTasks):
    """
    Downloads the converted docx file.
    Triggers cleanup of the session folder once the download is complete.
    """
    session_dir = TEMP_DIR / session_id
    output_docx_path = session_dir / "output.docx"

    if not output_docx_path.exists():
        raise HTTPException(status_code=404, detail="Converted file not found or already downloaded.")

    # Schedule cleanup of the entire session directory after file is sent
    background_tasks.add_task(cleanup_file_or_dir, session_dir)

    return FileResponse(
        path=str(output_docx_path),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename="converted.docx"
    )
