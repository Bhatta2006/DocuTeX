import os
import pytest
from fastapi.testclient import TestClient
from app import app

client = TestClient(app)

def test_get_index():
    """Verify that the home page serves correctly and contains LaTeX converter markers."""
    response = client.get("/")
    assert response.status_code == 200
    assert "LaTeX" in response.text
    assert "style.css" in response.text
    assert "app.js" in response.text

def test_convert_invalid_file_extension():
    """Verify that upload fails for invalid file extensions (non-.tex files)."""
    files = {"file": ("document.txt", b"Some plain text content", "text/plain")}
    response = client.post("/convert", files=files)
    assert response.status_code == 400
    assert "Only .tex files are supported" in response.json()["detail"]

def test_convert_valid_latex_and_download():
    """
    Test compiling a minimal valid LaTeX file:
    1. Uploads .tex file to /convert.
    2. Expects successful execution with valid download URL and captured logs.
    3. Downloads generated docx.
    4. Expects the temp directories to be deleted after download finishes.
    """
    tex_content = r"""\documentclass{article}
\begin{document}
\title{Automated Integration Test}
\author{Antigravity Agent}
\maketitle
\begin{abstract}
This is a test abstract to verify LaTeX conversion in the web application pipeline.
\end{abstract}
\section{Introduction}
Testing section formatting and subprocess execution logs.
\end{document}
"""
    
    files = {"file": ("test.tex", tex_content.encode("utf-8"), "application/x-latex")}
    response = client.post("/convert", files=files)
    assert response.status_code == 200
    
    res_json = response.json()
    assert res_json["success"] is True
    assert "download_url" in res_json
    assert "logs" in res_json
    assert len(res_json["logs"]) > 0
    
    # Print compilation logs for verification output
    print("\n--- Spat Subprocess Logs ---")
    for log in res_json["logs"]:
        print(log)
    print("----------------------------\n")
    
    # Download the resulting file
    download_url = res_json["download_url"]
    download_response = client.get(download_url)
    assert download_response.status_code == 200
    assert download_response.headers["content-type"] == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    
    # Ensure auto-cleanup completed (TestClient executes background tasks synchronously before return)
    session_id = download_url.split("/")[-1]
    session_dir = os.path.join("temp_conversions", session_id)
    assert not os.path.exists(session_dir)

def test_convert_resume_latex():
    """
    Test compiling a LaTeX resume containing custom resume commands (e.g. \resumeSubheading, \resumeItem)
    and formatting wrappers (e.g. {\Huge \scshape ...}).
    """
    resume_latex = r"""\documentclass{article}
\begin{document}

{\Huge \scshape Ramakrishna V Bhat}
\vspace{4pt}
\small
+91 9480106354 \quad \href{mailto:ramkrsnabhat@gmail.com}{ramkrsnabhat@gmail.com} \quad \href{https://github.com/bhatta2006}{github.com/bhatta2006}

\section{Education}
\resumeSubHeadingListStart
  \resumeSubheading
    {RV College of Engineering}{Bengaluru, India}
    {Bachelor of Engineering -- Computer Science Engineering}{2024 -- Present}
    \resumeItemListStart
      \resumeItem{Relevant Coursework: Data Structures & Algorithms, Software Engineering.}
      \resumeItem{LeetCode Rating: 1825 | 450+ problems solved.}
    \resumeItemListEnd
\resumeSubHeadingListEnd

\section{Work Experience}
\resumeSubHeadingListStart
  \resumeSubheading
    {Samsung R&D Institute (Samsung PRISM)}{Bengaluru, India}
    {AI/ML Research Intern}{November 2025 -- Present}
    \resumeItemListStart
      \resumeItem{Contributed to a \textbf{Federated Learning}-based video content moment retrieval system.}
      \resumeItem{Analyzed AI model outputs to identify failure patterns.}
    \resumeItemListEnd
\resumeSubHeadingListEnd

\end{document}
"""
    files = {"file": ("resume.tex", resume_latex.encode("utf-8"), "application/x-latex")}
    response = client.post("/convert", files=files)
    assert response.status_code == 200
    
    res_json = response.json()
    assert res_json["success"] is True
    assert "download_url" in res_json
    
    logs_str = "\n".join(res_json["logs"])
    assert "Extracted document title from header: Ramakrishna V Bhat" in logs_str
    
    # Download the resulting file
    download_url = res_json["download_url"]
    download_response = client.get(download_url)
    assert download_response.status_code == 200
    assert download_response.headers["content-type"] == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    
    # Ensure auto-cleanup completed
    session_id = download_url.split("/")[-1]
    session_dir = os.path.join("temp_conversions", session_id)
    assert not os.path.exists(session_dir)

