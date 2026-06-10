# DocuTeX

**DocuTeX** is a lightweight, pure Python LaTeX-to-DOCX converter with a minimal, high-performance web dashboard. Unlike traditional tools, DocuTeX does not require heavy external dependencies like Pandoc or MacTeX/MiKTeX to compile documents to Word. It parses LaTeX syntax directly and generates professionally formatted `.docx` files using native Python libraries.

The application features a minimalist, monochrome brutalist web UI, allowing users to upload `.tex` files, view compilation logs in real-time, and download the finished `.docx` files instantly.

---

## Features

- **No Heavy External Dependencies**: Pure Python conversion (built on `python-docx`).
- **Real-Time Logs**: View detailed compilation logs directly on the web interface as the document is being compiled.
- **Robust Math & Table Parsing**:
  - Handles mathematical formulas and inline equations gracefully.
  - Supports standard and complex table structures (`tabular`, `tabularx`, and `longtable`).
  - Supports list structures (`itemize`, `enumerate`) with correct numbering.
- **Resume-Tailored Parsing**:
  - Out-of-the-box support for popular LaTeX resume templates (recognizes tags like `\resumeSubheading`, `\resumeItem`, `\resumeItemListStart`, custom margins, etc.).
- **Automatic Cleanup**: Temporary files and output directories are automatically removed from the server after the download is completed.
- **Minimalist Web UI**: Fast, responsive, pitch-black dark mode interface with clean white borders and crisp micro-animations.

---

## Directory Structure

```text
DocuTeX/
├── static/                  # Frontend web interface assets
│   ├── index.html           # Minimalist HTML dashboard layout
│   ├── style.css            # Monochrome brutalist stylesheet
│   └── app.js               # Reactive upload and real-time logging script
├── tests/                   # Test suite
│   └── test_app.py          # FastAPI integration tests
├── app.py                   # FastAPI web server backend
├── generate_docx.py         # Direct LaTeX parsing and python-docx compiler
├── requirements.txt         # Python package dependencies
├── .gitignore               # Git untracked file configurations
└── README.md                # Project documentation
```

---

## Getting Started

### Prerequisites

- Python 3.10 or higher installed.

### Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/Bhatta2006/DocuTeX.git
   cd DocuTeX
   ```

2. Create a virtual environment and activate it:
   ```bash
   # Windows
   python -m venv .venv
   .venv\Scripts\activate

   # macOS/Linux
   python3 -m venv .venv
   source .venv/bin/activate
   ```

3. Install the dependencies:
   ```bash
   pip install -r requirements.txt
   ```

### Running the Web Server

Start the development server using Uvicorn:

```bash
uvicorn app:app --reload
```

Once running, navigate to `http://127.0.0.1:8000` in your web browser.

### Running Tests

To run the integration tests:

```bash
pytest
```

---

## License

This project is open-source and available under the [MIT License](LICENSE).
