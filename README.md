# AccessiOffice

Office Accessibility Checker - React + Vite frontend with Node.js + Express API.

## Prerequisites

- Node.js 18+
- Windows OS (required for Office Engine workers)
- Microsoft Office 2016+ desktop edition (required for Office Engine workers)

## Setup

> Windows PowerShell note: if `npm` is blocked with an unsigned `npm.ps1`
> execution-policy error, use `npm.cmd` in the commands below, for example
> `npm.cmd install` and `npm.cmd run dev:all`.

### 1. Frontend

```bash
npm install
cp .env.example .env
```

### 2. Backend

```bash
cd backend
npm install
cd ..
```

The backend reads the same root `.env` file as the frontend. Do not create a
separate `backend/.env`.

## Run locally

**Option A - both servers (recommended):**

```bash
npm install
cd backend && npm install && cd ..
npm run dev:all
```

**Option B - two terminals:**

```bash
# Terminal 1 - API on http://localhost:4000
cd backend
npm run dev

# Terminal 2 - frontend on http://localhost:5173
npm run dev
```

## Usage

1. Open http://localhost:5173
2. Click **התחלת סריקת נגישות**
3. Upload a `.docx`, `.pptx`, or `.xlsx` file
4. Select user type and scan type, then **הפעלת סריקה**
5. View results and printable report

## API

`POST http://localhost:4000/api/scan`

- **Content-Type:** `multipart/form-data`
- **Fields:** `file`, `userType`, `scanType`
- **userType:** `document-author` | `accessibility-auditor` | `lecturer-institution`
- **scanType:** `basic` | `full`

Health check: `GET http://localhost:4000/api/health`

## Office Engine (Windows only)

AccessiOffice includes three optional UI-Automation workers that read results directly
from the real Microsoft Office Accessibility Checker pane, giving accurate counts
instead of XML heuristics.

### Requirements

- Windows 10/11
- Microsoft Office 2016 or later (desktop/Click-to-Run edition)
  - PowerPoint for PPTX files
  - Word for DOCX files
  - Excel for XLSX files
- The corresponding Office application must not already be open with an
  Accessibility pane from a previous session during a scan (the worker dismisses
  stale panes automatically, but concurrent manual use may interfere)

### Enabling workers

Set the following variables in `backend/.env` before starting the backend:

```
# PowerPoint worker
PPTX_WORKER_ENABLED=true
PPTX_WORKER_TIMEOUT_MS=45000
PPTX_WORKER_DEBUG=false

# Word worker
WORD_WORKER_ENABLED=true
WORD_WORKER_TIMEOUT_MS=45000
WORD_WORKER_DEBUG=false

# Excel worker
EXCEL_WORKER_ENABLED=true
EXCEL_WORKER_TIMEOUT_MS=45000
EXCEL_WORKER_DEBUG=false
```

Alternatively set them in the PowerShell session before `npm run dev`:

```powershell
$env:PPTX_WORKER_ENABLED  = "true"
$env:WORD_WORKER_ENABLED  = "true"
$env:EXCEL_WORKER_ENABLED = "true"
npm run dev
```

### How it works

1. The backend receives an uploaded file and runs the XML scanner (always).
2. If the matching worker is enabled, a PowerShell script opens the Office
   application via COM, triggers the Accessibility Checker pane via `ExecuteMso`,
   and reads the pane content through Windows UIAutomation.
3. The worker result is merged into the response as `officeLikeSummary`.
   `engine` is set to `"office-engine"` on success.
4. If the worker fails (Office not installed, timeout, COM error), the response
   falls back to the XML scanner result and sets `engine` to `"hybrid"`.
   A fallback warning is shown in the UI.

### Fallback (XML scanner)

When workers are disabled or fail, AccessiOffice falls back to an XML/WCAG
heuristic scanner that parses the raw Office file format. Results are less
precise than the real Accessibility Checker but require no Office installation
and work on any OS (including Linux/cloud deployments).

### Limitations

- Office Engine workers require Windows and a desktop Office installation.
  They are intended for a local or university demo environment, not cloud
  or containerised deployments.
- Only one scan per Office application can run at a time (COM is single-threaded).
  Concurrent uploads for the same file type are queued automatically.
- The Office application window briefly appears on screen during a scan.
  This is required for UIAutomation to read the Accessibility pane.
- Office 365 insider builds that host the Accessibility pane in WebView2 may
  prevent UIAutomation from reading the results. The worker falls back to XML
  automatically in that case.
- Debug copies of uploaded files are saved to `backend/tmp/` only when the
  corresponding `_DEBUG` env var is `true`. They are not created in normal mode.

## Project structure

```
.
+-- src/                          # React frontend
|   +-- pages/
|   |   +-- ResultsPage.jsx       # Main scan results
|   |   +-- ReportPage.jsx        # Printable report
|   +-- utils/
|       +-- officeLikeReport.js   # Office Engine helpers (score, rows, groups)
+-- backend/
|   +-- src/
|   |   +-- index.ts              # Express entry point + dotenv
|   |   +-- routes/scan.routes.ts # /api/scan endpoint
|   |   +-- services/
|   |   |   +-- officeEngine.service.ts  # Merge worker results into ScanResult
|   |   |   +-- scan.service.ts          # XML scanner orchestrator
|   |   +-- workers/
|   |   |   +-- powerpointWorker.ts      # PPTX PS1 wrapper
|   |   |   +-- wordWorker.ts            # DOCX PS1 wrapper
|   |   |   +-- excelWorker.ts           # XLSX PS1 wrapper
|   |   +-- types/scan.types.ts          # Shared TypeScript types
|   +-- scripts/
|   |   +-- pptxAccessibilityCheck.ps1   # PowerPoint COM + UIAutomation
|   |   +-- wordAccessibilityCheck.ps1   # Word COM + UIAutomation
|   |   +-- excelAccessibilityCheck.ps1  # Excel COM + UIAutomation
|   +-- uploads/                  # Temp upload directory (auto-deleted after scan)
|   +-- tmp/                      # Debug file copies (created only when _DEBUG=true)
+-- package.json
```

## Stack

- **Frontend:** React 19, React Router 7, Vite 6, Hebrew RTL
- **Backend:** Node.js, Express 5, TypeScript, Multer, CORS
- **Office Engine:** Windows COM + UIAutomation via PowerShell
