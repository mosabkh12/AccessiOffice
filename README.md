# AccessiOffice

Office Accessibility Checker – React + Vite frontend with Node.js + Express API.

## Prerequisites

- Node.js 18+

## Setup

### 1. Frontend

```bash
npm install
cp .env.example .env
```

### 2. Backend

```bash
cd backend
npm install
cp .env.example .env
cd ..
```

## Run locally

**Option A – both servers (recommended):**

```bash
npm install
cd backend && npm install && cd ..
npm run dev:all
```

**Option B – two terminals:**

```bash
# Terminal 1 – API on http://localhost:4000
cd backend
npm run dev

# Terminal 2 – frontend on http://localhost:5173
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

## Project structure

```
├── src/                 # React frontend
├── backend/
│   ├── src/
│   │   ├── index.ts
│   │   ├── routes/scan.routes.ts
│   │   ├── services/scan.service.ts
│   │   └── types/scan.types.ts
│   └── uploads/
└── package.json
```

## Stack

- **Frontend:** React 19, React Router 7, Vite 6, Hebrew RTL
- **Backend:** Node.js, Express 5, TypeScript, Multer, CORS
