import { Router, type Request, type Response, type NextFunction } from 'express'
import multer from 'multer'
import path from 'path'
import fs from 'fs'
import crypto from 'crypto'
import { fileURLToPath } from 'url'
import {
  generateScanResult,
  getFileType,
  parseScanType,
  parseUserType,
} from '../services/scan.service.js'
import { resolveDisplayFileName } from '../utils/filename.js'
import type { ScanResult } from '../types/scan.types.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const uploadsDir = path.join(__dirname, '../../uploads')
const DEBUG_SCAN = process.env.ACCESSIOFFICE_DEBUG_SCAN === 'true'

if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir, { recursive: true })
}

const storage = multer.diskStorage({
  destination: (_req, _file, cb) => cb(null, uploadsDir),
  filename: (_req, file, cb) => {
    const unique = `${Date.now()}-${Math.round(Math.random() * 1e6)}`
    cb(null, `${unique}${path.extname(file.originalname)}`)
  },
})

const upload = multer({
  storage,
  limits: { fileSize: 25 * 1024 * 1024 },
  fileFilter: (_req, file, cb) => {
    const type = getFileType(file.originalname)
    if (!type) {
      cb(new Error('Unsupported file type. Please upload .docx, .pptx, or .xlsx only.'))
      return
    }
    cb(null, true)
  },
})

const router = Router()
const PPTX_RESPONSE_VERSION = 'pptx-diagnostics-v2'
const PPTX_RESPONSE_DEBUG_MARKER = 'API-SCAN-RESPONSE-V2-ACTIVE'

function ensurePptxDiagnosticsResponse(result: ScanResult): ScanResult {
  if (result.fileType !== 'pptx') return result
  const diagnostics = result.scanDiagnostics ?? result.diagnostics ?? {}
  return {
    ...result,
    scannerVersion: PPTX_RESPONSE_VERSION,
    ...(DEBUG_SCAN ? { debugMarker: PPTX_RESPONSE_DEBUG_MARKER } : {}),
    diagnostics,
    scanDiagnostics: diagnostics,
    checkStatuses: result.checkStatuses ?? [],
    manualReviewChecks: result.manualReviewChecks ?? [],
  }
}

router.post('/scan', upload.single('file'), async (req, res) => {
  const uploaded = req.file

  if (!uploaded) {
    res.status(400).json({ error: 'No file uploaded. Please attach an Office file.' })
    return
  }

  const userType = parseUserType(req.body.userType ?? '')
  const scanType = parseScanType(req.body.scanType ?? '')

  if (!userType) {
    cleanup(uploaded.path)
    res.status(400).json({
      error: 'Invalid userType. Use document-author, accessibility-auditor, or lecturer-institution.',
    })
    return
  }

  if (!scanType) {
    cleanup(uploaded.path)
    res.status(400).json({ error: 'Invalid scanType. Use basic or full.' })
    return
  }

  const fileType = getFileType(uploaded.originalname)
  if (!fileType) {
    cleanup(uploaded.path)
    res.status(400).json({ error: 'Unsupported file type. Please upload .docx, .pptx, or .xlsx only.' })
    return
  }

  const displayName = resolveDisplayFileName(
    uploaded.originalname,
    typeof req.body.displayFileName === 'string' ? req.body.displayFileName : undefined,
    fileType,
  )
  try {
    const fileSize = fs.statSync(uploaded.path).size
    const scanId = String(Date.now())
    const ext = path.extname(uploaded.originalname).toLowerCase()

    if (DEBUG_SCAN) {
      const fileBuffer = fs.readFileSync(uploaded.path)
      const debugFileHash = crypto.createHash('sha1').update(fileBuffer).digest('hex')
      console.log('[SCAN REQUEST]', {
        originalFilename: uploaded.originalname,
        displayFileName: displayName,
        savedPath: uploaded.path,
        fileSize,
        extension: ext,
        timestamp: scanId,
        debugFileHash,
      })
    }

    const result = await generateScanResult(
      uploaded.path,
      displayName,
      fileType,
      userType,
      scanType,
      fileSize,
    )
    const responseBody = ensurePptxDiagnosticsResponse(result)
    cleanup(uploaded.path)
    res.json(responseBody)
  } catch {
    cleanup(uploaded.path)
    res.status(500).json({ error: 'Failed to analyze the uploaded file.' })
  }
})

function cleanup(filePath: string) {
  fs.unlink(filePath, () => {})
}

export function scanErrorHandler(err: unknown, _req: Request, res: Response, _next: NextFunction) {
  if (err instanceof multer.MulterError) {
    res.status(400).json({ error: err.message })
    return
  }
  if (err instanceof Error) {
    res.status(400).json({ error: err.message })
    return
  }
  res.status(500).json({ error: 'An unexpected error occurred during scan.' })
}

export default router
