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
import { runPptxWorker } from '../workers/powerpointWorker.js'
import { runWordWorker } from '../workers/wordWorker.js'
import { runExcelWorker } from '../workers/excelWorker.js'
import {
  mergeWorkerIntoResult,
  markWorkerFallback,
  mergeWordWorkerIntoResult,
  markWordWorkerFallback,
  mergeExcelWorkerIntoResult,
  markExcelWorkerFallback,
} from '../services/officeEngine.service.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const uploadsDir = path.join(__dirname, '../../uploads')
const wordTmpDir = path.join(__dirname, '../../tmp')
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
const PPTX_WORKER_DEBUG = process.env.PPTX_WORKER_DEBUG === 'true'
const WORD_WORKER_DEBUG  = process.env.WORD_WORKER_DEBUG  === 'true'
const EXCEL_WORKER_DEBUG = process.env.EXCEL_WORKER_DEBUG === 'true'

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

    // Always compute SHA-256: lets callers verify the scanned file changed after edits.
    const fileBuffer = fs.readFileSync(uploaded.path)
    const fileHash = crypto.createHash('sha256').update(fileBuffer).digest('hex')

    if (DEBUG_SCAN) {
      console.log('[SCAN REQUEST]', {
        originalFilename: uploaded.originalname,
        displayFileName: displayName,
        fileSize,
        extension: ext,
        timestamp: scanId,
        fileHash,
      })
    }

    const xmlResult = await generateScanResult(
      uploaded.path,
      displayName,
      fileType,
      userType,
      scanType,
      fileSize,
    )

    // ── PowerPoint worker (PPTX only, opt-in via PPTX_WORKER_ENABLED) ──────
    let finalResult: ScanResult = xmlResult
    let workerRaw: {
      executedMso?: string
      counts: Record<string, number>
      statuses: Record<string, string>
      rawOfficeText: string[]
    } | null = null

    if (fileType === 'pptx' && process.env.PPTX_WORKER_ENABLED === 'true') {
      const workerOut = await runPptxWorker(uploaded.path)
      if (workerOut.ok) {
        finalResult = mergeWorkerIntoResult(xmlResult, workerOut)
        if (PPTX_WORKER_DEBUG) {
          workerRaw = {
            executedMso: workerOut.executedMso,
            counts: workerOut.counts,
            statuses: workerOut.statuses,
            rawOfficeText: workerOut.rawOfficeText,
          }
        }
        if (DEBUG_SCAN) {
          console.log('[PPT WORKER] Success — engine: office-engine', {
            pptVersion: workerOut.pptVersion,
            counts: workerOut.counts,
          })
        }
      } else {
        finalResult = markWorkerFallback(xmlResult, workerOut.error)
        console.warn('[PPT WORKER] Falling back to XML scanner:', workerOut.error)
      }
    }

    // ── Word worker (DOCX only, opt-in via WORD_WORKER_ENABLED) ─────────────
    let wordWorkerRaw: {
      executedMso?: string
      counts: Record<string, number>
      statuses: Record<string, string>
      rawOfficeText: string[]
      openedFilePath?: string
      openedFileName?: string
      verifiedWindowTitle?: string | null
      pollCount?: number
      parseTrace?: object[]
    } | null = null

    if (fileType === 'docx' && process.env.WORD_WORKER_ENABLED === 'true') {
      // Debug: save a copy of the uploaded DOCX so it can be tested manually
      // with the same file the worker will receive, even after cleanup.
      if (WORD_WORKER_DEBUG) {
        try {
          if (!fs.existsSync(wordTmpDir)) fs.mkdirSync(wordTmpDir, { recursive: true })
          const debugDest = path.join(wordTmpDir, 'debug-last-word-upload.docx')
          fs.copyFileSync(uploaded.path, debugDest)
          console.log(`[WORD DEBUG] Saved debug copy → ${debugDest}`)
          console.log(`[WORD DEBUG] Manual test: powershell -NonInteractive -NoProfile -ExecutionPolicy Bypass -File .\\backend\\scripts\\wordAccessibilityCheck.ps1 -FilePath "${debugDest}"`)
        } catch { /* best-effort */ }
      }

      const workerOut = await runWordWorker(uploaded.path)
      if (workerOut.ok) {
        finalResult = mergeWordWorkerIntoResult(xmlResult, workerOut)
        if (WORD_WORKER_DEBUG) {
          wordWorkerRaw = {
            executedMso:          workerOut.executedMso,
            counts:               workerOut.counts,
            statuses:             workerOut.statuses,
            rawOfficeText:        workerOut.rawOfficeText,
            openedFilePath:       workerOut.openedFilePath,
            openedFileName:       workerOut.openedFileName,
            verifiedWindowTitle:  workerOut.verifiedWindowTitle,
            pollCount:            workerOut.pollCount,
            parseTrace:           workerOut.parseTrace,
          }
        }
        if (DEBUG_SCAN) {
          console.log('[WORD WORKER] Success — engine: office-engine', {
            wordVersion: workerOut.wordVersion,
            counts: workerOut.counts,
          })
        }
      } else {
        finalResult = markWordWorkerFallback(xmlResult, workerOut.error)
        console.warn('[WORD WORKER] Falling back to XML scanner:', workerOut.error)
      }
    }

    // ── Excel worker (XLSX only, opt-in via EXCEL_WORKER_ENABLED) ──────────────
    let excelWorkerRaw: {
      executedMso?: string
      counts: Record<string, number>
      statuses: Record<string, string>
      rawOfficeText: string[]
      openedFilePath?: string
      openedFileName?: string
      verifiedWindowTitle?: string | null
      pollCount?: number
      parseTrace?: object[]
    } | null = null

    if (EXCEL_WORKER_DEBUG) {
      console.log('[EXCEL WORKER] Request received', {
        fileType,
        EXCEL_WORKER_ENABLED: process.env.EXCEL_WORKER_ENABLED,
        tempFilePath: uploaded.path,
      })
    }

    if (fileType === 'xlsx' && process.env.EXCEL_WORKER_ENABLED === 'true') {
      // Debug: save a copy of the uploaded XLSX so it can be tested manually
      // with the same file the worker will receive, even after cleanup.
      if (EXCEL_WORKER_DEBUG) {
        try {
          if (!fs.existsSync(wordTmpDir)) fs.mkdirSync(wordTmpDir, { recursive: true })
          const debugDest = path.join(wordTmpDir, 'debug-last-excel-upload.xlsx')
          fs.copyFileSync(uploaded.path, debugDest)
          console.log(`[EXCEL DEBUG] Saved debug copy -> ${debugDest}`)
          console.log(`[EXCEL DEBUG] Manual test: powershell -NonInteractive -NoProfile -ExecutionPolicy Bypass -File .\\backend\\scripts\\excelAccessibilityCheck.ps1 -FilePath "${debugDest}"`)
        } catch { /* best-effort */ }
      }

      const workerOut = await runExcelWorker(uploaded.path)

      if (EXCEL_WORKER_DEBUG) {
        console.log('[EXCEL WORKER] Result', {
          ok: workerOut.ok,
          ...(workerOut.ok
            ? { counts: workerOut.counts, excelVersion: workerOut.excelVersion }
            : { error: workerOut.error }),
        })
      }

      if (workerOut.ok) {
        finalResult = mergeExcelWorkerIntoResult(xmlResult, workerOut)

        if (EXCEL_WORKER_DEBUG) {
          excelWorkerRaw = {
            executedMso:         workerOut.executedMso,
            counts:              workerOut.counts,
            statuses:            workerOut.statuses,
            rawOfficeText:       workerOut.rawOfficeText,
            openedFilePath:      workerOut.openedFilePath,
            openedFileName:      workerOut.openedFileName,
            verifiedWindowTitle: workerOut.verifiedWindowTitle,
            pollCount:           workerOut.pollCount,
            parseTrace:          workerOut.parseTrace,
          }
          console.log('[EXCEL WORKER] officeLikeSummary.mergedCells', {
            count:  finalResult.officeLikeSummary?.mergedCells?.count,
            status: finalResult.officeLikeSummary?.mergedCells?.status,
          })
        }

        if (DEBUG_SCAN) {
          console.log('[EXCEL WORKER] Success — engine: office-engine', {
            excelVersion: workerOut.excelVersion,
            counts:       workerOut.counts,
          })
        }
      } else {
        finalResult = markExcelWorkerFallback(xmlResult, workerOut.error)
        console.warn('[EXCEL WORKER] Falling back to XML scanner:', workerOut.error)
      }
    }

    const responseBody = {
      ...ensurePptxDiagnosticsResponse(finalResult),
      fileHash,
      ...(workerRaw      ? { powerpointWorkerRaw: workerRaw }  : {}),
      ...(wordWorkerRaw  ? { wordWorkerRaw }                   : {}),
      ...(excelWorkerRaw ? { excelWorkerRaw }                  : {}),
    }
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
