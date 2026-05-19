import { Router, type Request, type Response, type NextFunction } from 'express'
import multer from 'multer'
import path from 'path'
import fs from 'fs'
import { fileURLToPath } from 'url'
import {
  generateScanResult,
  getFileType,
  parseScanType,
  parseUserType,
} from '../services/scan.service.js'
import { resolveDisplayFileName } from '../utils/filename.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const uploadsDir = path.join(__dirname, '../../uploads')

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

router.post('/scan', upload.single('file'), (req, res) => {
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
  const result = generateScanResult(displayName, fileType, userType, scanType)
  cleanup(uploaded.path)
  res.json(result)
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
