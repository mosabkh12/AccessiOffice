import express from 'express'
import cors from 'cors'
import dotenv from 'dotenv'
import scanRoutes, { scanErrorHandler } from './routes/scan.routes.js'

dotenv.config()

const app = express()
const PORT = Number(process.env.PORT) || 4000

app.use(cors())
app.use(express.json())

app.get('/api/health', (_req, res) => {
  res.json({ status: 'ok', service: 'AccessiOffice API' })
})

app.use('/api', scanRoutes)
app.use(scanErrorHandler)

import { PPTX_SCANNER_BUILD } from './services/pptxScanner.service.js'

app.listen(PORT, () => {
  console.log(`AccessiOffice API running on http://localhost:${PORT}`)
  console.log(`[AccessiOffice] Backend started with PPTX scanner version: ${PPTX_SCANNER_BUILD}`)
  if (process.env.EXCEL_WORKER_DEBUG === 'true') {
    console.log(`[EXCEL WORKER] enabled=${process.env.EXCEL_WORKER_ENABLED ?? 'unset'} timeout=${process.env.EXCEL_WORKER_TIMEOUT_MS ?? '45000'}ms`)
  }
})
