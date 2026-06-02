import { spawn } from 'child_process'
import path       from 'path'
import { fileURLToPath } from 'url'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const SCRIPT = path.resolve(__dirname, '../../scripts/excelAccessibilityCheck.ps1')

// Read all env vars dynamically at call time so that dotenv (loaded in index.ts
// after ES module imports are evaluated) can still influence these values.
const enabled  = () => process.env.EXCEL_WORKER_ENABLED   === 'true'
const timeout  = () => Number(process.env.EXCEL_WORKER_TIMEOUT_MS ?? 45_000)
const isDebug  = () => process.env.EXCEL_WORKER_DEBUG      === 'true'

// Single-slot concurrency queue — Excel COM is not thread-safe
let _busy = false
const _queue: Array<() => void> = []

function acquireSlot(): Promise<void> {
  return new Promise((resolve) => {
    if (!_busy) { _busy = true; resolve(); return }
    _queue.push(() => { _busy = true; resolve() })
  })
}

function releaseSlot(): void {
  _busy = false
  const next = _queue.shift()
  if (next) next()
}

// ── Public result types ───────────────────────────────────────────────────────

export interface ExcelWorkerSuccess {
  ok: true
  /** Always 'excel-ui-automation' */
  engine: string
  excelVersion: string
  sheetCount?: number | null
  /** idMso that successfully opened the pane */
  executedMso?: string
  /**
   * Accessibility issue occurrence counts.
   * Keys use final officeLikeSummary names directly.
   * Only keys with count > 0 are present (keys with 0 occurrences are omitted).
   */
  counts: Record<string, number>
  /**
   * Explicit pass/fail/manual status for checked keys.
   * Keys: missingAltText | contrast | tableHeader | mergedCells |
   *       unclearHyperlinkText | sheetName | restrictedAccess
   */
  statuses: Record<string, string>
  /** Per-occurrence items: structured objects from COM enumeration OR plain strings from text-scan. */
  occurrences?: Record<string, unknown[]>
  /** Raw UIAutomation text collected from the pane (capped at 200 entries) */
  rawOfficeText: string[]
  // Debug-only fields — present when EXCEL_WORKER_DEBUG=true
  openedFilePath?: string
  openedFileName?: string
  verifiedWindowTitle?: string | null
  pollCount?: number
  parseTrace?: object[]
}

export interface ExcelWorkerFailure {
  ok: false
  error: string
}

export type ExcelWorkerOutput = ExcelWorkerSuccess | ExcelWorkerFailure

// ── Public API ────────────────────────────────────────────────────────────────

/**
 * Run the Excel accessibility worker for the given file.
 * Always resolves — never rejects. Returns ExcelWorkerFailure on any error.
 */
export async function runExcelWorker(filePath: string): Promise<ExcelWorkerOutput> {
  if (!enabled()) {
    return { ok: false, error: 'EXCEL_WORKER_ENABLED is not set to true' }
  }

  await acquireSlot()
  try {
    return await spawnWorker(filePath)
  } finally {
    releaseSlot()
  }
}

// ── Internal ──────────────────────────────────────────────────────────────────

function spawnWorker(filePath: string): Promise<ExcelWorkerOutput> {
  return new Promise((resolve) => {
    const ps = spawn(
      'powershell.exe',
      [
        '-NonInteractive',
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', SCRIPT,
        '-FilePath', filePath,
      ],
      {
        windowsHide: true,
        // Child inherits all env vars, including EXCEL_WORKER_DEBUG
        env: { ...process.env },
      },
    )

    let stdout   = ''
    let stderr   = ''
    let timedOut = false

    const ms = timeout()
    const timer = setTimeout(() => {
      timedOut = true
      try { ps.kill() } catch { /* ignore */ }
      resolve({ ok: false, error: `Excel worker timed out after ${ms}ms` })
    }, ms)

    ps.stdout.on('data', (chunk: Buffer) => { stdout += chunk.toString() })

    ps.stderr.on('data', (chunk: Buffer) => {
      const text = chunk.toString()
      stderr += text
      if (isDebug()) process.stderr.write(`[EXCEL WORKER] ${text}`)
    })

    ps.on('close', () => {
      clearTimeout(timer)
      if (timedOut) return

      const raw = stdout.trim()
      if (!raw) {
        resolve({ ok: false, error: `Worker produced no stdout. stderr: ${stderr.slice(0, 600)}` })
        return
      }

      try {
        const parsed = JSON.parse(raw) as ExcelWorkerOutput
        if (isDebug() && parsed.ok) {
          const s = parsed as ExcelWorkerSuccess
          console.error(`[EXCEL WORKER] Parsed OK — engine=${s.engine} counts=${JSON.stringify(s.counts)}`)
          if (s.rawOfficeText?.length) {
            console.error(`[EXCEL WORKER] rawOfficeText (first 20):`, s.rawOfficeText.slice(0, 20))
          }
        }
        resolve(parsed)
      } catch {
        resolve({ ok: false, error: `Failed to parse worker JSON. stdout: ${raw.slice(0, 500)}` })
      }
    })

    ps.on('error', (err: Error) => {
      clearTimeout(timer)
      resolve({ ok: false, error: `Failed to spawn PowerShell: ${err.message}` })
    })
  })
}
