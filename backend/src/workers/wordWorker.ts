import { spawn } from 'child_process'
import path       from 'path'

const SCRIPT  = path.resolve(process.cwd(), 'scripts', 'wordAccessibilityCheck.ps1')
const TIMEOUT = Number(process.env.WORD_WORKER_TIMEOUT_MS ?? 45_000)
const ENABLED = process.env.WORD_WORKER_ENABLED === 'true'
const DEBUG   = process.env.WORD_WORKER_DEBUG === 'true'

// Single-slot concurrency queue — Word COM is not thread-safe
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

export interface WordWorkerSuccess {
  ok: true
  /** Always 'word-ui-automation' */
  engine: string
  wordVersion: string
  pageCount?: number | null
  /** idMso that successfully opened the pane */
  executedMso?: string
  /**
   * Accessibility issue occurrence counts.
   * Keys use final officeLikeSummary names directly (contrast, not hardToReadText).
   * Only keys with count > 0 are present (keys with 0 occurrences are omitted).
   */
  counts: Record<string, number>
  /**
   * Explicit pass/fail/manual status for checked keys.
   * Keys: missingAltText | contrast | tableHeader | mergedCells |
   *       unclearHyperlinkText | documentTitle | restrictedAccess
   */
  statuses: Record<string, string>
  /** Raw UIAutomation text collected from the pane (capped at 200 entries) */
  rawOfficeText: string[]
  // Debug-only fields — present when WORD_WORKER_DEBUG=true
  openedFilePath?: string
  openedFileName?: string
  verifiedWindowTitle?: string | null
  pollCount?: number
  parseTrace?: object[]
}

export interface WordWorkerFailure {
  ok: false
  error: string
}

export type WordWorkerOutput = WordWorkerSuccess | WordWorkerFailure

// ── Public API ────────────────────────────────────────────────────────────────

/**
 * Run the Word accessibility worker for the given file.
 * Always resolves — never rejects. Returns WordWorkerFailure on any error.
 */
export async function runWordWorker(filePath: string): Promise<WordWorkerOutput> {
  if (!ENABLED) {
    return { ok: false, error: 'WORD_WORKER_ENABLED is not set to true' }
  }

  await acquireSlot()
  try {
    return await spawnWorker(filePath)
  } finally {
    releaseSlot()
  }
}

// ── Internal ──────────────────────────────────────────────────────────────────

function spawnWorker(filePath: string): Promise<WordWorkerOutput> {
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
        // Child inherits all env vars, including WORD_WORKER_DEBUG
        env: { ...process.env },
      },
    )

    let stdout   = ''
    let stderr   = ''
    let timedOut = false

    const timer = setTimeout(() => {
      timedOut = true
      try { ps.kill() } catch { /* ignore */ }
      resolve({ ok: false, error: `Word worker timed out after ${TIMEOUT}ms` })
    }, TIMEOUT)

    ps.stdout.on('data', (chunk: Buffer) => { stdout += chunk.toString() })

    ps.stderr.on('data', (chunk: Buffer) => {
      const text = chunk.toString()
      stderr += text
      if (DEBUG) process.stderr.write(`[WORD WORKER] ${text}`)
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
        const parsed = JSON.parse(raw) as WordWorkerOutput
        if (DEBUG && parsed.ok) {
          const s = parsed as WordWorkerSuccess
          console.error(`[WORD WORKER] Parsed OK — engine=${s.engine} counts=${JSON.stringify(s.counts)}`)
          if (s.rawOfficeText?.length) {
            console.error(`[WORD WORKER] rawOfficeText (first 20):`, s.rawOfficeText.slice(0, 20))
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
