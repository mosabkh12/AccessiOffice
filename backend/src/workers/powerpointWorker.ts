import { spawn } from 'child_process'
import path       from 'path'
import { fileURLToPath } from 'url'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const SCRIPT   = path.resolve(__dirname, '../../scripts/pptxAccessibilityCheck.ps1')
const enabled  = () => process.env.PPTX_WORKER_ENABLED   === 'true'
const timeout  = () => Number(process.env.PPTX_WORKER_TIMEOUT_MS ?? 45_000)
const isDebug  = () => process.env.PPTX_WORKER_DEBUG      === 'true'

// Single-slot concurrency queue — PowerPoint COM is not thread-safe
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

/** A single occurrence extracted by invoking pane elements + reading COM state. */
export interface PptxOccurrence {
  index: number
  key: string
  slideNumber: number
  objectName: string
  shapeId: number
  shapeType?: string
  /** Hebrew display string: "שקופית N · ShapeName" */
  location: string
  source: 'Microsoft PowerPoint Accessibility Checker'
}

export interface WorkerSuccess {
  ok: true
  /** Always 'powerpoint-ui-automation' */
  engine: string
  pptVersion: string
  slideCount?: number
  /** idMso that successfully opened the pane */
  executedMso?: string
  /** Accessibility issue occurrence counts (only keys with count > 0 are present) */
  counts: Record<string, number>
  /**
   * Explicit pass/fail status for checks that PowerPoint actively tests.
   * Keys: contrast | mediaCaptions | missingTableHeader | mergedCells
   */
  statuses: Record<string, string>
  /** Structured per-occurrence data extracted via UIAutomation invoke + COM read. */
  occurrences?: Record<string, PptxOccurrence[]>
  /** true when at least one category yielded invokable occurrence elements. */
  occurrencesExtracted?: boolean
  /** Explanation when occurrencesExtracted is false. */
  occurrencesNote?: string | null
  /** Raw UIAutomation text collected from the pane (capped at 200 entries) */
  rawOfficeText: string[]
}

export interface WorkerFailure {
  ok: false
  error: string
}

export type WorkerOutput = WorkerSuccess | WorkerFailure

// ── Public API ────────────────────────────────────────────────────────────────

/**
 * Run the PowerPoint accessibility worker for the given file.
 * Always resolves — never rejects. Returns WorkerFailure on any error.
 */
export async function runPptxWorker(filePath: string): Promise<WorkerOutput> {
  if (!enabled()) {
    return { ok: false, error: 'PPTX_WORKER_ENABLED is not set to true' }
  }

  await acquireSlot()
  try {
    return await spawnWorker(filePath)
  } finally {
    releaseSlot()
  }
}

// ── Internal ──────────────────────────────────────────────────────────────────

function spawnWorker(filePath: string): Promise<WorkerOutput> {
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
        // Child inherits all env vars, including PPTX_WORKER_DEBUG
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
      resolve({ ok: false, error: `PowerPoint worker timed out after ${ms}ms` })
    }, ms)

    ps.stdout.on('data', (chunk: Buffer) => { stdout += chunk.toString() })

    ps.stderr.on('data', (chunk: Buffer) => {
      const text = chunk.toString()
      stderr += text
      // Forward PowerShell debug messages to Node stderr when debug is on
      if (isDebug()) process.stderr.write(`[PPT WORKER] ${text}`)
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
        const parsed = JSON.parse(raw) as WorkerOutput
        if (isDebug() && parsed.ok) {
          const s = parsed as WorkerSuccess
          console.error(`[PPT WORKER] Parsed OK — engine=${s.engine} counts=${JSON.stringify(s.counts)}`)
          if (s.rawOfficeText?.length) {
            console.error(`[PPT WORKER] rawOfficeText (first 20):`, s.rawOfficeText.slice(0, 20))
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
