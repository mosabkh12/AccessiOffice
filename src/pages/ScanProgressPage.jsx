import { useEffect, useState, useRef } from 'react'
import { useNavigate } from 'react-router-dom'
import Button from '../components/Button.jsx'
import { useScan } from '../context/ScanContext.jsx'
import { SCAN_STEPS } from '../data/scanSteps.js'
import { submitScan, safeFileName } from '../utils/api.js'

const STEP_DELAY = 750
const FILE_BADGE = { docx: 'Word', pptx: 'PowerPoint', xlsx: 'Excel' }

// Extra step label shown only for PPTX — reflects the PowerPoint worker phase
const PPTX_WORKER_STEP = 'בדיקה באמצעות Microsoft PowerPoint Accessibility Checker'

export default function ScanProgressPage() {
  const navigate = useNavigate()
  const { scanData, setScanData } = useScan()
  const [currentStep, setCurrentStep] = useState(0)
  const [error, setError] = useState('')

  // submitted.current prevents duplicate API calls when React re-renders the effect.
  // Set to true BEFORE the async call so any subsequent effect run sees it as locked.
  const submitted   = useRef(false)
  const scanFileKey = useRef(null)

  const isPptx = scanData?.fileType === 'pptx'
  const steps  = isPptx ? [...SCAN_STEPS, PPTX_WORKER_STEP] : SCAN_STEPS

  // ── Main scan effect ──────────────────────────────────────────────────────────
  useEffect(() => {
    if (!scanData?.file) {
      navigate('/upload')
      return
    }

    const fileKey = `${scanData.file.name}-${scanData.file.size}-${scanData.file.lastModified}`

    // New file detected → reset state for this scan
    if (fileKey !== scanFileKey.current) {
      scanFileKey.current = fileKey
      submitted.current   = false
      setCurrentStep(0)
      setError('')
      // Reset results without creating a new object when already null (avoids an
      // extra state-change that would re-trigger this effect and cause duplicate calls)
      setScanData((prev) => {
        if (!prev || prev.results === null) return prev
        return { ...prev, results: null }
      })
    }

    // Guard: lock immediately before the async call so re-renders don't duplicate it
    if (!submitted.current) {
      submitted.current = true
      submitScan({
        file: scanData.file,
        userType: scanData.userType,
        scanType: scanData.scanType,
      })
        .then((results) => {
          setScanData((prev) => ({ ...prev, results }))
        })
        .catch((err) => {
          setError(err.message || 'הסריקה נכשלה. נסו שוב.')
        })
    }
  }, [scanData, navigate, setScanData])

  // ── Navigate to results once steps done AND API returned ─────────────────────
  useEffect(() => {
    if (error || !scanData?.results || currentStep < steps.length) return
    const timer = setTimeout(() => navigate('/results'), 600)
    return () => clearTimeout(timer)
  }, [currentStep, scanData?.results, error, navigate, steps.length])

  // ── Animate progress steps ────────────────────────────────────────────────────
  useEffect(() => {
    if (error || !scanData?.fileName) return
    if (currentStep < steps.length) {
      const timer = setTimeout(() => setCurrentStep((s) => s + 1), STEP_DELAY)
      return () => clearTimeout(timer)
    }
  }, [currentStep, scanData?.fileName, error, steps.length])

  if (!scanData?.fileName) return null

  const displayName = safeFileName(scanData.fileName, scanData.fileType)
  const fileLabel   = FILE_BADGE[scanData.fileType] || scanData.fileType
  const progress    = Math.min(
    100,
    Math.round(((currentStep + (scanData.results ? 1 : 0)) / (steps.length + 1)) * 100),
  )

  // True when all steps animated but API still running (common for PPTX + worker)
  const waitingForWorker = currentStep >= steps.length && !scanData?.results

  if (error) {
    return (
      <div className="page page-scan">
        <header className="page-header">
          <h1>שגיאה בסריקה</h1>
        </header>
        <div className="card card--error" role="alert">
          <p className="form-error">{error}</p>
          <p className="page-lead">
            ודאו ששרת ה-API פועל (<span className="ltr-inline" dir="ltr">localhost:4000</span>)
            והקובץ הוא docx, pptx או xlsx.
          </p>
          <div className="btn-group">
            <Button to="/upload">חזרה להעלאה</Button>
          </div>
        </div>
      </div>
    )
  }

  return (
    <div className="page page-scan">
      <header className="page-header">
        <h1>המערכת סורקת את הקובץ</h1>
        <p className="page-lead">אנא המתינו בזמן שהמערכת מנתחת את המסמך.</p>
      </header>

      <div className="card scan-card">
        <div className="scan-file-info">
          <p><strong>קובץ:</strong> <span className="file-name-display">{displayName}</span></p>
          <span className={`file-badge file-badge--${scanData.fileType}`}>{fileLabel}</span>
        </div>

        <div
          className="progress-bar-wrap progress-bar-wrap--animated"
          role="progressbar"
          aria-valuenow={progress}
          aria-valuemin={0}
          aria-valuemax={100}
        >
          <div className="progress-bar" style={{ width: `${progress}%` }} />
        </div>
        <p className="scan-progress-label">{progress}% הושלם</p>

        <ol className="progress-steps progress-steps--scan">
          {steps.map((label, index) => {
            const done   = index < currentStep || (scanData.results && index <= currentStep)
            const active = index === currentStep && !scanData.results
            return (
              <li
                key={label}
                className={`progress-step${done ? ' done' : ''}${active ? ' active' : ''}`}
              >
                <span className="step-indicator" aria-hidden="true">
                  {done ? '✓' : index + 1}
                </span>
                <span>{label}</span>
              </li>
            )
          })}
        </ol>

        {/* Shown when animation is done but the PowerPoint worker is still running */}
        {waitingForWorker && isPptx && (
          <div className="scan-ppt-waiting" role="status" aria-live="polite">
            <span className="scan-ppt-waiting__icon" aria-hidden="true">⏳</span>
            <span className="scan-ppt-waiting__text">
              בודק באמצעות Microsoft PowerPoint Accessibility Checker...
              <br />
              <small>שלב זה עשוי לקחת עד 45 שניות.</small>
            </span>
          </div>
        )}
      </div>
    </div>
  )
}
