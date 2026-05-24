import { useEffect, useState, useRef } from 'react'
import { useNavigate } from 'react-router-dom'
import Button from '../components/Button.jsx'
import { useScan } from '../context/ScanContext.jsx'
import { SCAN_STEPS } from '../data/scanSteps.js'
import { submitScan, safeFileName } from '../utils/api.js'

const STEP_DELAY = 750
const FILE_BADGE = { docx: 'Word', pptx: 'PowerPoint', xlsx: 'Excel' }

export default function ScanProgressPage() {
  const navigate = useNavigate()
  const { scanData, setScanData } = useScan()
  const [currentStep, setCurrentStep] = useState(0)
  const [error, setError] = useState('')
  const apiDone = useRef(false)
  const scanFileKey = useRef(null)

  useEffect(() => {
    if (!scanData?.file) {
      navigate('/upload')
      return
    }

    const fileKey = `${scanData.file.name}-${scanData.file.size}-${scanData.file.lastModified}`
    if (fileKey !== scanFileKey.current) {
      scanFileKey.current = fileKey
      apiDone.current = false
      setCurrentStep(0)
      setError('')
      setScanData((prev) => (prev ? { ...prev, results: null } : prev))
    }

    if (!apiDone.current) {
      submitScan({
        file: scanData.file,
        userType: scanData.userType,
        scanType: scanData.scanType,
      })
        .then((results) => {
          apiDone.current = true
          setScanData((prev) => ({ ...prev, results }))
        })
        .catch((err) => {
          setError(err.message || 'הסריקה נכשלה. נסו שוב.')
        })
    }
  }, [scanData, navigate, setScanData])

  useEffect(() => {
    if (error || !scanData?.fileName) return

    if (currentStep < SCAN_STEPS.length) {
      const timer = setTimeout(() => setCurrentStep((s) => s + 1), STEP_DELAY)
      return () => clearTimeout(timer)
    }

    if (scanData.results?.scanId) {
      const timer = setTimeout(() => navigate('/results'), 500)
      return () => clearTimeout(timer)
    }
  }, [currentStep, scanData, error, navigate])

  if (!scanData?.fileName) return null

  const displayName = safeFileName(scanData.fileName, scanData.fileType)
  const fileLabel = FILE_BADGE[scanData.fileType] || scanData.fileType
  const progress = Math.min(
    100,
    Math.round(((currentStep + (scanData.results ? 1 : 0)) / (SCAN_STEPS.length + 1)) * 100),
  )

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

        <div className="progress-bar-wrap progress-bar-wrap--animated" role="progressbar" aria-valuenow={progress} aria-valuemin={0} aria-valuemax={100}>
          <div className="progress-bar" style={{ width: `${progress}%` }} />
        </div>
        <p className="scan-progress-label">{progress}% הושלם</p>

        <ol className="progress-steps progress-steps--scan">
          {SCAN_STEPS.map((label, index) => {
            const done = index < currentStep || (scanData.results && index <= currentStep)
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
      </div>
    </div>
  )
}
