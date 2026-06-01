import { useState, useRef } from 'react'
import { useNavigate } from 'react-router-dom'
import Button from '../components/Button.jsx'
import { useScan } from '../context/ScanContext.jsx'
import { getFileType, safeFileName } from '../utils/api.js'

const FILE_BADGE     = { docx: 'Word', pptx: 'PowerPoint', xlsx: 'Excel' }
const ACCEPTED       = '.docx,.pptx,.xlsx'
const WORKFLOW_STEPS = ['בחירת קובץ', 'הגדרת סריקה', 'הרצת בדיקה']

// Fixed values sent to the API — not selectable by the user
const DEFAULT_USER_TYPE = 'auditor'   // → 'accessibility-auditor' via USER_TYPE_TO_API
const DEFAULT_SCAN_TYPE = 'full'

export default function UploadPage() {
  const navigate = useNavigate()
  const { setScanData } = useScan()
  const inputRef = useRef(null)

  const [file, setFile]         = useState(null)
  const [dragOver, setDragOver] = useState(false)
  const [error, setError]       = useState('')

  const fileType     = file ? getFileType(file.name) : null
  // Step 1 while no file; once a file is chosen settings are pre-set so jump to step 3
  const workflowStep = file ? 3 : 1

  function handleFile(selected) {
    setError('')
    setScanData(null)
    if (!selected) return

    const type = getFileType(selected.name)
    if (!type) {
      setError('סוג קובץ לא נתמך. העלו קובץ .docx, .pptx או .xlsx בלבד.')
      setFile(null)
      return
    }
    setFile(selected)
  }

  function onDrop(e) {
    e.preventDefault()
    setDragOver(false)
    handleFile(e.dataTransfer.files[0])
  }

  function onSubmit(e) {
    e.preventDefault()
    if (!file) {
      setError('יש לבחור קובץ לפני הסריקה.')
      return
    }
    setScanData({
      file,
      fileName: file.name,
      fileType: getFileType(file.name),
      userType: DEFAULT_USER_TYPE,
      scanType: DEFAULT_SCAN_TYPE,
      results:  null,
    })
    navigate('/scan')
  }

  return (
    <div className="page page-upload">
      <header className="page-header">
        <h1>התחלת סריקת נגישות</h1>
        <p className="page-lead">העלו קובץ Office והמערכת תפעיל את בדיקת הנגישות המתאימה אוטומטית.</p>
      </header>

      <ol className="workflow-indicator" aria-label="שלבי תהליך">
        {WORKFLOW_STEPS.map((label, i) => {
          const stepNum = i + 1
          const active  = workflowStep === stepNum
          const done    = workflowStep > stepNum
          return (
            <li
              key={label}
              className={`workflow-indicator__item${active ? ' is-active' : ''}${done ? ' is-done' : ''}`}
            >
              <span className="workflow-indicator__num">{done ? '✓' : stepNum}</span>
              <span>{label}</span>
            </li>
          )
        })}
      </ol>

      <form onSubmit={onSubmit} className="upload-form">

        {/* ── Step 1: File selection ── */}
        <section className="card">
          <h2 className="card-title">1. בחירת קובץ</h2>

          <div
            className={`upload-zone${dragOver ? ' drag-over' : ''}${file ? ' has-file' : ''}`}
            onDragOver={(e) => { e.preventDefault(); setDragOver(true) }}
            onDragLeave={() => setDragOver(false)}
            onDrop={onDrop}
            onClick={() => inputRef.current?.click()}
            role="button"
            tabIndex={0}
            onKeyDown={(e) => e.key === 'Enter' && inputRef.current?.click()}
            aria-label="אזור העלאת קובץ"
          >
            <div className="upload-zone__icon" aria-hidden="true">⬆</div>
            {file ? (
              <div className="upload-zone__selected">
                <p className="upload-zone__status">הקובץ נבחר בהצלחה</p>
                <p className="file-name-display">{safeFileName(file.name, fileType)}</p>
                {fileType && (
                  <span className={`file-badge file-badge--${fileType}`}>
                    {FILE_BADGE[fileType]}
                  </span>
                )}
              </div>
            ) : (
              <>
                <p className="upload-zone__title">גררו קובץ לכאן או לחצו לבחירה</p>
                <p className="upload-zone__hint">פורמטים נתמכים: .docx · .pptx · .xlsx</p>
              </>
            )}
            <input
              ref={inputRef}
              type="file"
              accept={ACCEPTED}
              hidden
              onChange={(e) => handleFile(e.target.files[0])}
            />
          </div>

          {error && <p className="form-error" role="alert">{error}</p>}
        </section>

        {/* ── Step 2: Scan settings — fixed, non-interactive ── */}
        <section className="card">
          <h2 className="card-title">2. הגדרת סריקה</h2>

          <div className="form-section">
            <h3 className="form-label">סוג משתמש</h3>
            <div className="option-cards">
              <div className="option-card is-selected" style={{ cursor: 'default', pointerEvents: 'none' }}>
                <span className="option-card__title">בודק/ת נגישות</span>
                <span className="option-card__desc">דוח מפורט לבדיקות נגישות</span>
              </div>
            </div>
          </div>

          <div className="form-section">
            <h3 className="form-label">סוג סריקה</h3>
            <div className="option-cards">
              <div className="option-card is-selected" style={{ cursor: 'default', pointerEvents: 'none' }}>
                <span className="option-card__title">סריקה מלאה</span>
                <span className="option-card__desc">בדיקה מורחבת הכוללת ממצאים נוספים לפי סוג הקובץ</span>
              </div>
            </div>
          </div>
        </section>

        {/* ── Step 3: Run scan ── */}
        <section className="card card--actions">
          <h2 className="card-title">3. הרצת בדיקה</h2>
          <p className="card-hint">לאחר לחיצה, המערכת תעלה את הקובץ לשרת ותציג את תוצאות הסריקה.</p>
          <div className="btn-group">
            <Button type="submit" disabled={!file}>הפעלת סריקה</Button>
            <Button to="/" variant="secondary">חזרה לדף הבית</Button>
          </div>
        </section>

      </form>
    </div>
  )
}
