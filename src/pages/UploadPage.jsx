import { useState, useRef } from 'react'
import { useNavigate } from 'react-router-dom'
import Button from '../components/Button.jsx'
import { useScan } from '../context/ScanContext.jsx'
import { getFileType, safeFileName } from '../utils/api.js'

const USER_TYPES = [
  { value: 'author', label: 'מחבר/ת מסמך', desc: 'בדיקה לפני פרסום או שיתוף מסמך' },
  { value: 'auditor', label: 'בודק/ת נגישות', desc: 'דוח מפורט לביקורת נגישות' },
  { value: 'lecturer', label: 'מרצה / מוסד לימודי', desc: 'התאמה לחומרי הוראה ומצגות' },
]

const SCAN_TYPES = [
  {
    value: 'basic',
    label: 'סריקה בסיסית',
    desc: 'בדיקת בעיות נגישות נפוצות ומהירות.',
  },
  {
    value: 'full',
    label: 'סריקה מלאה',
    desc: 'בדיקה מורחבת הכוללת ממצאים נוספים לפי סוג הקובץ.',
  },
]

const FILE_BADGE = { docx: 'Word', pptx: 'PowerPoint', xlsx: 'Excel' }

const ACCEPTED = '.docx,.pptx,.xlsx'

const WORKFLOW_STEPS = ['בחירת קובץ', 'הגדרת סריקה', 'הרצת בדיקה']

export default function UploadPage() {
  const navigate = useNavigate()
  const { setScanData } = useScan()
  const inputRef = useRef(null)

  const [file, setFile] = useState(null)
  const [dragOver, setDragOver] = useState(false)
  const [userType, setUserType] = useState('author')
  const [scanType, setScanType] = useState('basic')
  const [error, setError] = useState('')
  const [workflowStep, setWorkflowStep] = useState(1)

  const fileType = file ? getFileType(file.name) : null

  function handleFile(selected) {
    setError('')
    // Clear previous scan results so the new scan starts fresh
    setScanData(null)
    if (!selected) return

    const type = getFileType(selected.name)
    if (!type) {
      setError('סוג קובץ לא נתמך. העלו קובץ .docx, .pptx או .xlsx בלבד.')
      setFile(null)
      setWorkflowStep(1)
      return
    }
    setFile(selected)
    setWorkflowStep(2)
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
      userType,
      scanType,
      results: null,
    })
    navigate('/scan')
  }

  return (
    <div className="page page-upload">
      <header className="page-header">
        <h1>התחלת סריקת נגישות</h1>
        <p className="page-lead">העלו קובץ Office, בחרו סוג משתמש וסריקה, והפעילו את הבדיקה.</p>
      </header>

      <ol className="workflow-indicator" aria-label="שלבי תהליך">
        {WORKFLOW_STEPS.map((label, i) => {
          const stepNum = i + 1
          const active = workflowStep === stepNum
          const done = workflowStep > stepNum
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

          {error && (
            <p className="form-error" role="alert">{error}</p>
          )}
        </section>

        <section className="card">
          <h2 className="card-title">2. הגדרת סריקה</h2>

          <div className="form-section">
            <h3 className="form-label">סוג משתמש</h3>
            <div className="option-cards" role="radiogroup" aria-label="סוג משתמש">
              {USER_TYPES.map((opt) => (
                <label
                  key={opt.value}
                  className={`option-card${userType === opt.value ? ' is-selected' : ''}`}
                >
                  <input
                    type="radio"
                    name="userType"
                    value={opt.value}
                    checked={userType === opt.value}
                    onChange={() => { setUserType(opt.value); if (file) setWorkflowStep(2) }}
                  />
                  <span className="option-card__title">{opt.label}</span>
                  <span className="option-card__desc">{opt.desc}</span>
                </label>
              ))}
            </div>
          </div>

          <div className="form-section">
            <h3 className="form-label">סוג סריקה</h3>
            <div className="option-cards" role="radiogroup" aria-label="סוג סריקה">
              {SCAN_TYPES.map((opt) => (
                <label
                  key={opt.value}
                  className={`option-card${scanType === opt.value ? ' is-selected' : ''}`}
                >
                  <input
                    type="radio"
                    name="scanType"
                    value={opt.value}
                    checked={scanType === opt.value}
                    onChange={() => { setScanType(opt.value); if (file) setWorkflowStep(3) }}
                  />
                  <span className="option-card__title">{opt.label}</span>
                  <span className="option-card__desc">{opt.desc}</span>
                </label>
              ))}
            </div>
          </div>
        </section>

        <section className="card card--actions">
          <h2 className="card-title">3. הרצת בדיקה</h2>
          <p className="card-hint">לאחר לחיצה, המערכת תעלה את הקובץ לשרת ותציג את תוצאות הסריקה.</p>
          <div className="btn-group">
            <Button type="submit" disabled={!file}>
              הפעלת סריקה
            </Button>
            <Button to="/" variant="secondary">חזרה לדף הבית</Button>
          </div>
        </section>
      </form>
    </div>
  )
}
