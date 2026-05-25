import { useEffect } from 'react'
import { useNavigate } from 'react-router-dom'
import Button from '../components/Button.jsx'
import ScoreRing from '../components/ScoreRing.jsx'
import AssistantPanel from '../components/AssistantPanel.jsx'
import { useScan } from '../context/ScanContext.jsx'

const USER_LABELS = {
  author: 'מחבר/ת מסמך',
  auditor: 'בודק/ת נגישות',
  lecturer: 'מרצה / מוסד לימודי',
}
const SCAN_LABELS  = { basic: 'סריקה בסיסית', full: 'סריקה מלאה' }
const FILE_LABELS  = { docx: 'Word', pptx: 'PowerPoint', xlsx: 'Excel' }
const DEBUG_SCAN = import.meta.env.VITE_ACCESSIOFFICE_DEBUG_SCAN === 'true'

export default function ResultsPage() {
  const navigate = useNavigate()
  const { scanData } = useScan()
  const results = scanData?.results

  useEffect(() => {
    if (!results) navigate('/upload')
  }, [results, navigate])

  if (!results) return null

  const {
    score,
    totalOccurrences,
    totalIssueTypes,
    highOccurrences,
    mediumOccurrences,
    quickFixOccurrences,
    issues,
    quickFix,
    scanDiagnostics,
    checkStatuses,
    scannerVersion,
    fileName,
    fileType,
    userType,
    scanType,
  } = results

  // All counts flow from the same normalized issues array — never computed separately
  if (DEBUG_SCAN) {
    console.log('[RESULTS PAGE RENDER]', {
      totalOccurrences,
      totalIssueTypes,
      highOccurrences,
      mediumOccurrences,
      quickFixOccurrences,
      issueRows: issues.map(i => `${i.title}: ${i.occurrenceCount}`),
      quickFixRows: quickFix.map(i => `${i.title}: ${i.occurrenceCount}`),
      scannerVersion,
      scanDiagnostics,
      checkStatuses,
    })
  }

  return (
    <div className="page page-results">
      <header className="page-header">
        <h1>תוצאות סריקת נגישות</h1>
        <dl className="meta-chips">
          <div><dt>קובץ</dt><dd className="file-name-display">{fileName}</dd></div>
          <div><dt>סוג</dt><dd>{FILE_LABELS[fileType]}</dd></div>
          <div><dt>משתמש</dt><dd>{USER_LABELS[userType]}</dd></div>
          <div><dt>סריקה</dt><dd>{SCAN_LABELS[scanType]}</dd></div>
        </dl>
      </header>

      {/* ── Dashboard summary cards ──────────────────────────────────────────── */}
      <section className="dashboard-grid">
        <article className="card stat-card stat-card--score">
          <ScoreRing score={score} />
          <p className="stat-label">ציון נגישות</p>
          <p className="stat-sublabel">מתוך 100</p>
        </article>

        <article className="card stat-card">
          <p className="stat-value">{totalOccurrences ?? 0}</p>
          <p className="stat-label">סה&quot;כ ממצאים</p>
          <p className="stat-sublabel">{totalIssueTypes ?? 0} סוגי בעיות</p>
        </article>

        <article className="card stat-card stat-card--high">
          <p className="stat-value">{highOccurrences ?? 0}</p>
          <p className="stat-label">חומרה גבוהה</p>
        </article>

        <article className="card stat-card stat-card--medium">
          <p className="stat-value">{mediumOccurrences ?? 0}</p>
          <p className="stat-label">חומרה בינונית</p>
        </article>

        {(quickFixOccurrences ?? 0) > 0 && (
          <article className="card stat-card stat-card--quickfix">
            <p className="stat-value">{quickFixOccurrences}</p>
            <p className="stat-label">Quick Fix</p>
          </article>
        )}
      </section>

      {/* ── Office-like Accessibility Assistant Panel ────────────────────────── */}
      <section className="card card--assistant">
        <AssistantPanel
          issues={issues}
          quickFix={quickFix}
          fileType={fileType}
          scanDiagnostics={scanDiagnostics}
          checkStatuses={checkStatuses}
          scannerVersion={scannerVersion}
        />
      </section>

      <div className="btn-group no-print">
        <Button to="/report">צפייה בדוח מלא</Button>
        <Button to="/upload" variant="secondary">סריקה נוספת</Button>
      </div>
    </div>
  )
}
