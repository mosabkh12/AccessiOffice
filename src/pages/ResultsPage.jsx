import { useEffect } from 'react'
import { useNavigate } from 'react-router-dom'
import Button from '../components/Button.jsx'
import ScoreRing from '../components/ScoreRing.jsx'
import IssuesTable from '../components/IssuesTable.jsx'
import { useScan } from '../context/ScanContext.jsx'

const USER_LABELS = {
  author: 'מחבר/ת מסמך',
  auditor: 'בודק/ת נגישות',
  lecturer: 'מרצה / מוסד לימודי',
}

const SCAN_LABELS = { basic: 'סריקה בסיסית', full: 'סריקה מלאה' }
const FILE_LABELS = { docx: 'Word', pptx: 'PowerPoint', xlsx: 'Excel' }

const SUMMARY_DEFAULT =
  'על פי ממצאי הסריקה, הקובץ כולל בעיות נגישות שיש לתקן לפני פרסום או שיתוף.'

const NO_ISSUES_MSG =
  'לא נמצאו בעיות נגישות משמעותיות בסריקה האוטומטית. מומלץ לבצע בדיקה ידנית לפני פרסום.'

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
    totalIssues,
    severityCounts,
    issues,
    fileName,
    fileType,
    userType,
    scanType,
    summary,
  } = results

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

      <section className="dashboard-grid">
        <article className="card stat-card stat-card--score">
          <ScoreRing score={score} />
          <p className="stat-label">ציון נגישות</p>
          <p className="stat-sublabel">מתוך 100</p>
        </article>
        <article className="card stat-card">
          <p className="stat-value">{totalIssues}</p>
          <p className="stat-label">סה&quot;כ בעיות</p>
        </article>
        <article className="card stat-card stat-card--high">
          <p className="stat-value">{severityCounts.high}</p>
          <p className="stat-label">חומרה גבוהה</p>
        </article>
        <article className="card stat-card stat-card--medium">
          <p className="stat-value">{severityCounts.medium}</p>
          <p className="stat-label">חומרה בינונית</p>
        </article>
        <article className="card stat-card stat-card--low">
          <p className="stat-value">{severityCounts.low}</p>
          <p className="stat-label">חומרה נמוכה</p>
        </article>
      </section>

      <article className="card summary-banner">
        <h2 className="summary-banner__title">סיכום</h2>
        <p>{totalIssues === 0 ? NO_ISSUES_MSG : summary || SUMMARY_DEFAULT}</p>
      </article>

      <section className="card">
        <h2 className="card-title">פירוט בעיות נגישות ({totalIssues})</h2>
        {totalIssues === 0 ? (
          <p className="card-hint">{NO_ISSUES_MSG}</p>
        ) : (
          <>
            <p className="card-hint">לחצו על &quot;פרטים&quot; לצפייה בהשפעה, המלצה והסבר WCAG.</p>
            <IssuesTable issues={issues} />
          </>
        )}
      </section>

      <div className="btn-group no-print">
        <Button to="/report">צפייה בדוח מלא</Button>
        <Button to="/upload" variant="secondary">סריקה נוספת</Button>
      </div>
    </div>
  )
}
