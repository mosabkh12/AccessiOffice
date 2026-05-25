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
const FILE_LABELS = { docx: 'Microsoft Word', pptx: 'Microsoft PowerPoint', xlsx: 'Microsoft Excel' }

const CRITICAL_DEFAULT =
  'סריקה אוטומטית מסייעת לזהות בעיות נגישות נפוצות, אך אינה מחליפה בדיקה אנושית מקצועית. יש לבדוק ידנית את איכות הטקסט החלופי, הקשר התוכן, סדר קריאה מורכב והתאמה מלאה לתקן הישראלי לפני פרסום המסמך.'

export default function ReportPage() {
  const navigate = useNavigate()
  const { scanData } = useScan()
  const results = scanData?.results

  useEffect(() => {
    if (!results) navigate('/upload')
  }, [results, navigate])

  if (!results) return null

  const {
    fileName,
    fileType,
    scannedAt,
    userType,
    scanType,
    score,
    summary,
    issues,
    quickFix,
    quickFixCount,
    quickFixOccurrences,
    recommendations,
    criticalAnalysis,
    totalOccurrences,
    totalIssueTypes,
    totalIssues,
    highOccurrences,
    mediumOccurrences,
    lowOccurrences,
  } = results

  function handlePrint() {
    requestAnimationFrame(() => window.print())
  }

  return (
    <div className="page page-report report-page">
      <div className="report-toolbar no-print">
        <Button onClick={handlePrint}>הדפסת דוח</Button>
        <Button to="/results" variant="secondary">חזרה לתוצאות</Button>
      </div>

      <article className="card report-document">
        <header className="report-header">
          <h1>דוח נגישות – AccessiOffice</h1>
          <p className="report-subtitle">בודק נגישות לקבצי Office</p>
        </header>

        <dl className="report-meta">
          <div>
            <dt>שם הקובץ</dt>
            <dd className="file-name-display">{fileName}</dd>
          </div>
          <div>
            <dt>תאריך סריקה</dt>
            <dd>{scannedAt}</dd>
          </div>
          <div>
            <dt>סוג הקובץ</dt>
            <dd>{FILE_LABELS[fileType]}</dd>
          </div>
          <div>
            <dt>סוג המשתמש</dt>
            <dd>{USER_LABELS[userType]}</dd>
          </div>
          <div>
            <dt>סוג הסריקה</dt>
            <dd>{SCAN_LABELS[scanType]}</dd>
          </div>
          <div>
            <dt>תקן</dt>
            <dd><span className="wcag-ltr" dir="ltr">WCAG 2.1</span> · ת&quot;י 5568</dd>
          </div>
        </dl>

        <section className="report-section report-score-block">
          <ScoreRing score={score} />
          <p className="report-score-text">
            <strong>ציון נגישות: {score}/100</strong>
          </p>
        </section>

        <section className="report-section">
          <h3>סיכום ממצאים</h3>
          <ul className="severity-summary-list">
            <li><span className="severity-tag high">גבוהה</span> {highOccurrences ?? 0} מופעים</li>
            <li><span className="severity-tag medium">בינונית</span> {mediumOccurrences ?? 0} מופעים</li>
            {(lowOccurrences ?? 0) > 0 && (
              <li><span className="severity-tag low">נמוכה</span> {lowOccurrences} מופעים</li>
            )}
            <li className="severity-total">
              <strong>סה&quot;כ ממצאי נגישות:</strong> {totalOccurrences ?? 0}
            </li>
            <li className="severity-total">
              <strong>סוגי בעיות:</strong> {totalIssueTypes ?? 0}
            </li>
            {(quickFixOccurrences ?? 0) > 0 && (
              <li className="severity-total">
                <strong>הצעות Quick Fix:</strong> {quickFixOccurrences}
              </li>
            )}
          </ul>
          <p className="report-summary-para">{summary}</p>
        </section>

        <section className="report-section">
          <h3>פירוט בעיות נגישות ({totalIssues})</h3>
          <IssuesTable issues={issues} variant="report" />
        </section>

        {quickFixCount > 0 && (
          <section className="report-section">
            <h3>הצעות Quick Fix – אובייקטים דקורטיביים ({quickFixCount})</h3>
            <p className="report-quickfix-note">
              האובייקטים הבאים אינם נושאים תוכן מהותי ומומלץ לסמן אותם כדקורטיביים ב-PowerPoint,
              כדי שקוראי מסך ידלגו עליהם ויחסכו מהמשתמש מידע מיותר.
            </p>
            <IssuesTable issues={quickFix} variant="report" />
          </section>
        )}

        <section className="report-section">
          <h3>המלצות לתיקון</h3>
          <ol className="report-list">
            {recommendations.map((rec, i) => (
              <li key={i}>{rec}</li>
            ))}
          </ol>
        </section>

        <section className="report-section">
          <h3>ניתוח ביקורתי – מגבלות הכלי</h3>
          <div className="limitations-box">
            <p>{criticalAnalysis || CRITICAL_DEFAULT}</p>
          </div>
        </section>

        <footer className="report-footer">
          AccessiOffice © 2026 · דוח נגישות אקדמי
        </footer>
      </article>
    </div>
  )
}
