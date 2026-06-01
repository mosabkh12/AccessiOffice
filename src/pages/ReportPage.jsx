import { useEffect } from 'react'
import { useNavigate } from 'react-router-dom'
import Button from '../components/Button.jsx'
import ScoreRing from '../components/ScoreRing.jsx'
import IssuesTable from '../components/IssuesTable.jsx'
import { useScan } from '../context/ScanContext.jsx'
import {
  isOfficeEngineResult,
  buildOfficeEngineSummary,
  calculateOfficeEngineScore,
  buildOfficeEngineIssueRows,
} from '../utils/officeLikeReport.js'

const USER_LABELS = {
  author: 'מחבר/ת מסמך',
  auditor: 'בודק/ת נגישות',
  lecturer: 'מרצה / מוסד לימודי',
}
const SCAN_LABELS = { basic: 'סריקה בסיסית', full: 'סריקה מלאה' }
const FILE_LABELS = { docx: 'Microsoft Word', pptx: 'Microsoft PowerPoint', xlsx: 'Microsoft Excel' }

const CRITICAL_DEFAULT =
  'סריקה אוטומטית מסייעת לזהות בעיות נגישות נפוצות, אך אינה מחליפה בדיקה אנושית מקצועית. יש לבדוק ידנית את איכות הטקסט החלופי, הקשר התוכן, סדר קריאה מורכב והתאמה מלאה לתקן הישראלי לפני פרסום המסמך.'

const SEV_LABEL = { high: 'גבוהה', medium: 'בינונית', manual: 'בדיקה ידנית', low: 'נמוכה' }
const DEBUG_SCAN = import.meta.env.VITE_ACCESSIOFFICE_DEBUG_SCAN === 'true'

// ── Office Engine row card (print-friendly) ────────────────────────────────────
// row carries: label, count, severity, category, recommendation, impact, wcag,
//              fixSteps[], note (generic "נבדק על ידי" notes already stripped).
function OfficeReportRow({ row }) {
  return (
    <article className="report-issue-card">
      <header className="report-issue-card__header">
        <h4>
          {row.label}
          {row.count > 0 && <span className="occurrence-badge">{row.count}</span>}
        </h4>
        <span className={`severity-tag ${row.severity}`}>{SEV_LABEL[row.severity]}</span>
      </header>
      <dl className="report-issue-card__body">
        <div>
          <dt>קטגוריה</dt>
          <dd>{row.category}</dd>
        </div>
        {row.impact && (
          <div>
            <dt>השפעה</dt>
            <dd>{row.impact}</dd>
          </div>
        )}
        {row.recommendation && (
          <div>
            <dt>המלצה</dt>
            <dd>{row.recommendation}</dd>
          </div>
        )}
        {row.wcag && row.wcag !== '—' && (
          <div>
            <dt><span dir="ltr">WCAG</span> · ת&quot;י&nbsp;5568</dt>
            <dd dir="ltr" className="wcag-ltr">{row.wcag}</dd>
          </div>
        )}
        {row.fixSteps?.length > 0 && (
          <div className="fix-steps">
            <dt className="fix-steps__title">שלבי תיקון</dt>
            <dd>
              <ol className="fix-steps__list">
                {row.fixSteps.map((step, i) => (
                  <li key={i} className="fix-steps__item">{step}</li>
                ))}
              </ol>
            </dd>
          </div>
        )}
        {row.note && (
          <div>
            <dt>הערה</dt>
            <dd>{row.note}</dd>
          </div>
        )}
      </dl>
    </article>
  )
}

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
    manualReviewChecks,
    totalOccurrences,
    totalIssueTypes,
    totalIssues,
    highOccurrences,
    mediumOccurrences,
    lowOccurrences,
    officeLikeSummary,
    engine,
  } = results

  const isOffice = isOfficeEngineResult(results)

  // Human-readable engine name used across several labels
  const engineName = isOffice
    ? (fileType === 'docx'
        ? 'Microsoft Word Accessibility Checker'
        : fileType === 'xlsx'
          ? 'Microsoft Excel Accessibility Checker'
          : 'Microsoft PowerPoint Accessibility Checker')
    : null

  // Office Engine derived values — computed only when relevant
  const officeSummary = isOffice ? buildOfficeEngineSummary(officeLikeSummary) : null
  const officeScore   = isOffice ? calculateOfficeEngineScore(officeLikeSummary) : null
  const officeRows    = isOffice ? buildOfficeEngineIssueRows(officeLikeSummary) : []
  const officeRecs    = isOffice
    ? [...new Set(officeRows.map(r => r.recommendation).filter(Boolean))]
    : recommendations

  const displayScore = officeScore ?? score

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
          <p className="report-subtitle">
            {isOffice ? `מבוסס על ${engineName}` : 'בודק נגישות לקבצי Office'}
          </p>
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
          {isOffice && (
            <div>
              <dt>מנוע בדיקה</dt>
              <dd>{engineName}</dd>
            </div>
          )}
        </dl>

        {/* ── Score ── */}
        <section className="report-section report-score-block">
          <ScoreRing score={displayScore} />
          <p className="report-score-text">
            <strong>ציון נגישות: {displayScore}/100</strong>
          </p>
          {isOffice && (
            <p className="report-score-source">
              הציון מחושב על בסיס ממצאי {engineName}
            </p>
          )}
        </section>

        {/* ══════════════════════════════════════════════════════════════════════
            OFFICE ENGINE PRIMARY REPORT
        ══════════════════════════════════════════════════════════════════════ */}
        {isOffice && (
          <>
            {/* Summary */}
            <section className="report-section">
              <h3>סיכום ממצאים — {engineName}</h3>
              <ul className="severity-summary-list">
                <li>
                  <span className="severity-tag high">גבוהה</span>
                  {' '}{officeSummary.highCount} מופעים
                </li>
                <li>
                  <span className="severity-tag medium">בינונית</span>
                  {' '}{officeSummary.mediumCount} מופעים
                </li>
                {officeSummary.manualCount > 0 && (
                  <li>
                    <span className="severity-tag manual">בדיקה ידנית</span>
                    {' '}{officeSummary.manualCount} בדיקות
                  </li>
                )}
                <li className="severity-total">
                  <strong>סה&quot;כ ממצאים:</strong> {officeSummary.totalFindings}
                </li>
                <li className="severity-total">
                  <strong>סוגי בעיות:</strong> {officeSummary.issueTypes}
                </li>
              </ul>
            </section>

            {/* Issue rows — one grouped row per check, no XML repeats */}
            <section className="report-section">
              <h3>פירוט ממצאים ({officeRows.length})</h3>
              {officeRows.length === 0
                ? <p>כל הבדיקות האוטומטיות עברו ללא ממצאים.</p>
                : (
                  <div className="report-issues-grid">
                    {officeRows.map(row => <OfficeReportRow key={row.key} row={row} />)}
                  </div>
                )
              }
            </section>

            {/* Recommendations */}
            <section className="report-section">
              <h3>המלצות לתיקון</h3>
              <ol className="report-list">
                {officeRecs.map((rec, i) => <li key={i}>{rec}</li>)}
              </ol>
            </section>

            {/* Limitations */}
            <section className="report-section">
              <h3>ניתוח ביקורתי – מגבלות הכלי</h3>
              <div className="limitations-box">
                <p>{criticalAnalysis || CRITICAL_DEFAULT}</p>
              </div>
            </section>

            {/* Secondary XML section — debug only, hidden from print */}
            {DEBUG_SCAN && (issues.length > 0 || quickFix.length > 0) && (
              <details className="report-section xml-secondary-details debug-only">
                <summary className="xml-secondary-summary xml-secondary-summary--report">
                  <span className="xml-secondary-title">Debug: XML/WCAG scanner output</span>
                  <span className="xml-secondary-badge debug-badge">debug only</span>
                </summary>
                <p className="xml-secondary-note" dir="ltr">
                  Developer view — visible only when VITE_ACCESSIOFFICE_DEBUG_SCAN=true. Not shown in production or print.
                </p>

                {issues.length > 0 && (
                  <div className="report-section">
                    <h4 className="report-xml-heading">
                      בעיות WCAG — ניתוח XML ({issues.length} סוגים)
                    </h4>
                    <IssuesTable issues={issues} variant="report" />
                  </div>
                )}

                {quickFix.length > 0 && (
                  <div className="report-section">
                    <h4 className="report-xml-heading">
                      הצעות Quick Fix ({quickFixCount})
                    </h4>
                    <p className="report-quickfix-note">
                      אובייקטים שניתן לסמן כדקורטיביים ב-PowerPoint.
                    </p>
                    <IssuesTable issues={quickFix} variant="report" />
                  </div>
                )}
              </details>
            )}
          </>
        )}

        {/* ══════════════════════════════════════════════════════════════════════
            XML / STANDARD REPORT (non-office-engine)
        ══════════════════════════════════════════════════════════════════════ */}
        {!isOffice && (
          <>
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

            {(manualReviewChecks?.length ?? 0) > 0 && (
              <section className="report-section">
                <h3>בדיקות ידניות / חלקיות</h3>
                <div className="limitations-box">
                  <ul className="report-list">
                    {manualReviewChecks.map((check) => (
                      <li key={check.id}>
                        <strong>{check.label}</strong>
                        {` — ${check.status === 'partial' ? 'בדיקה חלקית' : check.status === 'not_checked' ? 'לא נבדק אוטומטית' : 'נדרשת סקירה ידנית'}`}
                        {check.note ? `: ${check.note}` : ''}
                      </li>
                    ))}
                  </ul>
                </div>
              </section>
            )}

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
          </>
        )}

        <footer className="report-footer">
          AccessiOffice © 2026 · דוח נגישות אקדמי
        </footer>
      </article>
    </div>
  )
}
