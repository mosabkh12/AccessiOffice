import { useEffect } from 'react'
import { useNavigate } from 'react-router-dom'
import Button from '../components/Button.jsx'
import ScoreRing from '../components/ScoreRing.jsx'
import AssistantPanel from '../components/AssistantPanel.jsx'
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
const SCAN_LABELS  = { basic: 'סריקה בסיסית', full: 'סריקה מלאה' }
const FILE_LABELS  = { docx: 'Word', pptx: 'PowerPoint', xlsx: 'Excel' }
const DEBUG_SCAN = import.meta.env.VITE_ACCESSIOFFICE_DEBUG_SCAN === 'true'

const SEV_LABEL = { high: 'גבוהה', medium: 'בינונית', manual: 'בדיקה ידנית' }

// ── Office Engine grouped report table ─────────────────────────────────────────
function OfficeEngineReportTable({ officeLikeSummary }) {
  const allRows = buildOfficeEngineIssueRows(officeLikeSummary)
  if (allRows.length === 0) {
    return <p className="office-report-empty">כל הבדיקות האוטומטיות עברו ללא ממצאים.</p>
  }

  // Re-group preserving OFFICE_REPORT_GROUPS order (rows already come out in that order)
  const seen = new Set()
  const groupLabels = allRows.map(r => r.category).filter(c => !seen.has(c) && seen.add(c))

  return (
    <div className="office-report-table" dir="rtl">
      {groupLabels.map(groupLabel => (
        <section key={groupLabel} className="office-report-group">
          <h4 className="office-report-group__label">{groupLabel}</h4>
          {allRows.filter(r => r.category === groupLabel).map(row => (
            <div key={row.key} className={`office-report-row office-report-row--${row.severity}`}>
              <div className="office-report-row__main">
                <span className="office-report-row__label">{row.label}</span>
                {row.count > 0 && (
                  <span className={`office-report-row__count office-report-row__count--${row.severity}`}>
                    {row.count}
                  </span>
                )}
                <span className={`severity-tag ${row.severity}`}>{SEV_LABEL[row.severity]}</span>
              </div>
              {row.recommendation && <p className="office-report-row__rec">{row.recommendation}</p>}
              {row.note && <p className="office-report-row__note">{row.note}</p>}
            </div>
          ))}
        </section>
      ))}
    </div>
  )
}

// ── Main page ──────────────────────────────────────────────────────────────────
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
    officeLikeSummary,
    fileHash,
    engine,
    workerError,
    fileName,
    fileType,
    userType,
    scanType,
  } = results

  const isOffice     = isOfficeEngineResult(results)
  const officeSummary = isOffice ? buildOfficeEngineSummary(officeLikeSummary) : null
  const officeScore   = isOffice ? calculateOfficeEngineScore(officeLikeSummary) : null

  // Top-card display values — prefer office engine when available
  const displayScore      = officeScore              ?? score             ?? 0
  const displayTotal      = officeSummary?.totalFindings ?? totalOccurrences  ?? 0
  const displayIssueTypes = officeSummary?.issueTypes    ?? totalIssueTypes   ?? 0
  const displayHigh       = officeSummary?.highCount     ?? highOccurrences   ?? 0
  const displayMedium     = officeSummary?.mediumCount   ?? mediumOccurrences ?? 0

  if (DEBUG_SCAN) {
    console.log('[RESULTS PAGE]', {
      engine, isOffice, displayScore, displayTotal, displayHigh, displayMedium,
      officeSummary,
      officeLikeSummary: officeLikeSummary
        ? Object.entries(officeLikeSummary).map(([k, v]) => `${k}: ${v?.status} (${v?.count})`)
        : null,
      fileHash,
    })
  }

  return (
    <div className="page page-results">
      {/* ── Fallback banner: worker failed, XML results shown ─────────────────── */}
      {engine === 'hybrid' && workerError && (
        <div className="worker-fallback-banner" role="alert" dir="rtl">
          <span className="worker-fallback-banner__icon" aria-hidden="true">⚠</span>
          <span className="worker-fallback-banner__text">
            הנתונים מבוססים על ניתוח XML/WCAG.
          </span>
        </div>
      )}

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
          <ScoreRing score={displayScore} />
          <p className="stat-label">ציון נגישות</p>
          <p className="stat-sublabel">מתוך 100</p>
          {!isOffice && (
            <p className="stat-score-source">XML/WCAG</p>
          )}
        </article>

        <article className="card stat-card">
          <p className="stat-value">{displayTotal}</p>
          <p className="stat-label">סה&quot;כ ממצאים</p>
          <p className="stat-sublabel">{displayIssueTypes} סוגי בעיות</p>
        </article>

        <article className="card stat-card stat-card--high">
          <p className="stat-value">{displayHigh}</p>
          <p className="stat-label">חומרה גבוהה</p>
        </article>

        <article className="card stat-card stat-card--medium">
          <p className="stat-value">{displayMedium}</p>
          <p className="stat-label">חומרה בינונית</p>
        </article>

        {!isOffice && (quickFixOccurrences ?? 0) > 0 && (
          <article className="card stat-card stat-card--quickfix">
            <p className="stat-value">{quickFixOccurrences}</p>
            <p className="stat-label">Quick Fix</p>
          </article>
        )}
      </section>

      {/* ── Card A: Accessibility Assistant checklist ─────────────────────────── */}
      <section className="card card--assistant">
        <AssistantPanel
          issues={isOffice ? [] : issues}
          quickFix={isOffice ? [] : quickFix}
          fileType={fileType}
          scanDiagnostics={scanDiagnostics}
          checkStatuses={checkStatuses}
          scannerVersion={scannerVersion}
          officeLikeSummary={isOffice ? officeLikeSummary : null}
          engine={engine}
          workerError={workerError}
        />
      </section>

      {/* ── Card B: Office Engine grouped report table (office-engine only) ──── */}
      {isOffice && (
        <section className="card card--office-report">
          <OfficeEngineReportTable officeLikeSummary={officeLikeSummary} />
        </section>
      )}

      {/* ── Debug-only: XML/WCAG scanner output — hidden in normal mode and print ── */}
      {isOffice && DEBUG_SCAN && (issues.length > 0 || quickFix.length > 0) && (
        <details className="card card--xml-secondary debug-only">
          <summary className="xml-secondary-summary">
            <span className="xml-secondary-title">Debug: XML/WCAG scanner output</span>
            <span className="xml-secondary-badge debug-badge">debug only</span>
          </summary>
          <p className="xml-secondary-note" dir="ltr">
            Developer view — visible only when VITE_ACCESSIOFFICE_DEBUG_SCAN=true. Not shown in production or print.
          </p>
          <AssistantPanel
            issues={issues}
            quickFix={quickFix}
            fileType={fileType}
            scanDiagnostics={scanDiagnostics}
            checkStatuses={checkStatuses}
            scannerVersion={scannerVersion}
            detailedOnly
          />
        </details>
      )}

      <div className="btn-group no-print">
        <Button to="/report">צפייה בדוח מלא</Button>
        <Button to="/upload" variant="secondary">סריקה נוספת</Button>
      </div>
    </div>
  )
}
