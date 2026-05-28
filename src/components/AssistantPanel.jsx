import { useState } from 'react'

// ── Known checks per file type. PPTX uses diagnostics for three-state status.

const CHECKS_BY_TYPE = {
  pptx: [
    { id: 'missing-alt', label: 'חסר טקסט חלופי', group: 'מדיה ואיורים' },
    { id: 'media-captions', label: 'חסרות כתוביות לאודיו/וידאו', group: 'מדיה ואיורים' },
    { id: 'missing-title', label: 'שקופית ללא כותרת', group: 'מבנה מסמך' },
    { id: 'duplicate-title', label: 'כותרת שקופית כפולה', group: 'מבנה מסמך' },
    { id: 'reading-order', label: 'בדיקת סדר קריאה', group: 'מבנה מסמך' },
    { id: 'contrast', label: 'ניגודיות טקסט קשה לקריאה', group: 'צבע וניגודיות' },
    { id: 'table-header', label: 'חסרת שורת כותרת בטבלה', group: 'טבלאות' },
    { id: 'merged-cells', label: 'שימוש בתאים ממוזגים או מפוצלים', group: 'טבלאות' },
    { id: 'restricted-access', label: 'גישה מוגבלת למסמך', group: 'גישה למסמך', manual: true },
    { id: 'unclear-link', label: 'קישורים לא ברורים', group: 'ניווט' },
  ],
  docx: [
    { label: 'תמונה ללא טקסט חלופי',   prefixes: ['missing-alt'],      group: 'תוכן' },
    { label: 'מבנה כותרות',             prefixes: ['incorrect-heading', 'heading-skip'], group: 'מבנה' },
    { label: 'קישורים לא ברורים',      prefixes: ['unclear-link'],     group: 'ניווט' },
    { label: 'טבלאות',                  prefixes: ['table-header'],     group: 'טבלאות' },
    { label: 'ניגודיות צבעים',          prefixes: ['low-contrast', 'contrast'], group: 'תצוגה' },
  ],
  xlsx: [
    { label: 'שם גיליון לא תיאורי',    prefixes: ['sheet-name'],       group: 'מבנה' },
    { label: 'תאים ממוזגים',            prefixes: ['merged'],           group: 'טבלאות' },
    { label: 'כותרת חסרה',             prefixes: ['header', 'columns-no-header'], group: 'טבלאות' },
    { label: 'תאים ריקים',             prefixes: ['empty-cells'],      group: 'מבנה' },
    { label: 'ניגודיות צבעים',          prefixes: ['contrast'],         group: 'תצוגה' },
  ],
}

function matchesCheck(issue, check) {
  if (check.prefixes) return check.prefixes.some((p) => issue.id?.startsWith(p))

  const id = issue.id || ''
  const title = issue.title || ''
  const wcag = issue.wcagKey || issue.wcagCriterion || ''
  const category = issue.category || ''

  switch (check.id) {
    case 'contrast':
      return id.includes('contrast') || wcag.includes('Contrast') || wcag === 'contrast-minimum' || /ניגודיות|contrast/i.test(title) || /ניגודיות|contrast/i.test(category)
    case 'missing-alt':
      return id.includes('missing-alt') || title.includes('טקסט חלופי')
    case 'missing-title':
      return id.includes('slide-without-title') || id.includes('missing-title') || title.includes('שקופית ללא כותרת')
    case 'duplicate-title':
      return id.includes('duplicate-title') || title.includes('כותרת שקופית כפולה')
    case 'reading-order':
      return id.includes('reading-order') || title.includes('סדר קריאה')
    case 'table-header':
      return id.includes('table-header') || title.includes('שורת כותרת') || title.includes('כותרת בטבלה')
    case 'merged-cells':
      return id.includes('merged-cells') || title.includes('ממוזגים') || title.includes('מפוצלים')
    case 'media-captions':
      return id.includes('media-captions') || title.includes('כתוביות') || wcag.includes('Captions')
    case 'unclear-link':
      return id.includes('unclear-link') || title.includes('קישור לא ברור') || title.includes('קישורים לא ברורים')
    default:
      return false
  }
}

function issueCount(issues) {
  return issues.reduce((sum, issue) => sum + (issue.occurrenceCount ?? 1), 0)
}

// ── Sub-components ─────────────────────────────────────────────────────────────

function SeverityMarker({ severity, isQuickFix }) {
  if (isQuickFix) return <span className="ap-marker ap-marker--qf" aria-hidden="true" />
  if (severity === 'high')   return <span className="ap-marker ap-marker--high" aria-hidden="true" />
  if (severity === 'medium') return <span className="ap-marker ap-marker--medium" aria-hidden="true" />
  return                            <span className="ap-marker ap-marker--low" aria-hidden="true" />
}

function IssueRow({ issue, isExpanded, onToggle, isQuickFix = false }) {
  const count = issue.occurrenceCount ?? 1
  const countClass = isQuickFix ? 'qf' : issue.severity

  return (
    <div className={`ap-issue${isExpanded ? ' ap-issue--open' : ''}`}>
      <button
        className="ap-issue__row"
        onClick={onToggle}
        aria-expanded={isExpanded}
        type="button"
      >
        <SeverityMarker severity={issue.severity} isQuickFix={isQuickFix} />
        <span className="ap-issue__title">{issue.title}</span>
        <span className={`ap-issue__count ap-issue__count--${countClass}`}>{count}</span>
        <span className="ap-issue__chevron" aria-hidden="true">{isExpanded ? '▲' : '▼'}</span>
      </button>

      {isExpanded && (
        <div className="ap-issue__details" role="region">
          {/* Location list (per-occurrence) */}
          {issue.locations?.length > 0 ? (
            <ul className="ap-locations" aria-label={`מיקומים עבור: ${issue.title}`}>
              {issue.locations.map((loc, i) => (
                <li key={i} className="ap-location">{loc}</li>
              ))}
            </ul>
          ) : (
            <p className="ap-location ap-location--single">{issue.location}</p>
          )}

          {/* WCAG + impact + recommendation */}
          <dl className="ap-issue__meta">
            {issue.impact && (
              <div>
                <dt>השפעה</dt>
                <dd>{issue.impact}</dd>
              </div>
            )}
            {issue.recommendation && (
              <div>
                <dt>המלצה</dt>
                <dd>{issue.recommendation}</dd>
              </div>
            )}
            {issue.wcagCriterion && issue.wcagCriterion !== '—' && (
              <div>
                <dt>WCAG</dt>
                <dd dir="ltr" className="wcag-ltr">{issue.wcagCriterion} · Level {issue.wcagLevel}</dd>
              </div>
            )}
          </dl>
        </div>
      )}
    </div>
  )
}

function PassedRow({ label }) {
  return (
    <div className="ap-check-pass">
      <span className="ap-marker ap-marker--pass" aria-hidden="true" />
      <span className="ap-check-pass__label">{label}</span>
    </div>
  )
}

function ManualReviewRow({ label, note }) {
  return (
    <div className="ap-check-manual">
      <span className="ap-marker ap-marker--manual" aria-hidden="true" />
      <span className="ap-check-pass__label">{label}</span>
      {note && <span className="ap-check-manual__note">{note}</span>}
    </div>
  )
}

// ── Office-like check row (used when officeLikeSummary is present) ──────────────

const OFFICE_GROUPS = [
  { label: 'צבע וניגודיות', keys: ['contrast'] },
  { label: 'מדיה ואיורים', keys: ['missingAltText', 'mediaCaptions'] },
  { label: 'טבלאות', keys: ['tableHeader', 'mergedCells'] },
  { label: 'מבנה מסמך', keys: ['missingSlideTitle', 'duplicateSlideTitle', 'readingOrder'] },
  { label: 'גישה למסמך', keys: ['restrictedAccess'] },
]

function OfficeLikeCheckRow({ item }) {
  if (item.status === 'passed') {
    return (
      <div className="ap-check-pass">
        <span className="ap-marker ap-marker--pass" aria-hidden="true" />
        <span className="ap-check-pass__label">{item.label}</span>
      </div>
    )
  }
  if (item.status === 'failed') {
    return (
      <div className="ap-check-pass">
        <span className="ap-marker ap-marker--medium" aria-hidden="true" />
        <span className="ap-check-pass__label">{item.label}</span>
        <span className="ap-issue__count ap-issue__count--medium">{item.count}</span>
      </div>
    )
  }
  // manual / partial / not_checked — show count if > 0, no red coloring
  return (
    <div className="ap-check-manual">
      <span className="ap-marker ap-marker--manual" aria-hidden="true" />
      <span className="ap-check-pass__label">{item.label}</span>
      {item.count > 0 && (
        <span className="ap-issue__count ap-issue__count--medium">{item.count}</span>
      )}
    </div>
  )
}

// ── Main component ─────────────────────────────────────────────────────────────

export default function AssistantPanel({
  issues = [],
  quickFix = [],
  fileType = 'pptx',
  scanDiagnostics = {},
  checkStatuses = [],
  scannerVersion,
  officeLikeSummary = null,
  detailedOnly = false,
  engine = 'xml',
  workerError = null,
}) {
  const [expanded, setExpanded] = useState(new Set())

  function toggle(id) {
    setExpanded((prev) => {
      const next = new Set(prev)
      next.has(id) ? next.delete(id) : next.add(id)
      return next
    })
  }

  const qfOccurrences = quickFix.reduce((s, i) => s + (i.occurrenceCount ?? 1), 0)

  // ── Detailed-only rendering (Card B — raw WCAG issues + quickFix, no check-status logic) ───
  if (detailedOnly) {
    const totalOcc = issues.reduce((s, i) => s + (i.occurrenceCount ?? 1), 0)
    const highOcc  = issues.filter((i) => i.severity === 'high').reduce((s, i) => s + (i.occurrenceCount ?? 1), 0)
    const medOcc   = issues.filter((i) => i.severity === 'medium').reduce((s, i) => s + (i.occurrenceCount ?? 1), 0)

    if (issues.length === 0 && quickFix.length === 0) return null

    return (
      <div className="assistant-panel" dir="rtl">
        <div className="ap-header">
          <span className="ap-header__title">ממצאים מפורטים</span>
          <span className="ap-header__subtitle">
            {issues.length} סוגי בעיות · {totalOcc} מופעים
          </span>
        </div>

        {issues.length > 0 && (
          <section className="ap-section">
            <div className="ap-section__label">
              בעיות נגישות
              <span className="ap-section__tally">
                {highOcc > 0 && <span className="ap-tally ap-tally--high">{highOcc} גבוהה</span>}
                {medOcc  > 0 && <span className="ap-tally ap-tally--medium">{medOcc} בינונית</span>}
              </span>
            </div>
            {issues.map((issue, index) => (
              <IssueRow
                key={`${issue.id}-${issue.location}-${index}`}
                issue={issue}
                isExpanded={expanded.has(issue.id)}
                onToggle={() => toggle(issue.id)}
              />
            ))}
          </section>
        )}

        {quickFix.length > 0 && (
          <section className="ap-section ap-section--qf">
            <div className="ap-section__label">
              הצעות Quick Fix
              <span className="ap-tally ap-tally--qf">{qfOccurrences}</span>
            </div>
            {quickFix.map((issue, index) => (
              <IssueRow
                key={`${issue.id}-${issue.location}-${index}`}
                issue={issue}
                isExpanded={expanded.has(issue.id)}
                onToggle={() => toggle(issue.id)}
                isQuickFix
              />
            ))}
          </section>
        )}
      </div>
    )
  }

  // ── Office-like rendering when officeLikeSummary is present for PPTX ─────────
  if (fileType === 'pptx' && officeLikeSummary) {
    const summaryItems = Object.values(officeLikeSummary)
    const totalFailCount = summaryItems
      .filter((it) => it.status === 'failed')
      .reduce((s, it) => s + it.count, 0)

    return (
      <div className="assistant-panel" dir="rtl">
        {/* Worker error banner — shown when PPT worker failed and XML fallback was used */}
        {workerError && (
          <div className="ap-worker-warning" role="alert" dir="rtl">
            <span className="ap-worker-warning__icon" aria-hidden="true">⚠</span>
            <span className="ap-worker-warning__text">
              מנוע PowerPoint לא זמין — הנתונים מבוססים על ניתוח XML.
            </span>
          </div>
        )}
        {/* Engine badge — shown when the real PPT worker was used */}
        {engine === 'office-engine' && !workerError && (
          <div className="ap-engine-badge" dir="rtl">
            <span className="ap-engine-badge__icon" aria-hidden="true">✓</span>
            מופעל על ידי Microsoft PowerPoint Accessibility Checker
          </div>
        )}
        <div className="ap-header">
          <span className="ap-header__title">בדיקת נגישות</span>
          <span className="ap-header__subtitle">
            {totalFailCount > 0
              ? `נמצאו ${totalFailCount} ממצאים לתיקון`
              : 'כל הבדיקות האוטומטיות עברו'}
          </span>
        </div>

        {OFFICE_GROUPS.map((group) => (
          <section key={group.label} className="ap-section">
            <div className="ap-section__label">{group.label}</div>
            {group.keys.map((key) => {
              const it = officeLikeSummary[key]
              if (!it) return null
              return <OfficeLikeCheckRow key={key} item={it} />
            })}
          </section>
        ))}

        {quickFix.length > 0 && (
          <section className="ap-section ap-section--qf">
            <div className="ap-section__label">
              הצעות Quick Fix
              <span className="ap-tally ap-tally--qf">{qfOccurrences}</span>
            </div>
            {quickFix.map((issue, index) => (
              <IssueRow
                key={`${issue.id}-${issue.location}-${index}`}
                issue={issue}
                isExpanded={expanded.has(issue.id)}
                onToggle={() => toggle(issue.id)}
                isQuickFix
              />
            ))}
          </section>
        )}
      </div>
    )
  }

  // ── Standard rendering (non-PPTX or no officeLikeSummary) ────────────────────

  // Counts derived directly from the issues array — single source of truth
  const totalOccurrences   = issues.reduce((s, i) => s + (i.occurrenceCount ?? 1), 0)
  const highOccurrences    = issues.filter((i) => i.severity === 'high').reduce((s, i) => s + (i.occurrenceCount ?? 1), 0)
  const mediumOccurrences  = issues.filter((i) => i.severity === 'medium').reduce((s, i) => s + (i.occurrenceCount ?? 1), 0)

  const knownChecks = CHECKS_BY_TYPE[fileType] ?? []
  const backendStatusById = new Map((checkStatuses || []).map((check) => [check.id, check]))
  const checkResults = knownChecks.map((check) => {
    const matches = issues.filter((issue) => matchesCheck(issue, check))
    const backendStatus = backendStatusById.get(check.id)
    let status = matches.length > 0 ? 'issue' : backendStatus?.status ?? (check.manual ? 'manual' : 'passed')
    let note = backendStatus?.note || ''
    const backendCount = typeof backendStatus?.count === 'number' ? backendStatus.count : undefined

    if (fileType === 'pptx' && check.id === 'contrast' && matches.length === 0) {
      const contrastStatus = scanDiagnostics?.contrastCheckStatus
      if (!backendStatus && contrastStatus !== 'checked') {
        status = 'manual'
        note =
          contrastStatus === 'partial'
            ? 'נסרקו רק צבעי sRGB מפורשים; צבעי ערכת נושא או ירושה דורשים בדיקה ידנית.'
            : 'לא נמצאו מספיק צבעי טקסט מפורשים לסריקת ניגודיות אוטומטית.'
      }
    }

    if (!backendStatus && fileType === 'pptx' && check.id === 'media-captions' && matches.length === 0 && (scanDiagnostics?.pptxMediaCount ?? 0) === 0) {
      status = 'passed'
    }

    if (!backendStatus && fileType === 'pptx' && check.id === 'reading-order' && matches.length === 0) {
      status = 'manual'
      note = 'סדר הקריאה נבדק בהיוריסטיקה בלבד; בדקו את חלונית Reading Order ב-PowerPoint לפני פרסום.'
    }

    if (!backendStatus && fileType === 'pptx' && check.id === 'restricted-access') {
      status = 'manual'
      note = 'בדיקת הרשאות והגבלות מסמך אינה מיושמת בסורק ה-PPTX.'
    }

    return { ...check, matches, count: backendCount ?? issueCount(matches), status, note }
  })

  const passedChecks = checkResults.filter((check) => check.status === 'passed')
  const manualChecks = checkResults.filter((check) => check.status === 'manual')
  const partialChecks = checkResults.filter((check) => check.status === 'partial' || check.status === 'not_checked')

  if (import.meta.env.VITE_ACCESSIOFFICE_DEBUG_SCAN === 'true') {
    console.log('[ASSISTANT PANEL STATES]', {
      scannerVersion,
      scanDiagnostics,
      checkStatuses,
      states: checkResults.map(({ id, label, status, count }) => ({ id, label, status, count })),
    })
  }

  const noIssues = issues.length === 0 && quickFix.length === 0 && manualChecks.length === 0 && partialChecks.length === 0

  return (
    <div className="assistant-panel" dir="rtl">
      {/* Header */}
      <div className="ap-header">
        <span className="ap-header__title">בדיקת נגישות</span>
        <span className="ap-header__subtitle">
          {noIssues
            ? 'כל הבדיקות עברו'
            : issues.length === 0 && quickFix.length === 0
              ? 'לא נמצאו בעיות אוטומטיות · יש בדיקות לסקירה ידנית'
            : `${issues.length} סוגי בעיות · ${totalOccurrences} מופעים`}
        </span>
      </div>

      {/* Real accessibility issues */}
      {issues.length > 0 && (
        <section className="ap-section">
          <div className="ap-section__label">
            בעיות נגישות
            <span className="ap-section__tally">
              {highOccurrences > 0 && <span className="ap-tally ap-tally--high">{highOccurrences} גבוהה</span>}
              {mediumOccurrences > 0 && <span className="ap-tally ap-tally--medium">{mediumOccurrences} בינונית</span>}
            </span>
          </div>
          {issues.map((issue, index) => (
            <IssueRow
              key={`${issue.id}-${issue.location}-${index}`}
              issue={issue}
              isExpanded={expanded.has(issue.id)}
              onToggle={() => toggle(issue.id)}
            />
          ))}
        </section>
      )}

      {/* Quick Fix suggestions */}
      {quickFix.length > 0 && (
        <section className="ap-section ap-section--qf">
          <div className="ap-section__label">
            הצעות Quick Fix
            <span className="ap-tally ap-tally--qf">{qfOccurrences}</span>
          </div>
          {quickFix.map((issue, index) => (
            <IssueRow
              key={`${issue.id}-${issue.location}-${index}`}
              issue={issue}
              isExpanded={expanded.has(issue.id)}
              onToggle={() => toggle(issue.id)}
              isQuickFix
            />
          ))}
        </section>
      )}

      {/* Passed checks */}
      {passedChecks.length > 0 && (
        <section className="ap-section ap-section--passed">
          <div className="ap-section__label">בדיקות שעברו</div>
          {passedChecks.map((check) => (
            <PassedRow key={`passed-${check.id ?? check.label}`} label={check.label} />
          ))}
        </section>
      )}

      {manualChecks.length > 0 && (
        <section className="ap-section ap-section--manual">
          <div className="ap-section__label">בדיקות שמצריכות סקירה ידנית</div>
          {manualChecks.map((check) => (
            <ManualReviewRow key={`manual-${check.id ?? check.label}`} label={check.label} note={check.note} />
          ))}
        </section>
      )}

      {partialChecks.length > 0 && (
        <section className="ap-section ap-section--manual">
          <div className="ap-section__label">בדיקות חלקיות או שלא הושלמו</div>
          {partialChecks.map((check) => (
            <ManualReviewRow key={`partial-${check.id ?? check.label}`} label={check.label} note={check.note} />
          ))}
        </section>
      )}

      {noIssues && (
        <div className="ap-all-pass">
          <span className="ap-marker ap-marker--pass" aria-hidden="true" />
          לא נמצאו בעיות נגישות. מומלץ לבצע בדיקה ידנית לפני פרסום.
        </div>
      )}
    </div>
  )
}
