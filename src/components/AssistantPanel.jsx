import { useState } from 'react'

// ── Known checks per file type (used to infer "passed" status) ────────────────
// A check is "passed" when no issue in the issues array matches its id prefix.

const CHECKS_BY_TYPE = {
  pptx: [
    { label: 'חסר טקסט חלופי',        prefixes: ['missing-alt'],      group: 'מדיה ואיורים' },
    { label: 'שקופית ללא כותרת',       prefixes: ['missing-title'],    group: 'מבנה מסמך' },
    { label: 'כותרת שקופית כפולה',     prefixes: ['duplicate-title'],  group: 'מבנה מסמך' },
    { label: 'בדיקת סדר קריאה',       prefixes: ['reading-order'],    group: 'מבנה מסמך' },
    { label: 'ניגודיות טקסט',          prefixes: ['low-contrast', 'contrast'], group: 'צבע וניגודיות' },
    { label: 'טבלאות',                  prefixes: ['table'],            group: 'טבלאות' },
    { label: 'קישורים לא ברורים',      prefixes: ['unclear-link'],     group: 'ניווט' },
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
  return check.prefixes.some((p) => issue.id?.startsWith(p))
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

// ── Main component ─────────────────────────────────────────────────────────────

export default function AssistantPanel({ issues = [], quickFix = [], fileType = 'pptx' }) {
  const [expanded, setExpanded] = useState(new Set())

  function toggle(id) {
    setExpanded((prev) => {
      const next = new Set(prev)
      next.has(id) ? next.delete(id) : next.add(id)
      return next
    })
  }

  // Counts derived directly from the issues array — single source of truth
  const totalOccurrences   = issues.reduce((s, i) => s + (i.occurrenceCount ?? 1), 0)
  const highOccurrences    = issues.filter((i) => i.severity === 'high').reduce((s, i) => s + (i.occurrenceCount ?? 1), 0)
  const mediumOccurrences  = issues.filter((i) => i.severity === 'medium').reduce((s, i) => s + (i.occurrenceCount ?? 1), 0)
  const qfOccurrences      = quickFix.reduce((s, i) => s + (i.occurrenceCount ?? 1), 0)

  // Passed checks for this file type
  const knownChecks = CHECKS_BY_TYPE[fileType] ?? []
  const passedChecks = knownChecks.filter(
    (check) => !issues.some((issue) => matchesCheck(issue, check))
  )

  const noIssues = issues.length === 0 && quickFix.length === 0

  return (
    <div className="assistant-panel" dir="rtl">
      {/* Header */}
      <div className="ap-header">
        <span className="ap-header__title">בדיקת נגישות</span>
        <span className="ap-header__subtitle">
          {noIssues
            ? 'כל הבדיקות עברו'
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
          {issues.map((issue) => (
            <IssueRow
              key={issue.id}
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
          {quickFix.map((issue) => (
            <IssueRow
              key={issue.id}
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
            <PassedRow key={check.label} label={check.label} />
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
