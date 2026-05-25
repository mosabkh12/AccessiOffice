import { useState } from 'react'

function OccurrenceBadge({ count }) {
  if (!count || count <= 1) return null
  return <span className="occurrence-badge">{count}</span>
}

function LocationsList({ locations }) {
  if (!locations?.length) return null
  return (
    <ul className="locations-list">
      {locations.map((loc, i) => (
        <li key={i}>{loc}</li>
      ))}
    </ul>
  )
}

function IssueReportCard({ issue }) {
  return (
    <article className="report-issue-card">
      <header className="report-issue-card__header">
        <h4>
          {issue.title}
          <OccurrenceBadge count={issue.occurrenceCount} />
        </h4>
        <span className={`severity-tag ${issue.severity}`}>{issue.severityLabel}</span>
      </header>
      <dl className="report-issue-card__meta">
        <div>
          <dt>WCAG</dt>
          <dd className="wcag-ltr" dir="ltr">{issue.wcagCriterion}</dd>
        </div>
        <div>
          <dt>רמה</dt>
          <dd className="wcag-ltr" dir="ltr">Level {issue.wcagLevel}</dd>
        </div>
        <div>
          <dt>מיקום</dt>
          <dd>{issue.location}</dd>
        </div>
        {issue.occurrenceCount > 1 && (
          <div>
            <dt>מופעים</dt>
            <dd>{issue.occurrenceCount}</dd>
          </div>
        )}
      </dl>
      <dl className="report-issue-card__body">
        <div>
          <dt>משתמשים מושפעים</dt>
          <dd>{issue.affectedUsersText}</dd>
        </div>
        <div>
          <dt>השפעה</dt>
          <dd>{issue.impact}</dd>
        </div>
        {issue.locations?.length > 1 && (
          <div>
            <dt>מיקומים</dt>
            <dd><LocationsList locations={issue.locations} /></dd>
          </div>
        )}
        <div>
          <dt>המלצה</dt>
          <dd>{issue.recommendation}</dd>
        </div>
        <div>
          <dt>הסבר WCAG</dt>
          <dd>
            <p className="wcag-ltr wcag-meta" dir="ltr">
              {issue.wcagCriterion} · {issue.wcagPrinciple}
            </p>
            <p>{issue.wcagExplanation}</p>
          </dd>
        </div>
      </dl>
    </article>
  )
}

export default function IssuesTable({ issues, variant = 'interactive' }) {
  const [expandedId, setExpandedId] = useState(null)

  if (!issues?.length) {
    return <p>לא נמצאו פריטים.</p>
  }

  if (variant === 'report') {
    return (
      <div className="report-issues-grid">
        {issues.map((issue, index) => (
          <IssueReportCard key={`${issue.id}-${issue.location}-${index}`} issue={issue} />
        ))}
      </div>
    )
  }

  function toggle(id) {
    setExpandedId((prev) => (prev === id ? null : id))
  }

  return (
    <div className="issues-list">
      <div className="issues-table-compact" role="table" aria-label="טבלת בעיות נגישות">
        <div className="issues-table-head" role="row">
          <span role="columnheader">בעיה</span>
          <span role="columnheader">חומרה</span>
          <span role="columnheader">WCAG</span>
          <span role="columnheader">מיקום</span>
          <span role="columnheader">פעולה</span>
        </div>

        {issues.map((issue, index) => {
          const isOpen = expandedId === issue.id
          return (
            <article key={`${issue.id}-${issue.location}-${index}`} className="issue-row-wrap">
              <div className="issue-row" role="row">
                <span className="issue-title" role="cell">
                  {issue.title}
                  <OccurrenceBadge count={issue.occurrenceCount} />
                </span>
                <span role="cell">
                  <span className={`severity-tag ${issue.severity}`}>
                    {issue.severityLabel}
                  </span>
                </span>
                <span className="wcag-cell wcag-ltr" role="cell" dir="ltr">
                  <span className="wcag-criterion">{issue.wcagCriterion}</span>
                  <span className="wcag-level">Level {issue.wcagLevel}</span>
                </span>
                <span className="issue-location" role="cell">{issue.location}</span>
                <span role="cell">
                  <button
                    type="button"
                    className="btn-details"
                    onClick={() => toggle(issue.id)}
                    aria-expanded={isOpen}
                    aria-controls={`issue-details-${issue.id}`}
                  >
                    {isOpen ? 'סגירה' : 'פרטים'}
                  </button>
                </span>
              </div>

              {isOpen && (
                <div
                  id={`issue-details-${issue.id}`}
                  className="issue-details-card"
                  role="region"
                  aria-label={`פרטי בעיה: ${issue.title}`}
                >
                  <dl className="issue-details-dl">
                    <div>
                      <dt>משתמשים מושפעים</dt>
                      <dd>{issue.affectedUsersText}</dd>
                    </div>
                    <div>
                      <dt>השפעה</dt>
                      <dd>{issue.impact}</dd>
                    </div>
                    {issue.locations?.length > 1 && (
                      <div>
                        <dt>מיקומים ({issue.occurrenceCount})</dt>
                        <dd><LocationsList locations={issue.locations} /></dd>
                      </div>
                    )}
                    <div>
                      <dt>המלצה</dt>
                      <dd>{issue.recommendation}</dd>
                    </div>
                    <div>
                      <dt>הסבר WCAG</dt>
                      <dd>
                        <p className="wcag-ltr wcag-meta" dir="ltr">
                          {issue.wcagCriterion} · {issue.wcagPrinciple} · Level {issue.wcagLevel}
                        </p>
                        <p>{issue.wcagExplanation}</p>
                      </dd>
                    </div>
                  </dl>
                </div>
              )}
            </article>
          )
        })}
      </div>
    </div>
  )
}
