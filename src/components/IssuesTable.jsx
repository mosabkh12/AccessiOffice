import { useState } from 'react'
import { groupIssuesByCategory } from '../data/officeCategories.js'
import { groupIssuesByTitle } from '../data/powerpointAssistantChecks.js'

function IssueReportCard({ issue }) {
  return (
    <article className="report-issue-card">
      <header className="report-issue-card__header">
        <h4>{issue.title}</h4>
        <div className="report-issue-card__badges">
          <span className={`severity-tag ${issue.severity}`}>{issue.severityLabel}</span>
          {issue.confidenceLabel && (
            <span className={`confidence-tag confidence-tag--${issue.confidence}`}>
              {issue.confidenceLabel}
            </span>
          )}
        </div>
      </header>
      <dl className="report-issue-card__meta">
        <div>
          <dt>קטגוריה</dt>
          <dd>{issue.category}</dd>
        </div>
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
        {issue.contextSnippet && (
          <div>
            <dt>קטע מהמסמך</dt>
            <dd>
              <blockquote className="issue-snippet">&quot;{issue.contextSnippet}&quot;</blockquote>
            </dd>
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

function InteractiveIssueRow({ issue, expandedId, onToggle }) {
  const isOpen = expandedId === issue.id
  return (
    <article className="issue-row-wrap">
      <div className="issue-row" role="row">
        <span className="issue-title" role="cell">{issue.title}</span>
        <span role="cell" className="issue-badges-cell">
          <span className={`severity-tag ${issue.severity}`}>{issue.severityLabel}</span>
          {issue.confidenceLabel && (
            <span className={`confidence-tag confidence-tag--${issue.confidence}`}>
              {issue.confidenceLabel}
            </span>
          )}
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
            onClick={() => onToggle(issue.id)}
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
              <dt>קטגוריה</dt>
              <dd>{issue.category}</dd>
            </div>
            <div>
              <dt>מיקום</dt>
              <dd>{issue.location}</dd>
            </div>
            {issue.contextSnippet && (
              <div>
                <dt>קטע מהמסמך</dt>
                <dd>
                  <blockquote className="issue-snippet">
                    &quot;{issue.contextSnippet}&quot;
                  </blockquote>
                </dd>
              </div>
            )}
            <div>
              <dt>משתמשים מושפעים</dt>
              <dd>{issue.affectedUsersText}</dd>
            </div>
            {issue.confidenceLabel && (
              <div>
                <dt>ודאות זיהוי</dt>
                <dd>{issue.confidenceLabel}</dd>
              </div>
            )}
            <div>
              <dt>השפעה</dt>
              <dd>{issue.impact}</dd>
            </div>
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
}

export default function IssuesTable({ issues, variant = 'interactive' }) {
  const [expandedId, setExpandedId] = useState(null)
  const grouped = groupIssuesByCategory(issues)

  if (!issues?.length) {
    return <p>לא נמצאו בעיות נגישות.</p>
  }

  function toggle(id) {
    setExpandedId((prev) => (prev === id ? null : id))
  }

  if (variant === 'report') {
    return (
      <div className="issues-by-category">
        {grouped.map(({ category, issues: catIssues }) => {
          const byTitle = groupIssuesByTitle(catIssues)
          return (
            <section key={category} className="issue-category-group">
              <h3 className="issue-category-heading">
                {category}
                <span className="issue-category-count">{catIssues.length}</span>
              </h3>
              {byTitle.map(({ title, count, issues: titleIssues }) => (
                <div key={title} className="report-title-group">
                  <h4 className="report-title-group__heading">
                    {title}
                    {count > 1 ? ` — ${count} מופעים` : ''}
                  </h4>
                  <ul className="report-title-locations">
                    {titleIssues.map((issue) => (
                      <li key={issue.id}>{issue.location}</li>
                    ))}
                  </ul>
                  <div className="report-issues-grid">
                    {titleIssues.map((issue) => (
                      <IssueReportCard key={issue.id} issue={issue} />
                    ))}
                  </div>
                </div>
              ))}
            </section>
          )
        })}
      </div>
    )
  }

  return (
    <div className="issues-list issues-by-category">
      {grouped.map(({ category, issues: catIssues }) => (
        <section key={category} className="issue-category-group">
          <h3 className="issue-category-heading">
            {category}
            <span className="issue-category-count">{catIssues.length}</span>
          </h3>
          <div className="issues-table-compact" role="table" aria-label={`בעיות: ${category}`}>
            <div className="issues-table-head" role="row">
              <span role="columnheader">בעיה</span>
              <span role="columnheader">חומרה</span>
              <span role="columnheader">WCAG</span>
              <span role="columnheader">מיקום</span>
              <span role="columnheader">פעולה</span>
            </div>
            {catIssues.map((issue) => (
              <InteractiveIssueRow
                key={issue.id}
                issue={issue}
                expandedId={expandedId}
                onToggle={toggle}
              />
            ))}
          </div>
        </section>
      ))}
    </div>
  )
}


