import { useState } from 'react'
import { buildAssistantChecklist } from '../data/powerpointAssistantChecks.js'

function CheckRow({ check }) {
  const [open, setOpen] = useState(false)
  const hasFailures = !check.passed

  return (
    <li className={`assistant-check ${check.passed ? 'assistant-check--pass' : 'assistant-check--fail'}`}>
      <button
        type="button"
        className="assistant-check__btn"
        onClick={() => hasFailures && setOpen((v) => !v)}
        disabled={!hasFailures}
        aria-expanded={hasFailures ? open : undefined}
      >
        <span className="assistant-check__status" aria-hidden="true">
          {check.passed ? '✓' : check.count}
        </span>
        <span className="assistant-check__title">{check.title}</span>
        {!check.passed && (
          <span className="assistant-check__count">{check.count}</span>
        )}
      </button>
      {hasFailures && open && (
        <ul className="assistant-check__locations">
          {check.issues.map((issue) => (
            <li key={issue.id}>{issue.location}</li>
          ))}
        </ul>
      )}
    </li>
  )
}

export default function AccessibilityAssistantPanel({ issues, fileType }) {
  if (fileType !== 'pptx') return null

  const groups = buildAssistantChecklist(issues)

  return (
    <section className="card assistant-panel" aria-label="מסייע נגישות PowerPoint">
      <h2 className="card-title">מסייע נגישות (PowerPoint)</h2>
      <p className="card-hint">
        רשימת בדיקות בסגנון Microsoft Accessibility Assistant. סימון ✓ = לא נמצאו ממצאים בסריקה זו.
      </p>
      <div className="assistant-groups">
        {groups.map((group) => (
          <div key={group.category} className="assistant-group">
            <h3 className="assistant-group__title">{group.category}</h3>
            <ul className="assistant-checks">
              {group.checks.map((check) => (
                <CheckRow key={check.title} check={check} />
              ))}
            </ul>
          </div>
        ))}
      </div>
    </section>
  )
}

