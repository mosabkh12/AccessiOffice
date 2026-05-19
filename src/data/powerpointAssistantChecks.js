/** PowerPoint Accessibility Assistant–style checklist (Hebrew UI, Office groups). */
export const PPTX_ASSISTANT_GROUPS = [
  {
    category: 'צבע וניגודיות',
    checks: [{ title: 'ניגודיות טקסט נמוכה', aliases: ['ניגודיות טקסט קשה לקריאה'] }],
  },
  {
    category: 'מדיה ואיורים',
    checks: [
      { title: 'חסר טקסט חלופי' },
      { title: 'אובייקטים שניתן לסמן כדקורטיביים' },
      { title: 'חסרות כתוביות לאודיו/וידאו' },
    ],
  },
  {
    category: 'טבלאות',
    checks: [
      { title: 'חסרה כותרת טבלה' },
      { title: 'שימוש בתאים ממוזגים או מפוצלים' },
    ],
  },
  {
    category: 'מבנה מסמך',
    checks: [
      { title: 'שקופית ללא כותרת' },
      { title: 'כותרת שקופית כפולה' },
      { title: 'בדיקת סדר קריאה' },
      { title: 'שם מקטע ברירת מחדל' },
      { title: 'שם מקטע כפול' },
    ],
  },
  {
    category: 'גישה למסמך',
    checks: [{ title: 'גישה מוגבלת' }],
  },
]

function findIssuesForCheck(byTitle, check) {
  const titles = [check.title, ...(check.aliases ?? [])]
  const seen = new Set()
  const matches = []
  for (const t of titles) {
    for (const issue of byTitle.get(t) ?? []) {
      if (seen.has(issue.id)) continue
      seen.add(issue.id)
      matches.push(issue)
    }
  }
  return matches
}

export function buildAssistantChecklist(issues) {
  const byTitle = new Map()
  for (const issue of issues ?? []) {
    if (!byTitle.has(issue.title)) byTitle.set(issue.title, [])
    byTitle.get(issue.title).push(issue)
  }

  return PPTX_ASSISTANT_GROUPS.map((group) => ({
    category: group.category,
    checks: group.checks.map((check) => {
      const matches = findIssuesForCheck(byTitle, check)
      return {
        title: check.title,
        passed: matches.length === 0,
        count: matches.length,
        issues: matches,
      }
    }),
  }))
}

export function groupIssuesByTitle(issues) {
  const map = new Map()
  for (const issue of issues ?? []) {
    if (!map.has(issue.title)) map.set(issue.title, [])
    map.get(issue.title).push(issue)
  }
  return [...map.entries()].map(([title, items]) => ({
    title,
    count: items.length,
    issues: items,
  }))
}
