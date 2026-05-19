/** Display order matching Microsoft Office Accessibility Checker groups */
export const OFFICE_CATEGORY_ORDER = [
  'צבע וניגודיות',
  'מדיה ואיורים',
  'טבלאות',
  'מבנה מסמך',
  'גישה למסמך',
  'מערכת',
]

export function groupIssuesByCategory(issues) {
  const map = new Map()
  for (const issue of issues) {
    const cat = issue.category || 'כללי'
    if (!map.has(cat)) map.set(cat, [])
    map.get(cat).push(issue)
  }

  const grouped = []
  for (const cat of OFFICE_CATEGORY_ORDER) {
    if (map.has(cat)) {
      grouped.push({ category: cat, issues: map.get(cat) })
      map.delete(cat)
    }
  }
  for (const [category, list] of map) {
    grouped.push({ category, issues: list })
  }
  return grouped
}
