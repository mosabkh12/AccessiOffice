import type { WcagRuleKey } from '../../data/wcag.rules.js'
import { getWcagRule } from '../../data/wcag.rules.js'
import type { ScanIssue, Severity } from '../../types/scan.types.js'

export interface IssueDraft {
  id: string
  wcagKey: WcagRuleKey | string
  title: string
  severity: Severity
  category: string
  affectedUsers: string[]
  impact: string
  recommendation: string
  location: string
}

export interface ExtractedSignals {
  fileType: string
  fileSize?: number
  slideCount?: number
  paragraphCount?: number
  sheetCount?: number
  imageCount?: number
  tableCount?: number
  hyperlinkCount?: number
  mergedCellCount?: number
  genericSheetNames?: string[]
  missingTitleSlides?: number[]
  emptyParagraphCount?: number
}

export function buildScanIssue(draft: IssueDraft): ScanIssue {
  const wcag = getWcagRule(draft.wcagKey)
  return {
    id: draft.id,
    title: draft.title,
    severity: draft.severity,
    category: draft.category,
    wcagPrinciple: wcag.principle,
    wcagCriterion: wcag.criterion,
    wcagLevel: wcag.level,
    wcagExplanation: wcag.explanationHe,
    affectedUsers: draft.affectedUsers,
    impact: draft.impact,
    recommendation: draft.recommendation,
    location: draft.location,
  }
}

export function calculateScore(issues: ScanIssue[]): number {
  const weights = { High: 12, Medium: 7, Low: 3 } as const
  const penalty = issues.reduce((sum, i) => sum + weights[i.severity], 0)
  return Math.max(0, Math.min(100, 100 - penalty))
}

export function buildSummary(issues: ScanIssue[]) {
  return {
    totalIssues: issues.length,
    high: issues.filter((i) => i.severity === 'High').length,
    medium: issues.filter((i) => i.severity === 'Medium').length,
    low: issues.filter((i) => i.severity === 'Low').length,
  }
}

export function dedupeDrafts(drafts: IssueDraft[]): IssueDraft[] {
  const seen = new Set<string>()
  return drafts.filter((d) => {
    const key = `${d.id}::${d.location}`
    if (seen.has(key)) return false
    seen.add(key)
    return true
  })
}
