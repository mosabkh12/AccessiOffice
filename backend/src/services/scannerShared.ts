import { getWcagRule, type WcagRuleKey } from './wcagRules.js'
import { OFFICE_TITLES } from './officeCategories.js'
import type { Confidence, ScanIssue, Severity } from '../types/scan.types.js'

export type IssueTier = 'blocking' | 'quickFix'

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
  confidence: Confidence
  /** Office-like: Quick Fix suggestions are not counted as main accessibility errors. */
  issueTier?: IssueTier
}

export function isQuickFixDraft(draft: Pick<IssueDraft, 'title' | 'issueTier'>): boolean {
  return draft.issueTier === 'quickFix' || draft.title === OFFICE_TITLES.decorativeCandidates
}

export function isBlockingIssue(issue: ScanIssue): boolean {
  return issue.title !== OFFICE_TITLES.decorativeCandidates
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
  titleCount?: number
  duplicateTitleSlides?: number[]
  imageLikeObjectCount?: number
  missingAltCount?: number
  decorativeCandidateCount?: number
  readingOrderRiskSlides?: number[]
  mainIssueTotal?: number
  quickFixTotal?: number
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
    confidence: draft.confidence,
  }
}

export function calculateScore(issues: ScanIssue[]): number {
  const blocking = issues.filter(isBlockingIssue)
  const weights = { High: 12, Medium: 7, Low: 3 } as const
  const penalty = blocking.reduce((sum, i) => sum + weights[i.severity], 0)
  return Math.max(0, Math.min(100, 100 - penalty))
}

export function buildSummary(issues: ScanIssue[]) {
  const blocking = issues.filter(isBlockingIssue)
  const quickFix = issues.filter((i) => !isBlockingIssue(i))
  return {
    totalIssues: blocking.length,
    mainIssueTotal: blocking.length,
    quickFixTotal: quickFix.length,
    high: blocking.filter((i) => i.severity === 'High').length,
    medium: blocking.filter((i) => i.severity === 'Medium').length,
    low: blocking.filter((i) => i.severity === 'Low').length,
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
