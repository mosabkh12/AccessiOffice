import type { WcagRuleKey } from '../data/wcag.rules.js'
import { getWcagRule } from '../data/wcag.rules.js'
import type { ScanIssue, ScanSummary, Severity } from '../types/scan.types.js'

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
  /** How many individual occurrences this issue covers */
  occurrenceCount?: number
  /** Per-occurrence locations for expanded display */
  locations?: string[]
  isQuickFix?: boolean
  confidence?: string
  issueTier?: string
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
  quickFixCount?: number
  // PPTX-specific diagnostic fields
  titleCount?: number
  duplicateTitleSlides?: number[]
  imageLikeObjectCount?: number
  missingAltCount?: number
  decorativeCandidateCount?: number
  readingOrderRiskSlides?: number[]
  contrastCheckStatus?: 'checked' | 'partial' | 'not_checked'
  contrastCheckedRuns?: number
  contrastSkippedRuns?: number
  pptxTableCount?: number
  pptxTablesMissingHeaderCount?: number
  pptxMergedCellCount?: number
  pptxMediaCount?: number
  pptxMediaMissingCaptionsCount?: number
  mainIssueTotal?: number
  quickFixTotal?: number
}

export function buildScanIssue(draft: IssueDraft): ScanIssue {
  const wcag = getWcagRule(draft.wcagKey)
  return {
    id: draft.id,
    wcagKey: draft.wcagKey,
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
    ...(draft.occurrenceCount !== undefined ? { occurrenceCount: draft.occurrenceCount } : {}),
    ...(draft.locations ? { locations: draft.locations } : {}),
    ...(draft.isQuickFix || draft.issueTier === 'quickFix' ? { isQuickFix: true as const } : {}),
    ...(draft.confidence ? { confidence: draft.confidence } : {}),
  }
}

function occ(issue: ScanIssue): number {
  return issue.occurrenceCount ?? 1
}

/**
 * Occurrence-aware score calculation.
 * Each issue group contributes min(cap, count × rate) to the penalty.
 * Quick Fix items do not affect the score.
 */
export function calculateScore(issues: ScanIssue[]): number {
  const weights = {
    High:   { rate: 2, cap: 30 },
    Medium: { rate: 4, cap: 16 },
    Low:    { rate: 1, cap: 8  },
  } as const
  const penalty = issues.reduce((sum, issue) => {
    const w = weights[issue.severity as keyof typeof weights]
    if (!w) return sum
    return sum + Math.min(w.cap, occ(issue) * w.rate)
  }, 0)
  return Math.max(0, Math.min(100, 100 - penalty))
}

export function buildSummary(issues: ScanIssue[], quickFix: ScanIssue[] = []): ScanSummary {
  const high   = issues.filter((i) => i.severity === 'High')
  const medium = issues.filter((i) => i.severity === 'Medium')
  const low    = issues.filter((i) => i.severity === 'Low')

  const sumOcc = (arr: ScanIssue[]) => arr.reduce((s, i) => s + occ(i), 0)

  const totalIssueTypes    = issues.length
  const highIssueTypes     = high.length
  const mediumIssueTypes   = medium.length
  const lowIssueTypes      = low.length
  const totalOccurrences   = sumOcc(issues)
  const highOccurrences    = sumOcc(high)
  const mediumOccurrences  = sumOcc(medium)
  const lowOccurrences     = sumOcc(low)
  const quickFixTypes      = quickFix.length
  const quickFixOccurrences = sumOcc(quickFix)

  return {
    totalIssueTypes,
    highIssueTypes,
    mediumIssueTypes,
    lowIssueTypes,
    totalOccurrences,
    highOccurrences,
    mediumOccurrences,
    lowOccurrences,
    quickFixTypes,
    quickFixOccurrences,
    // Backward-compatible aliases
    totalIssues: totalIssueTypes,
    high: highIssueTypes,
    medium: mediumIssueTypes,
    low: lowIssueTypes,
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
