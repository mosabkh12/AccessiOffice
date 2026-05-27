export type FileType = 'docx' | 'pptx' | 'xlsx'
export type Severity = 'High' | 'Medium' | 'Low'
export type UserType = 'document-author' | 'accessibility-auditor' | 'lecturer-institution'
export type ScanType = 'basic' | 'full'

export interface ScanIssue {
  id: string
  wcagKey?: string
  title: string
  severity: Severity
  category: string
  wcagPrinciple: string
  wcagCriterion: string
  wcagLevel: string
  wcagExplanation: string
  affectedUsers: string[]
  impact: string
  recommendation: string
  location: string
  /** How many individual occurrences this grouped issue covers (default 1) */
  occurrenceCount?: number
  /** Specific per-occurrence locations for expanded display */
  locations?: string[]
  confidence?: string
  isQuickFix?: boolean
}

export interface ScanSummary {
  /** Number of grouped issue objects (issue types) */
  totalIssueTypes: number
  highIssueTypes: number
  mediumIssueTypes: number
  lowIssueTypes: number
  /** Sum of occurrenceCount across all real issues */
  totalOccurrences: number
  highOccurrences: number
  mediumOccurrences: number
  lowOccurrences: number
  /** Quick Fix decorative suggestions */
  quickFixTypes: number
  quickFixOccurrences: number
  /** Backward-compatible aliases */
  totalIssues: number
  high: number
  medium: number
  low: number
}

export interface ScanCheckStatus {
  id: string
  label: string
  status: 'failed' | 'passed' | 'partial' | 'not_checked' | 'manual'
  count: number
  note?: string
}

export type OfficeLikeSummaryStatus = 'failed' | 'passed' | 'manual' | 'partial' | 'not_checked'

export interface OfficeLikeSummaryItem {
  label: string
  status: OfficeLikeSummaryStatus
  count: number
  note?: string
}

export interface OfficeLikeSummary {
  contrast: OfficeLikeSummaryItem
  missingAltText: OfficeLikeSummaryItem
  mediaCaptions: OfficeLikeSummaryItem
  tableHeader: OfficeLikeSummaryItem
  mergedCells: OfficeLikeSummaryItem
  missingSlideTitle: OfficeLikeSummaryItem
  duplicateSlideTitle: OfficeLikeSummaryItem
  readingOrder: OfficeLikeSummaryItem
  restrictedAccess: OfficeLikeSummaryItem
}

export interface ScanResult {
  fileName: string
  fileType: FileType
  userType: UserType
  scanType: ScanType
  score: number
  summary: ScanSummary
  issues: ScanIssue[]
  criticalAnalysis: string
  quickFix?: ScanIssue[]
  /** SHA-256 of uploaded file bytes — changes when the file content changes */
  fileHash?: string
  scannerVersion?: string
  debugMarker?: string
  diagnostics?: Record<string, unknown>
  scanDiagnostics?: Record<string, unknown>
  checkStatuses?: ScanCheckStatus[]
  manualReviewChecks?: ScanCheckStatus[]
  officeLikeSummary?: OfficeLikeSummary
}
