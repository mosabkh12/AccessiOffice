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
  /** Per-occurrence item names extracted from the Office Accessibility Checker pane. */
  locations?: string[]
}

/**
 * Map of accessibility check key → item.
 * PPTX keys: contrast, missingAltText, mediaCaptions, tableHeader, mergedCells,
 *             missingSlideTitle, duplicateSlideTitle, readingOrder, restrictedAccess
 * DOCX keys: contrast, missingAltText, tableHeader, mergedCells,
 *             unclearHyperlinkText, documentTitle, restrictedAccess
 * The index signature allows both engines to share the same type.
 */
export interface OfficeLikeSummary {
  [key: string]: OfficeLikeSummaryItem
}

/** Which engine produced the officeLikeSummary counts */
export type ScanEngine = 'xml' | 'office-engine' | 'hybrid'

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
  /** Which engine produced officeLikeSummary: xml = XML-only, office-engine = PPT worker, hybrid = PPT worker failed + XML fallback */
  engine?: ScanEngine
  /** Set when the PowerPoint worker failed and XML fallback was used */
  workerError?: string
  /** Office version string reported by the COM worker (e.g. "16.0") */
  workerScannerVersion?: string
  /** true when PPTX occurrence extraction via UIAutomation+COM succeeded for at least one key */
  officeOccurrencesExtracted?: boolean
  /** Explanation when officeOccurrencesExtracted is false */
  officeOccurrencesNote?: string | null
}
