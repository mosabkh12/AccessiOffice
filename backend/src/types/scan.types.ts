export type FileType = 'docx' | 'pptx' | 'xlsx'
export type Severity = 'High' | 'Medium' | 'Low'
export type Confidence = 'High' | 'Medium' | 'Low'
export type UserType = 'document-author' | 'accessibility-auditor' | 'lecturer-institution'
export type ScanType = 'basic' | 'full'

export interface ScanIssue {
  id: string
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
  confidence: Confidence
}

export interface ScanSummary {
  totalIssues: number
  mainIssueTotal?: number
  quickFixTotal?: number
  high: number
  medium: number
  low: number
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
  /** Debug: unique id per scan request */
  scanId?: string
  /** Debug: SHA-1 of uploaded file bytes */
  debugFileHash?: string
  /** Debug: active PPTX scanner build label */
  scannerVersion?: string
}
