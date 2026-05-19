import type { WcagRuleKey } from '../data/wcag.rules.js'
import { getWcagRule } from '../data/wcag.rules.js'
import type {
  Confidence,
  FileType,
  ScanIssue,
  ScanResult,
  ScanType,
  Severity,
  UserType,
} from '../types/scan.types.js'
import {
  CAT_COLOR_CONTRAST,
  CAT_DOCUMENT_STRUCTURE,
  CAT_MEDIA,
  CAT_TABLES,
  OFFICE_TITLES,
} from './officeCategories.js'
import { calculateScore, logScanSummary, scanOfficeFile } from './officeScanner.service.js'
import { PPTX_SCANNER_BUILD } from './pptxScanner.service.js'
import { buildScanIssue, buildSummary, type ExtractedSignals, type IssueDraft } from './scannerShared.js'

interface IssueTemplate {
  id: string
  wcagKey: WcagRuleKey | string
  title: string
  severity: Severity
  category: string
  affectedUsers: string[]
  impact: string
  recommendation: string
  confidence: Confidence
}

const DEMO_LOCATIONS: Record<FileType, string[]> = {
  docx: ['עמוד 1 / תמונה 2', 'עמוד 1 / כותרות', 'עמוד 4 / קישור', 'עמוד 3 / טבלה', 'עמוד 2 / גוף טקסט'],
  pptx: ['שקופית 2', 'שקופית 3 / תמונה 1', 'שקופית 5 / תיבת טקסט', 'שקופית 4 / תוכן', 'שקופית 6'],
  xlsx: ['גיליון "Sheet1" / A1:F20', 'גיליון "Sheet1" / B5:D8', 'גיליון "Sheet1"', 'גיליון "Sheet2"', 'גיליון "Sheet1" / C3'],
}

const DEMO_DOCX: IssueTemplate[] = [
  { id: 'missing-alt-text', wcagKey: 'non-text-content', title: OFFICE_TITLES.missingAltText, severity: 'High', category: CAT_MEDIA, affectedUsers: ['משתמשים עיוורים', 'משתמשי קורא מסך'], impact: 'תמונות ללא תיאור חלופי.', recommendation: 'הוסיפו טקסט חלופי.', confidence: 'High' },
  { id: 'incorrect-heading-structure', wcagKey: 'headings-labels', title: OFFICE_TITLES.noHeadingsInDocument, severity: 'High', category: CAT_DOCUMENT_STRUCTURE, affectedUsers: ['משתמשי קורא מסך'], impact: 'מבנה כותרות לקוי.', recommendation: 'השתמשו בסגנונות כותרת.', confidence: 'Medium' },
  { id: 'unclear-link-text', wcagKey: 'link-purpose', title: OFFICE_TITLES.unclearHyperlink, severity: 'Medium', category: CAT_DOCUMENT_STRUCTURE, affectedUsers: ['משתמשי קורא מסך'], impact: 'קישור "לחץ כאן".', recommendation: 'טקסט קישור תיאורי.', confidence: 'High' },
  { id: 'table-without-header', wcagKey: 'info-relationships', title: OFFICE_TITLES.missingTableHeader, severity: 'Medium', category: CAT_TABLES, affectedUsers: ['משתמשי קורא מסך'], impact: 'טבלה ללא כותרת.', recommendation: 'סמנו שורת כותרת.', confidence: 'High' },
  { id: 'low-color-contrast', wcagKey: 'contrast-minimum', title: OFFICE_TITLES.hardTextContrast, severity: 'High', category: CAT_COLOR_CONTRAST, affectedUsers: ['משתמשים עם לקות ראייה'], impact: 'ניגודיות נמוכה.', recommendation: 'שפרו ניגודיות.', confidence: 'Low' },
]

const DEMO_PPTX: IssueTemplate[] = [
  { id: 'slide-without-title', wcagKey: 'page-titled-headings', title: OFFICE_TITLES.slideWithoutTitle, severity: 'Medium', category: CAT_DOCUMENT_STRUCTURE, affectedUsers: ['משתמשי קורא מסך'], impact: 'שקופיות ללא כותרת.', recommendation: 'הוסיפו כותרת.', confidence: 'High' },
  { id: 'image-without-alt-text', wcagKey: 'non-text-content', title: OFFICE_TITLES.missingAltText, severity: 'High', category: CAT_MEDIA, affectedUsers: ['משתמשים עיוורים'], impact: 'תמונות ללא alt.', recommendation: 'הוסיפו alt.', confidence: 'High' },
  { id: 'reading-order-issue', wcagKey: 'meaningful-sequence', title: OFFICE_TITLES.readingOrderCheck, severity: 'Medium', category: CAT_DOCUMENT_STRUCTURE, affectedUsers: ['משתמשי קורא מסך'], impact: 'סדר קריאה.', recommendation: 'בדקו סדר קריאה.', confidence: 'Medium' },
  { id: 'unclear-link', wcagKey: 'link-purpose', title: OFFICE_TITLES.unclearHyperlink, severity: 'Medium', category: CAT_DOCUMENT_STRUCTURE, affectedUsers: ['משתמשי קורא מסך'], impact: 'קישור לא ברור.', recommendation: 'טקסט תיאורי.', confidence: 'High' },
  { id: 'small-font-size', wcagKey: 'resize-text', title: OFFICE_TITLES.smallFont, severity: 'Low', category: CAT_COLOR_CONTRAST, affectedUsers: ['משתמשים עם לקות ראייה'], impact: 'גופן קטן.', recommendation: 'הגדילו גופן.', confidence: 'High' },
]

const DEMO_XLSX: IssueTemplate[] = [
  { id: 'table-without-header', wcagKey: 'info-relationships', title: OFFICE_TITLES.missingTableHeader, severity: 'High', category: CAT_TABLES, affectedUsers: ['משתמשי קורא מסך'], impact: 'ללא כותרת.', recommendation: 'הגדירו כותרת.', confidence: 'High' },
  { id: 'merged-cells', wcagKey: 'info-relationships', title: OFFICE_TITLES.mergedOrSplitCells, severity: 'High', category: CAT_TABLES, affectedUsers: ['משתמשי קורא מסך'], impact: 'מיזוג תאים.', recommendation: 'בטלו מיזוג.', confidence: 'High' },
  { id: 'empty-cells', wcagKey: 'info-relationships', title: OFFICE_TITLES.emptyCellsInRange, severity: 'Medium', category: CAT_TABLES, affectedUsers: ['משתמשי קורא מסך'], impact: 'תאים ריקים.', recommendation: 'צמצמו ריווח.', confidence: 'High' },
  { id: 'sheet-without-title', wcagKey: 'headings-labels', title: OFFICE_TITLES.genericSheetName, severity: 'Medium', category: CAT_DOCUMENT_STRUCTURE, affectedUsers: ['משתמשי קורא מסך'], impact: 'Sheet1.', recommendation: 'שם תיאורי.', confidence: 'High' },
  { id: 'low-contrast', wcagKey: 'contrast-minimum', title: OFFICE_TITLES.hardTextContrast, severity: 'High', category: CAT_COLOR_CONTRAST, affectedUsers: ['משתמשים עם לקות ראייה'], impact: 'ניגודיות.', recommendation: 'שפרו צבעים.', confidence: 'Low' },
]

export const CRITICAL_ANALYSIS_HE =
  'סריקה אוטומטית מסייעת לזהות בעיות נגישות נפוצות, אך אינה מחליפה בדיקה אנושית מקצועית. יש לבדוק ידנית את איכות הטקסט החלופי, הקשר התוכן, סדר קריאה מורכב והתאמה מלאה לתקן הישראלי לפני פרסום המסמך.'

const LIMITED_SCAN_NOTE =
  'הסריקה האוטומטית זיהתה מעט נתונים מבניים. מומלץ לבצע גם בדיקה ידנית.'

const NO_ISSUES_NOTE =
  'לא נמצאו בעיות נגישות משמעותיות בסריקה האוטומטית. מומלץ לבצע בדיקה ידנית לפני פרסום.'

function isDemoFile(fileName: string): boolean {
  const lower = fileName.toLowerCase()
  return (
    lower.includes('sample-bad') ||
    lower.includes('bad-accessibility') ||
    /(?:^|[/\\._-])demo(?:[/\\._-]|\.)/i.test(lower)
  )
}

function buildDemoIssues(fileType: FileType): ScanIssue[] {
  const templates = fileType === 'docx' ? DEMO_DOCX : fileType === 'pptx' ? DEMO_PPTX : DEMO_XLSX
  const locations = DEMO_LOCATIONS[fileType]
  return templates.map((t, i) => {
    const wcag = getWcagRule(t.wcagKey)
    return {
      id: t.id,
      title: t.title,
      severity: t.severity,
      category: t.category,
      wcagPrinciple: wcag.principle,
      wcagCriterion: wcag.criterion,
      wcagLevel: wcag.level,
      wcagExplanation: wcag.explanationHe,
      affectedUsers: t.affectedUsers,
      impact: t.impact,
      recommendation: t.recommendation,
      location: locations[i] ?? 'מסמך',
      confidence: t.confidence,
    }
  })
}

function hasMeaningfulSignals(signals: ExtractedSignals, fileType: FileType): boolean {
  if (fileType === 'pptx') return (signals.slideCount ?? 0) > 0
  if (fileType === 'docx') return (signals.paragraphCount ?? 0) > 0
  if (fileType === 'xlsx') return (signals.sheetCount ?? 0) > 0
  return false
}

function buildFallbackDrafts(signals: ExtractedSignals, fileName: string): IssueDraft[] {
  const meta: string[] = []
  if (signals.slideCount != null) meta.push(`${signals.slideCount} שקופיות`)
  if (signals.paragraphCount != null) meta.push(`${signals.paragraphCount} פסקאות`)
  if (signals.sheetCount != null) meta.push(`${signals.sheetCount} גיליונות`)
  if (signals.fileSize != null) meta.push(`${Math.round(signals.fileSize / 1024)}KB`)

  return [
    {
      id: 'limited-scan-metadata',
      wcagKey: 'info-relationships',
      title: OFFICE_TITLES.limitedScan,
      severity: 'Low',
      category: 'מערכת',
      affectedUsers: ['כל המשתמשים'],
      impact: `${LIMITED_SCAN_NOTE} מטא-נתונים: ${meta.join(', ') || fileName}.`,
      recommendation: 'בצעו בדיקת נגישות ידנית מלאה.',
      location: 'מטא-נתוני קובץ',
      confidence: 'Low',
    },
  ]
}

export async function generateScanResult(
  filePath: string,
  fileName: string,
  fileType: FileType,
  userType: UserType,
  scanType: ScanType,
  fileSize: number,
  debug?: { scanId: string; debugFileHash: string },
): Promise<ScanResult> {
  const scanMeta = {
    scanId: debug?.scanId ?? String(Date.now()),
    debugFileHash: debug?.debugFileHash ?? '',
    scannerVersion: fileType === 'pptx' ? PPTX_SCANNER_BUILD : undefined,
  }

  if (isDemoFile(fileName)) {
    console.warn('[SCAN SERVICE] DEMO file path — static issues, NOT real scanner', { fileName })
    const issues = buildDemoIssues(fileType)
    const score = calculateScore(issues)
    logScanSummary(fileName, { fileType, fileSize }, issues.length, score)
    return {
      fileName,
      fileType,
      userType,
      scanType,
      score,
      summary: buildSummary(issues),
      issues,
      criticalAnalysis: CRITICAL_ANALYSIS_HE,
      ...scanMeta,
    }
  }

  let issues: ScanIssue[]
  let signals: ExtractedSignals = { fileType, fileSize }

  try {
    if (fileType === 'pptx') {
      console.log('[SCAN SERVICE] Calling PPTX scanner', {
        version: PPTX_SCANNER_BUILD,
        filePath,
        fileSize,
        scanId: scanMeta.scanId,
      })
    }
    const result = await scanOfficeFile(filePath, fileType, fileSize)
    signals = result.signals

    if (!hasMeaningfulSignals(signals, fileType)) {
      issues = buildFallbackDrafts(signals, fileName).map(buildScanIssue)
    } else {
      issues = result.issues
    }
  } catch (err) {
    console.error('[AccessiOffice scan] error:', (err as Error).message)
    issues = buildFallbackDrafts(signals, fileName).map(buildScanIssue)
  }

  const score = issues.length === 0 ? 100 : calculateScore(issues)
  logScanSummary(fileName, signals, issues.length, score)

  let criticalAnalysis = CRITICAL_ANALYSIS_HE
  if (issues.length === 0) {
    criticalAnalysis = `${NO_ISSUES_NOTE} ${CRITICAL_ANALYSIS_HE}`
  } else if (issues.some((i) => i.id === 'limited-scan-metadata')) {
    criticalAnalysis = `${LIMITED_SCAN_NOTE} ${CRITICAL_ANALYSIS_HE}`
  }

  return {
    fileName,
    fileType,
    userType,
    scanType,
    score,
    summary: buildSummary(issues),
    issues,
    criticalAnalysis,
    ...scanMeta,
  }
}

export function getFileType(filename: string): FileType | null {
  const ext = filename.split('.').pop()?.toLowerCase()
  if (ext === 'docx' || ext === 'pptx' || ext === 'xlsx') return ext
  return null
}

export function parseUserType(value: string): UserType | null {
  const allowed: UserType[] = ['document-author', 'accessibility-auditor', 'lecturer-institution']
  return allowed.includes(value as UserType) ? (value as UserType) : null
}

export function parseScanType(value: string): ScanType | null {
  return value === 'basic' || value === 'full' ? value : null
}
