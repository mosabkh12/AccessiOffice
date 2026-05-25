import type { WcagRuleKey } from '../data/wcag.rules.js'
import { getWcagRule } from '../data/wcag.rules.js'
import type { FileType, ScanIssue, ScanResult, ScanType, Severity, UserType } from '../types/scan.types.js'
import { analyzeDocx } from './docxScanner.service.js'
import {
  buildScanIssue,
  buildSummary,
  calculateScore,
  dedupeDrafts,
  type ExtractedSignals,
  type IssueDraft,
} from './scannerShared.js'
import { analyzePptx } from './pptxScanner.service.js'
import { analyzeXlsx } from './xlsxScanner.service.js'

interface IssueTemplate {
  id: string
  wcagKey: WcagRuleKey | string
  title: string
  severity: Severity
  category: string
  affectedUsers: string[]
  impact: string
  recommendation: string
}

const DEMO_LOCATIONS: Record<FileType, string[]> = {
  docx: ['עמוד 1 / תמונה 2', 'עמוד 1 / כותרות', 'עמוד 4 / קישור', 'עמוד 3 / טבלה', 'עמוד 2 / גוף טקסט'],
  pptx: ['שקופית 2', 'שקופית 3 / תמונה 1', 'שקופית 5 / תיבת טקסט', 'שקופית 4 / תוכן', 'שקופית 6'],
  xlsx: ['גיליון "Sheet1" / A1:F20', 'גיליון "Sheet1" / B5:D8', 'גיליון "Sheet1"', 'גיליון "Sheet2"', 'גיליון "Sheet1" / C3'],
}

const DEMO_DOCX: IssueTemplate[] = [
  { id: 'missing-alt-text', wcagKey: 'non-text-content', title: 'תמונה ללא טקסט חלופי', severity: 'High', category: 'תוכן', affectedUsers: ['משתמשים עיוורים', 'משתמשי קורא מסך'], impact: 'תמונות ללא תיאור חלופי.', recommendation: 'הוסיפו טקסט חלופי.', },
  { id: 'incorrect-heading-structure', wcagKey: 'headings-labels', title: 'מבנה כותרות לא תקין', severity: 'High', category: 'מבנה', affectedUsers: ['משתמשי קורא מסך'], impact: 'מבנה כותרות לקוי.', recommendation: 'השתמשו בסגנונות כותרת.', },
  { id: 'unclear-link-text', wcagKey: 'link-purpose', title: 'קישור לא ברור', severity: 'Medium', category: 'ניווט', affectedUsers: ['משתמשי קורא מסך'], impact: 'קישור "לחץ כאן".', recommendation: 'טקסט קישור תיאורי.', },
  { id: 'table-without-header', wcagKey: 'info-relationships', title: 'טבלה ללא כותרות ברורות', severity: 'Medium', category: 'טבלאות', affectedUsers: ['משתמשי קורא מסך'], impact: 'טבלה ללא כותרת.', recommendation: 'סמנו שורת כותרת.', },
  { id: 'low-color-contrast', wcagKey: 'contrast-minimum', title: 'ניגודיות צבעים נמוכה', severity: 'High', category: 'תצוגה', affectedUsers: ['משתמשים עם לקות ראייה'], impact: 'ניגודיות נמוכה.', recommendation: 'שפרו ניגודיות.', },
]

const DEMO_PPTX: IssueTemplate[] = [
  { id: 'slide-without-title', wcagKey: 'page-titled-headings', title: 'שקופית ללא כותרת', severity: 'Medium', category: 'מבנה', affectedUsers: ['משתמשי קורא מסך'], impact: 'שקופיות ללא כותרת.', recommendation: 'הוסיפו כותרת.', },
  { id: 'image-without-alt-text', wcagKey: 'non-text-content', title: 'תמונה ללא טקסט חלופי', severity: 'High', category: 'תוכן', affectedUsers: ['משתמשים עיוורים'], impact: 'תמונות ללא alt.', recommendation: 'הוסיפו alt.', },
  { id: 'reading-order-issue', wcagKey: 'meaningful-sequence', title: 'בעיית סדר קריאה', severity: 'Medium', category: 'מבנה', affectedUsers: ['משתמשי קורא מסך'], impact: 'סדר קריאה.', recommendation: 'בדקו סדר קריאה.', },
  { id: 'unclear-link', wcagKey: 'link-purpose', title: 'קישור לא ברור', severity: 'Medium', category: 'ניווט', affectedUsers: ['משתמשי קורא מסך'], impact: 'קישור לא ברור.', recommendation: 'טקסט תיאורי.', },
  { id: 'small-font-size', wcagKey: 'resize-text', title: 'גודל גופן קטן מדי', severity: 'Low', category: 'תצוגה', affectedUsers: ['משתמשים עם לקות ראייה'], impact: 'גופן קטן.', recommendation: 'הגדילו גופן.', },
]

const DEMO_XLSX: IssueTemplate[] = [
  { id: 'table-without-header', wcagKey: 'info-relationships', title: 'טבלה ללא שורת כותרות', severity: 'High', category: 'טבלאות', affectedUsers: ['משתמשי קורא מסך'], impact: 'ללא כותרת.', recommendation: 'הגדירו כותרת.', },
  { id: 'merged-cells', wcagKey: 'info-relationships', title: 'תאים ממוזגים המשמשים לעיצוב', severity: 'High', category: 'טבלאות', affectedUsers: ['משתמשי קורא מסך'], impact: 'מיזוג תאים.', recommendation: 'בטלו מיזוג.', },
  { id: 'empty-cells', wcagKey: 'info-relationships', title: 'תאים ריקים המשמשים לריווח חזותי', severity: 'Medium', category: 'מבנה', affectedUsers: ['משתמשי קורא מסך'], impact: 'תאים ריקים.', recommendation: 'צמצמו ריווח.', },
  { id: 'sheet-without-title', wcagKey: 'headings-labels', title: 'שם גיליון לא תיאורי', severity: 'Medium', category: 'מבנה', affectedUsers: ['משתמשי קורא מסך'], impact: 'Sheet1.', recommendation: 'שם תיאורי.', },
  { id: 'low-contrast', wcagKey: 'contrast-minimum', title: 'ניגודיות צבעים נמוכה', severity: 'High', category: 'תצוגה', affectedUsers: ['משתמשים עם לקות ראייה'], impact: 'ניגודיות.', recommendation: 'שפרו צבעים.', },
]

export const CRITICAL_ANALYSIS_HE =
  'סריקה אוטומטית מסייעת לזהות בעיות נגישות נפוצות, אך אינה מחליפה בדיקה אנושית מקצועית מלאה. בדיקות מסוימות, כגון סדר קריאה וסיווג אובייקטים דקורטיביים, הן היוריסטיות ועשויות להיות שונות מהמנוע הפנימי של Microsoft Office. יש לבדוק ידנית את איכות הטקסט החלופי, הקשר התוכן, וסדר הקריאה לפני פרסום המסמך.'

const LIMITED_SCAN_NOTE =
  'הסריקה האוטומטית זיהתה מעט נתונים מבניים. מומלץ לבצע גם בדיקה ידנית.'

const NO_ISSUES_NOTE =
  'לא נמצאו בעיות נגישות משמעותיות בסריקה האוטומטית. מומלץ לבצע בדיקה ידנית לפני פרסום.'

function isDemoFile(fileName: string): boolean {
  const lower = fileName.toLowerCase()
  return lower.includes('demo') || lower.includes('sample-bad') || lower.includes('bad-accessibility')
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
  const drafts: IssueDraft[] = []
  const meta: string[] = []
  if (signals.slideCount != null) meta.push(`${signals.slideCount} שקופיות`)
  if (signals.paragraphCount != null) meta.push(`${signals.paragraphCount} פסקאות`)
  if (signals.sheetCount != null) meta.push(`${signals.sheetCount} גיליונות`)
  if (signals.fileSize != null) meta.push(`${Math.round(signals.fileSize / 1024)}KB`)

  drafts.push({
    id: 'limited-scan-metadata',
    wcagKey: 'info-relationships',
    title: 'ניתוח מבני מוגבל',
    severity: 'Low',
    category: 'מערכת',
    affectedUsers: ['כל המשתמשים'],
    impact: `${LIMITED_SCAN_NOTE} מטא-נתונים: ${meta.join(', ') || fileName}.`,
    recommendation: 'בצעו בדיקת נגישות ידנית מלאה.',
    location: 'מטא-נתוני קובץ',
  })
  return drafts
}

async function analyzeFile(
  filePath: string,
  fileType: FileType,
  fileSize: number,
): Promise<{ drafts: IssueDraft[]; signals: ExtractedSignals }> {
  switch (fileType) {
    case 'docx':
      return analyzeDocx(filePath, fileSize)
    case 'pptx':
      return analyzePptx(filePath, fileSize)
    case 'xlsx':
      return analyzeXlsx(filePath, fileSize)
  }
}

function logScan(fileName: string, signals: ExtractedSignals, issueCount: number, qfCount: number, score: number) {
  console.log('[AccessiOffice scan]', {
    fileName,
    fileType: signals.fileType,
    fileSize: signals.fileSize,
    slideCount: signals.slideCount,
    paragraphCount: signals.paragraphCount,
    sheetCount: signals.sheetCount,
    imageCount: signals.imageCount,
    tableCount: signals.tableCount,
    hyperlinkCount: signals.hyperlinkCount,
    mergedCellCount: signals.mergedCellCount,
    issueCount,
    quickFixCount: qfCount,
    score,
  })
}

export async function generateScanResult(
  filePath: string,
  fileName: string,
  fileType: FileType,
  userType: UserType,
  scanType: ScanType,
  fileSize: number,
): Promise<ScanResult> {
  let issues: ScanIssue[]
  let signals: ExtractedSignals = { fileType, fileSize }

  if (isDemoFile(fileName)) {
    issues = buildDemoIssues(fileType)
    const score = calculateScore(issues)
    logScan(
      fileName,
      { ...signals, slideCount: fileType === 'pptx' ? 5 : undefined, paragraphCount: fileType === 'docx' ? 20 : undefined, sheetCount: fileType === 'xlsx' ? 2 : undefined },
      issues.length, 0, score,
    )
    return {
      fileName, fileType, userType, scanType,
      score, summary: buildSummary(issues, []), issues,
      criticalAnalysis: CRITICAL_ANALYSIS_HE,
    }
  }

  let quickFix: ScanIssue[] = []

  try {
    const result = await analyzeFile(filePath, fileType, fileSize)
    signals = result.signals
    const allDrafts = dedupeDrafts(result.drafts)

    if (!hasMeaningfulSignals(signals, fileType)) {
      issues = buildFallbackDrafts(signals, fileName).map(buildScanIssue)
    } else {
      issues = allDrafts.filter((d) => !d.isQuickFix).map(buildScanIssue)
      quickFix = allDrafts.filter((d) => d.isQuickFix).map(buildScanIssue)
    }
  } catch (err) {
    console.error('[AccessiOffice scan] error:', (err as Error).message)
    issues = buildFallbackDrafts(signals, fileName).map(buildScanIssue)
  }

  const score = calculateScore(issues)
  logScan(fileName, signals, issues.length, quickFix.length, score)

  let criticalAnalysis = CRITICAL_ANALYSIS_HE
  if (issues.length === 0 && quickFix.length === 0) {
    criticalAnalysis = `${NO_ISSUES_NOTE} ${CRITICAL_ANALYSIS_HE}`
  } else if (issues.some((i) => i.id === 'limited-scan-metadata')) {
    criticalAnalysis = `${LIMITED_SCAN_NOTE} ${CRITICAL_ANALYSIS_HE}`
  }

  return {
    fileName, fileType, userType, scanType,
    score: issues.length === 0 ? 100 : score,
    summary: buildSummary(issues, quickFix),
    issues,
    criticalAnalysis,
    ...(quickFix.length > 0 ? { quickFix } : {}),
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
