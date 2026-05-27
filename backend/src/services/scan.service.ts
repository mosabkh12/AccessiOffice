import type { WcagRuleKey } from '../data/wcag.rules.js'
import { getWcagRule } from '../data/wcag.rules.js'
import type {
  FileType,
  OfficeLikeSummary,
  OfficeLikeSummaryItem,
  OfficeLikeSummaryStatus,
  ScanCheckStatus,
  ScanIssue,
  ScanResult,
  ScanType,
  Severity,
  UserType,
} from '../types/scan.types.js'
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

const DEBUG_SCAN = process.env.ACCESSIOFFICE_DEBUG_SCAN === 'true'
const PPTX_DIAGNOSTICS_VERSION = 'pptx-diagnostics-v2'

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
  if (!DEBUG_SCAN) return
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

function buildScanDiagnostics(signals: ExtractedSignals): Record<string, unknown> {
  if (signals.fileType !== 'pptx') return {}
  return {
    contrastCheckStatus: signals.contrastCheckStatus ?? 'not_checked',
    contrastCheckedRuns: signals.contrastCheckedRuns ?? 0,
    contrastSkippedRuns: signals.contrastSkippedRuns ?? 0,
    pptxTableCount: signals.pptxTableCount ?? 0,
    pptxTablesMissingHeaderCount: signals.pptxTablesMissingHeaderCount ?? 0,
    pptxMergedCellCount: signals.pptxMergedCellCount ?? 0,
    pptxMediaCount: signals.pptxMediaCount ?? 0,
    pptxMediaMissingCaptionsCount: signals.pptxMediaMissingCaptionsCount ?? 0,
    officeLikeDiagnostics: {
      // ── Missing alt text ──────────────────────────────────────────────────────
      missingAltTextRawCount: signals.missingAltCount ?? 0,
      missingAltTextOfficeLikeCount: signals.officeLikeMissingAltCount ?? 0,
      // Alt-text exclusion breakdown
      excludedDecorative: signals.decorativeCandidateCount ?? 0,
      excludedMasterOrLayout: signals.officeLikeAltExcludedLayout ?? 0,
      excludedFooterOrLogo: 0,  // included in excludedMasterOrLayout (tiny + non-picture layout shapes)
      excludedRepeatedTemplateStrongEvidence: signals.officeLikeAltExcludedLayout ?? 0,
      includedActionableImages: signals.officeLikeMissingAltCount ?? 0,
      // legacy field kept for backward-compat
      excludedRepeatedTemplateCount: signals.officeLikeAltExcludedLayout ?? 0,
      excludedDecorativeCount: signals.decorativeCandidateCount ?? 0,
      // ── Missing slide title ───────────────────────────────────────────────────
      missingSlideTitleRawCount: signals.missingTitleSlides?.length ?? 0,
      missingSlideTitleOfficeLikeCount: signals.officeLikeMissingTitleSlides?.length ?? 0,
      slidesWithRealTitlePlaceholder: signals.officeLikeTitleWithRealPlaceholderCount ?? 0,
      slidesMissingOfficeLikeTitle: signals.officeLikeMissingTitleSlides?.length ?? 0,
      // ── Reading order ─────────────────────────────────────────────────────────
      readingOrderRawCount: signals.readingOrderRiskSlides?.length ?? 0,
      readingOrderOfficeLikeCount: signals.officeLikeReadingOrderCount ?? 0,
    },
  }
}

function isQuickFixDraft(draft: IssueDraft): boolean {
  return draft.isQuickFix === true || draft.issueTier === 'quickFix'
}

function issueOccurrenceCount(issues: ScanIssue[], predicate: (issue: ScanIssue) => boolean): number {
  return issues.filter(predicate).reduce((sum, issue) => sum + (issue.occurrenceCount ?? 1), 0)
}

function buildPptxCheckStatuses(signals: ExtractedSignals, issues: ScanIssue[]): ScanCheckStatus[] {
  if (signals.fileType !== 'pptx') return []

  const contrastCount = issueOccurrenceCount(issues, (i) => i.wcagKey === 'contrast-minimum' || /contrast|ניגודיות/i.test(`${i.id} ${i.title}`))
  const mediaCaptionCount = issueOccurrenceCount(issues, (i) => i.wcagKey === 'captions-prerecorded' || /media-captions|כתוביות/i.test(`${i.id} ${i.title}`))
  const tableHeaderCount = issueOccurrenceCount(issues, (i) => /table-header/.test(i.id) || /שורת כותרת|כותרת בטבלה/.test(i.title))
  const mergedCellsCount = issueOccurrenceCount(issues, (i) => /merged-cells/.test(i.id) || /ממוזגים|מפוצלים/.test(i.title))
  const readingOrderCount = issueOccurrenceCount(issues, (i) => /reading-order/.test(i.id) || i.wcagKey === 'meaningful-sequence')

  const contrastStatus =
    contrastCount > 0
      ? 'failed'
      : signals.contrastCheckStatus === 'checked'
        ? 'passed'
        : signals.contrastCheckStatus === 'partial'
          ? 'partial'
          : 'not_checked'

  const statuses: ScanCheckStatus[] = [
    {
      id: 'contrast',
      label: 'ניגודיות טקסט קשה לקריאה',
      status: contrastStatus,
      count: contrastCount,
      note:
        contrastStatus === 'partial'
          ? 'נסרקו רק צבעי sRGB מפורשים; צבעי ערכת נושא, ירושה ורקע לא מפורש דורשים בדיקה ידנית.'
          : contrastStatus === 'not_checked'
            ? 'לא נמצאו מספיק צבעי טקסט מפורשים לסריקת ניגודיות אמינה.'
            : undefined,
    },
    {
      id: 'media-captions',
      label: 'חסרות כתוביות לאודיו/וידאו',
      status: mediaCaptionCount > 0 ? 'failed' : 'passed',
      count: mediaCaptionCount,
      note: mediaCaptionCount === 0 ? `נבדקו קשרי מדיה נפוצים; זוהו ${signals.pptxMediaCount ?? 0} אובייקטי מדיה.` : undefined,
    },
    {
      id: 'table-header',
      label: 'חסרת שורת כותרת בטבלה',
      status: tableHeaderCount > 0 ? 'failed' : 'passed',
      count: tableHeaderCount,
      note: tableHeaderCount === 0 ? `נבדקו ${signals.pptxTableCount ?? 0} טבלאות DrawingML.` : undefined,
    },
    {
      id: 'merged-cells',
      label: 'שימוש בתאים ממוזגים או מפוצלים',
      status: mergedCellsCount > 0 ? 'failed' : 'passed',
      count: mergedCellsCount,
      note: mergedCellsCount === 0 ? `נבדקו ${signals.pptxTableCount ?? 0} טבלאות DrawingML.` : undefined,
    },
    {
      id: 'reading-order',
      label: 'בדיקת סדר קריאה',
      status: readingOrderCount > 0 ? 'failed' : 'manual',
      count: readingOrderCount,
      note: 'בדיקת סדר הקריאה היא היוריסטית; יש לאמת ידנית בחלונית Reading Order של PowerPoint.',
    },
    {
      id: 'restricted-access',
      label: 'גישה מוגבלת למסמך',
      status: 'manual',
      count: 0,
      note: 'בדיקת הרשאות והגבלות מסמך אינה מיושמת בסורק ה-PPTX.',
    },
  ]

  return statuses
}

function item(
  label: string,
  status: OfficeLikeSummaryStatus,
  count: number,
  note?: string,
): OfficeLikeSummaryItem {
  return { label, status, count, ...(note ? { note } : {}) }
}

function buildOfficeLikeSummary(signals: ExtractedSignals): OfficeLikeSummary {
  // Contrast: only flag as failed when explicit fg+bg pair was used (not white-fallback bg).
  // Falls back to passed when all checked runs used the white-fallback background.
  const contrastExplicitCount = signals.contrastIssuesWithExplicitBg ?? 0
  const contrastStatus: OfficeLikeSummaryStatus = contrastExplicitCount > 0 ? 'failed' : 'passed'

  // Missing alt text: office-like count excludes layout-inherited images (template logos,
  // branding graphics). Falls back to raw missingAltCount if the refined signal is absent.
  const missingAltCount = signals.officeLikeMissingAltCount ?? signals.missingAltCount ?? 0
  const missingAltStatus: OfficeLikeSummaryStatus = missingAltCount > 0 ? 'failed' : 'passed'

  // Media captions: if no media objects detected, passed.
  const mediaMissingCount = signals.pptxMediaMissingCaptionsCount ?? 0
  const mediaStatus: OfficeLikeSummaryStatus = mediaMissingCount > 0 ? 'failed' : 'passed'

  // Table checks: based on DrawingML table analysis.
  const tableHeaderMissing = signals.pptxTablesMissingHeaderCount ?? 0
  const tableHeaderStatus: OfficeLikeSummaryStatus = tableHeaderMissing > 0 ? 'failed' : 'passed'
  const mergedCellCount = signals.pptxMergedCellCount ?? 0
  const mergedCellStatus: OfficeLikeSummaryStatus = mergedCellCount > 0 ? 'failed' : 'passed'
  const tableNote = `נבדקו ${signals.pptxTableCount ?? 0} טבלאות DrawingML.`

  // Slide title: office-like count also covers slides whose title came from cSld-name only.
  // PowerPoint requires a title placeholder; cSld-name is not a recognized title source.
  const missingTitleCount =
    signals.officeLikeMissingTitleSlides?.length ?? signals.missingTitleSlides?.length ?? 0
  const missingTitleStatus: OfficeLikeSummaryStatus = missingTitleCount > 0 ? 'failed' : 'passed'
  const dupTitleCount = signals.duplicateTitleSlides?.length ?? 0
  const dupTitleStatus: OfficeLikeSummaryStatus = dupTitleCount > 0 ? 'failed' : 'passed'

  // Reading order: office-like uses an inclusive object count (>= 3 visible objects per slide,
  // including placeholders) rather than the strict accessibility-score threshold used for raw issues.
  // Always manual status — this is a review warning, not a guaranteed content failure.
  const readingOrderCount =
    signals.officeLikeReadingOrderCount ?? signals.readingOrderRiskSlides?.length ?? 0

  return {
    contrast: item(
      'ניגודיות טקסט קשה לקריאה',
      contrastStatus,
      contrastExplicitCount,
      contrastStatus === 'passed'
        ? 'נבדקו רק זוגות של צבע קדמי/רקע מפורשים; צבעי ערכת נושא ורקע לא מפורש דורשים בדיקה ידנית.'
        : undefined,
    ),
    missingAltText: item('חסר טקסט חלופי', missingAltStatus, missingAltCount),
    mediaCaptions: item(
      'חסרות כתוביות לאודיו/וידאו',
      mediaStatus,
      mediaMissingCount,
      mediaStatus === 'passed'
        ? `זוהו ${signals.pptxMediaCount ?? 0} אובייקטי מדיה; לא זוהתה מדיה ללא כתוביות.`
        : undefined,
    ),
    tableHeader: item(
      'חסרת שורת כותרת בטבלה',
      tableHeaderStatus,
      tableHeaderMissing,
      tableHeaderStatus === 'passed' ? tableNote : undefined,
    ),
    mergedCells: item(
      'שימוש בתאים ממוזגים או מפוצלים',
      mergedCellStatus,
      mergedCellCount,
      mergedCellStatus === 'passed' ? tableNote : undefined,
    ),
    missingSlideTitle: item('שקופית ללא כותרת', missingTitleStatus, missingTitleCount),
    duplicateSlideTitle: item('כותרת שקופית כפולה', dupTitleStatus, dupTitleCount),
    readingOrder: item(
      'בדיקת סדר קריאה',
      'manual',
      readingOrderCount,
      'בדיקה היוריסטית; אמתו את סדר הקריאה בחלונית Reading Order ב-PowerPoint לפני פרסום.',
    ),
    restrictedAccess: item(
      'גישה מוגבלת למסמך',
      'manual',
      0,
      'הרשאות מסמך אינן ניתנות לסריקה אוטומטית.',
    ),
  }
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
      issues = allDrafts.filter((d) => !isQuickFixDraft(d)).map(buildScanIssue)
      quickFix = allDrafts.filter(isQuickFixDraft).map(buildScanIssue)
    }
  } catch (err) {
    console.error('[AccessiOffice scan] error:', (err as Error).message)
    issues = buildFallbackDrafts(signals, fileName).map(buildScanIssue)
  }

  const score = calculateScore(issues)
  logScan(fileName, signals, issues.length, quickFix.length, score)
  const checkStatuses = buildPptxCheckStatuses(signals, issues)
  const manualReviewChecks = checkStatuses.filter((c) => c.status === 'manual' || c.status === 'partial' || c.status === 'not_checked')
  const scanDiagnostics = buildScanDiagnostics(signals)
  const officeLikeSummary = fileType === 'pptx' ? buildOfficeLikeSummary(signals) : undefined

  if (DEBUG_SCAN && fileType === 'pptx') {
    console.log('[AccessiOffice PPTX diagnostics response]', {
      scannerVersion: PPTX_DIAGNOSTICS_VERSION,
      issueCount: issues.reduce((sum, issue) => sum + (issue.occurrenceCount ?? 1), 0),
      contrastCheckStatus: scanDiagnostics.contrastCheckStatus,
      tableStatus: checkStatuses.filter((c) => c.id === 'table-header' || c.id === 'merged-cells'),
      mediaStatus: checkStatuses.find((c) => c.id === 'media-captions'),
      checkStatuses,
    })
  }

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
    ...(fileType === 'pptx' ? { scannerVersion: PPTX_DIAGNOSTICS_VERSION } : {}),
    ...(fileType === 'pptx' && DEBUG_SCAN ? { debugMarker: 'API-SCAN-RESPONSE-V2-ACTIVE' } : {}),
    ...(fileType === 'pptx' ? { diagnostics: scanDiagnostics } : {}),
    scanDiagnostics,
    ...(checkStatuses.length > 0 ? { checkStatuses } : {}),
    ...(manualReviewChecks.length > 0 ? { manualReviewChecks } : {}),
    ...(officeLikeSummary ? { officeLikeSummary } : {}),
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
