import type { OfficeLikeSummary, OfficeLikeSummaryItem, OfficeLikeSummaryStatus, ScanResult } from '../types/scan.types.js'
import type { WorkerSuccess } from '../workers/powerpointWorker.js'
import type { WordWorkerSuccess } from '../workers/wordWorker.js'
import type { ExcelWorkerSuccess } from '../workers/excelWorker.js'

function item(label: string, status: OfficeLikeSummaryStatus, count: number, note?: string): OfficeLikeSummaryItem {
  return { label, status, count, ...(note ? { note } : {}) }
}

/**
 * Build OfficeLikeSummary from the PowerPoint UI-automation worker result.
 *
 * Worker → OfficeLikeSummary key mapping:
 *   counts.missingAltText      → missingAltText
 *   counts.missingSlideTitle   → missingSlideTitle
 *   counts.duplicateSlideTitle → duplicateSlideTitle
 *   counts.readingOrder        → readingOrder  (always 'manual')
 *   counts.hardToReadText      → contrast      (PPT calls this "Hard-to-Read Text")
 *   counts.missingTableHeader  → tableHeader
 *   counts.mediaCaptions       → mediaCaptions
 *   counts.mergedCells         → mergedCells
 *
 *   statuses.contrast           → contrast status ('passed'/'failed')
 *   statuses.mediaCaptions      → mediaCaptions status
 *   statuses.missingTableHeader → tableHeader status
 *   statuses.mergedCells        → mergedCells status
 *
 * restrictedAccess is always 'manual' — PPT reports it through a different path.
 * xmlSummary is used as fallback only for keys the worker cannot determine.
 */
function buildWorkerOfficeLikeSummary(
  counts: Record<string, number>,
  statuses: Record<string, string>,
  xmlSummary: OfficeLikeSummary | undefined,
): OfficeLikeSummary {
  const cnt = (key: string) => counts[key] ?? 0
  const st  = (key: string, fallback: OfficeLikeSummaryStatus): OfficeLikeSummaryStatus => {
    const v = statuses[key]
    if (v === 'failed' || v === 'passed' || v === 'manual' || v === 'partial' || v === 'not_checked') return v
    return fallback
  }

  const contrastCount = cnt('hardToReadText')
  const contrastStatus = st('contrast', contrastCount > 0 ? 'failed' : 'passed')

  const mediaCount  = cnt('mediaCaptions')
  const mediaStatus = st('mediaCaptions', mediaCount > 0 ? 'failed' : 'passed')

  const tableCount  = cnt('missingTableHeader')
  const tableStatus = st('missingTableHeader', tableCount > 0 ? 'failed' : 'passed')

  const mergedCount  = cnt('mergedCells')
  const mergedStatus = st('mergedCells', mergedCount > 0 ? 'failed' : 'passed')

  const missingAltCount  = cnt('missingAltText')
  const missingSlideCnt  = cnt('missingSlideTitle')
  const dupTitleCnt      = cnt('duplicateSlideTitle')
  const readingOrderCnt  = cnt('readingOrder')

  return {
    contrast: item(
      'ניגודיות טקסט קשה לקריאה',
      contrastStatus,
      contrastCount,
      contrastStatus === 'passed'
        ? 'נבדק על ידי Microsoft PowerPoint Accessibility Checker.'
        : undefined,
    ),

    missingAltText: item(
      'חסר טקסט חלופי',
      missingAltCount > 0 ? 'failed' : 'passed',
      missingAltCount,
    ),

    mediaCaptions: item(
      'חסרות כתוביות לאודיו/וידאו',
      mediaStatus,
      mediaCount,
      mediaStatus === 'passed'
        ? 'נבדק על ידי Microsoft PowerPoint Accessibility Checker.'
        : undefined,
    ),

    tableHeader: item(
      'חסרת שורת כותרת בטבלה',
      tableStatus,
      tableCount,
      tableStatus === 'passed'
        ? 'נבדק על ידי Microsoft PowerPoint Accessibility Checker.'
        : undefined,
    ),

    mergedCells: item(
      'שימוש בתאים ממוזגים או מפוצלים',
      mergedStatus,
      mergedCount,
      mergedStatus === 'passed'
        ? 'נבדק על ידי Microsoft PowerPoint Accessibility Checker.'
        : undefined,
    ),

    missingSlideTitle:   item('שקופית ללא כותרת',    missingSlideCnt > 0 ? 'failed' : 'passed', missingSlideCnt),
    duplicateSlideTitle: item('כותרת שקופית כפולה',  dupTitleCnt > 0 ? 'failed' : 'passed',    dupTitleCnt),

    readingOrder: item(
      'בדיקת סדר קריאה',
      'manual',
      readingOrderCnt,
      'בדיקה היוריסטית; אמתו את סדר הקריאה בחלונית Reading Order ב-PowerPoint לפני פרסום.',
    ),

    restrictedAccess: xmlSummary?.restrictedAccess ??
      item('גישה מוגבלת למסמך', 'manual', 0, 'הרשאות מסמך אינן ניתנות לסריקה אוטומטית.'),
  }
}

// ── Public API ────────────────────────────────────────────────────────────────

/**
 * Merge a successful PowerPoint UI-automation worker result into the XML ScanResult.
 * - officeLikeSummary counts/statuses come from the real PPT Accessibility Checker
 * - issues[], quickFix[], score, scanDiagnostics remain from the XML scanner
 * - rawOfficeText is not propagated into the API response
 */
export function mergeWorkerIntoResult(
  xmlResult: ScanResult,
  workerResult: WorkerSuccess,
): ScanResult {
  const officeLikeSummary = buildWorkerOfficeLikeSummary(
    workerResult.counts,
    workerResult.statuses,
    xmlResult.officeLikeSummary,
  )
  return {
    ...xmlResult,
    engine:               'office-engine',
    workerScannerVersion: workerResult.pptVersion,
    officeLikeSummary,
  }
}

/**
 * Mark the result as hybrid (PowerPoint worker failed, XML fallback used).
 * Preserves the XML-derived officeLikeSummary as-is.
 */
export function markWorkerFallback(xmlResult: ScanResult, errorMessage: string): ScanResult {
  return {
    ...xmlResult,
    engine:      'hybrid',
    workerError: `PowerPoint worker unavailable — XML fallback used. Reason: ${errorMessage}`,
  }
}

// ── Word UI-Automation worker ─────────────────────────────────────────────────

/**
 * Build OfficeLikeSummary from the Word UI-automation worker result.
 *
 * The Word worker uses final key names directly in counts (contrast, not hardToReadText),
 * so no key translation is needed here.
 *
 * Word OfficeLikeSummary keys:
 *   contrast            — Hard-to-read text contrast
 *   missingAltText      — Missing alternative text on images/objects
 *   tableHeader         — Missing table header row
 *   mergedCells         — Merged or split cells
 *   unclearHyperlinkText — Unclear hyperlink text (e.g. "click here")
 *   documentTitle       — Missing document title in file properties
 *   restrictedAccess    — Document access restrictions (always manual)
 */
function buildWordWorkerOfficeLikeSummary(
  counts: Record<string, number>,
  statuses: Record<string, string>,
): OfficeLikeSummary {
  const cnt = (key: string) => counts[key] ?? 0
  const st  = (key: string, fallback: OfficeLikeSummaryStatus): OfficeLikeSummaryStatus => {
    const v = statuses[key]
    if (v === 'failed' || v === 'passed' || v === 'manual' || v === 'partial' || v === 'not_checked') return v
    return fallback
  }

  const contrastCount  = cnt('contrast')
  const contrastStatus = st('contrast', contrastCount > 0 ? 'failed' : 'passed')

  const altCount  = cnt('missingAltText')
  const altStatus = st('missingAltText', altCount > 0 ? 'failed' : 'passed')

  const tableCount  = cnt('tableHeader')
  const tableStatus = st('tableHeader', tableCount > 0 ? 'failed' : 'passed')

  const mergedCount  = cnt('mergedCells')
  const mergedStatus = st('mergedCells', mergedCount > 0 ? 'failed' : 'passed')

  const hyperlinkCount  = cnt('unclearHyperlinkText')
  const hyperlinkStatus = st('unclearHyperlinkText', hyperlinkCount > 0 ? 'failed' : 'passed')

  const docTitleCount  = cnt('documentTitle')
  const docTitleStatus = st('documentTitle', docTitleCount > 0 ? 'failed' : 'passed')

  return {
    contrast: item(
      'ניגודיות טקסט קשה לקריאה',
      contrastStatus,
      contrastCount,
      contrastStatus === 'passed'
        ? 'נבדק על ידי Microsoft Word Accessibility Checker.'
        : undefined,
    ),

    missingAltText: item(
      'חסר טקסט חלופי',
      altStatus,
      altCount,
      altStatus === 'passed'
        ? 'נבדק על ידי Microsoft Word Accessibility Checker.'
        : undefined,
    ),

    tableHeader: item(
      'חסרת שורת כותרת בטבלה',
      tableStatus,
      tableCount,
      tableStatus === 'passed'
        ? 'נבדק על ידי Microsoft Word Accessibility Checker.'
        : undefined,
    ),

    mergedCells: item(
      'שימוש בתאים ממוזגים או מפוצלים',
      mergedStatus,
      mergedCount,
      mergedStatus === 'passed'
        ? 'נבדק על ידי Microsoft Word Accessibility Checker.'
        : undefined,
    ),

    unclearHyperlinkText: item(
      'טקסט קישור לא ברור',
      hyperlinkStatus,
      hyperlinkCount,
    ),

    documentTitle: item(
      'כותרת מסמך חסרה',
      docTitleStatus,
      docTitleCount,
    ),

    restrictedAccess: item(
      'גישה מוגבלת למסמך',
      'manual',
      0,
      'הרשאות מסמך אינן ניתנות לסריקה אוטומטית.',
    ),
  }
}

/**
 * Merge a successful Word UI-automation worker result into the XML ScanResult.
 * - officeLikeSummary counts/statuses come from the real Word Accessibility Checker
 * - issues[], quickFix[], score, scanDiagnostics remain from the XML scanner
 */
export function mergeWordWorkerIntoResult(
  xmlResult: ScanResult,
  workerResult: WordWorkerSuccess,
): ScanResult {
  const officeLikeSummary = buildWordWorkerOfficeLikeSummary(
    workerResult.counts,
    workerResult.statuses,
  )
  return {
    ...xmlResult,
    engine:               'office-engine',
    workerScannerVersion: workerResult.wordVersion,
    officeLikeSummary,
  }
}

/**
 * Mark the result as hybrid (Word worker failed, XML fallback used).
 */
export function markWordWorkerFallback(xmlResult: ScanResult, errorMessage: string): ScanResult {
  return {
    ...xmlResult,
    engine:      'hybrid',
    workerError: `Word worker unavailable — XML fallback used. Reason: ${errorMessage}`,
  }
}

// ── Excel UI-Automation worker ────────────────────────────────────────────────

/**
 * Build OfficeLikeSummary from the Excel UI-automation worker result.
 *
 * The Excel worker uses final key names directly in counts, so no key
 * translation is needed here.
 *
 * Excel OfficeLikeSummary keys:
 *   contrast            — Hard-to-read text contrast
 *   missingAltText      — Missing alternative text on images/charts/objects
 *   tableHeader         — Missing table header row
 *   mergedCells         — Merged or split cells
 *   unclearHyperlinkText — Unclear hyperlink text (e.g. "click here")
 *   sheetName           — Default or inaccessible sheet tab name (Excel-specific)
 *   restrictedAccess    — Document access restrictions (always manual)
 */
function buildExcelWorkerOfficeLikeSummary(
  counts: Record<string, number>,
  statuses: Record<string, string>,
): OfficeLikeSummary {
  const cnt = (key: string) => counts[key] ?? 0
  const st  = (key: string, fallback: OfficeLikeSummaryStatus): OfficeLikeSummaryStatus => {
    const v = statuses[key]
    if (v === 'failed' || v === 'passed' || v === 'manual' || v === 'partial' || v === 'not_checked') return v
    return fallback
  }

  const contrastCount  = cnt('contrast')
  const contrastStatus = st('contrast', contrastCount > 0 ? 'failed' : 'passed')

  const altCount  = cnt('missingAltText')
  const altStatus = st('missingAltText', altCount > 0 ? 'failed' : 'passed')

  const tableCount  = cnt('tableHeader')
  const tableStatus = st('tableHeader', tableCount > 0 ? 'failed' : 'passed')

  const mergedCount  = cnt('mergedCells')
  const mergedStatus = st('mergedCells', mergedCount > 0 ? 'failed' : 'passed')

  const hyperlinkCount  = cnt('unclearHyperlinkText')
  const hyperlinkStatus = st('unclearHyperlinkText', hyperlinkCount > 0 ? 'failed' : 'passed')

  const sheetNameCount  = cnt('sheetName')
  const sheetNameStatus = st('sheetName', sheetNameCount > 0 ? 'failed' : 'passed')

  return {
    contrast: item(
      'ניגודיות טקסט קשה לקריאה',
      contrastStatus,
      contrastCount,
      contrastStatus === 'passed'
        ? 'נבדק על ידי Microsoft Excel Accessibility Checker.'
        : undefined,
    ),

    missingAltText: item(
      'חסר טקסט חלופי',
      altStatus,
      altCount,
      altStatus === 'passed'
        ? 'נבדק על ידי Microsoft Excel Accessibility Checker.'
        : undefined,
    ),

    tableHeader: item(
      'חסרת שורת כותרת בטבלה',
      tableStatus,
      tableCount,
      tableStatus === 'passed'
        ? 'נבדק על ידי Microsoft Excel Accessibility Checker.'
        : undefined,
    ),

    mergedCells: item(
      'שימוש בתאים ממוזגים או מפוצלים',
      mergedStatus,
      mergedCount,
      mergedStatus === 'passed'
        ? 'נבדק על ידי Microsoft Excel Accessibility Checker.'
        : undefined,
    ),

    unclearHyperlinkText: item(
      'טקסט קישור לא ברור',
      hyperlinkStatus,
      hyperlinkCount,
    ),

    sheetName: item(
      'שם לשונית גיליון לא תיאורי',
      sheetNameStatus,
      sheetNameCount,
    ),

    restrictedAccess: item(
      'גישה מוגבלת למסמך',
      'manual',
      0,
      'הרשאות מסמך אינן ניתנות לסריקה אוטומטית.',
    ),
  }
}

/**
 * Merge a successful Excel UI-automation worker result into the XML ScanResult.
 * - officeLikeSummary counts/statuses come from the real Excel Accessibility Checker
 * - issues[], quickFix[], score, scanDiagnostics remain from the XML scanner
 */
export function mergeExcelWorkerIntoResult(
  xmlResult: ScanResult,
  workerResult: ExcelWorkerSuccess,
): ScanResult {
  const officeLikeSummary = buildExcelWorkerOfficeLikeSummary(
    workerResult.counts,
    workerResult.statuses,
  )
  return {
    ...xmlResult,
    engine:               'office-engine',
    workerScannerVersion: workerResult.excelVersion,
    officeLikeSummary,
  }
}

/**
 * Mark the result as hybrid (Excel worker failed, XML fallback used).
 */
export function markExcelWorkerFallback(xmlResult: ScanResult, errorMessage: string): ScanResult {
  return {
    ...xmlResult,
    engine:      'hybrid',
    workerError: `Excel worker unavailable — XML fallback used. Reason: ${errorMessage}`,
  }
}
