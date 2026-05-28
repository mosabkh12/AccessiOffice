/**
 * Shared helpers for Office Engine results.
 * Used by ResultsPage, ReportPage, and any future export/print view.
 * All calculations are derived at runtime from officeLikeSummary — never hardcoded.
 */

// Severity classification used in buildOfficeEngineSummary counting
// PPTX keys: missingAltText, contrast, mediaCaptions, missingSlideTitle, duplicateSlideTitle, tableHeader, mergedCells
// DOCX keys: missingAltText, contrast, tableHeader, mergedCells, unclearHyperlinkText, documentTitle
// XLSX keys: missingAltText, contrast, tableHeader, mergedCells, unclearHyperlinkText, sheetName
const HIGH_KEYS   = ['missingAltText', 'contrast', 'mediaCaptions']
const MEDIUM_KEYS = [
  'missingSlideTitle', 'duplicateSlideTitle',
  'tableHeader', 'mergedCells',
  'unclearHyperlinkText', 'documentTitle',
  'sheetName',
]
// readingOrder and restrictedAccess are always status='manual' in the backend

export const OFFICE_KEY_SEVERITY = {
  missingAltText:       'high',
  contrast:             'high',
  mediaCaptions:        'high',
  tableHeader:          'medium',
  mergedCells:          'medium',
  missingSlideTitle:    'medium',
  duplicateSlideTitle:  'medium',
  unclearHyperlinkText: 'medium',
  documentTitle:        'medium',
  sheetName:            'medium',
  readingOrder:         'manual',
  restrictedAccess:     'manual',
}

export const OFFICE_KEY_RECOMMENDATION = {
  // PPTX + shared
  missingAltText:       'הוסיפו טקסט חלופי לכל תמונה ואובייקט גרפי.',
  contrast:             'שפרו ניגודיות הטקסט כך שתעמוד ביחס 4.5:1 לפחות (WCAG AA).',
  mediaCaptions:        'הוסיפו כתוביות לכל קטעי האודיו והווידאו בקובץ.',
  tableHeader:          'הגדירו שורת כותרת לכל טבלה.',
  mergedCells:          'הימנעו ממיזוג תאים — השתמשו בפריסה חלופית.',
  missingSlideTitle:    'הוסיפו כותרת לכל שקופית כדי לאפשר ניווט למשתמשי קוראי מסך.',
  duplicateSlideTitle:  'ודאו שלכל שקופית יש כותרת ייחודית.',
  readingOrder:         'בדקו את סדר הקריאה בחלונית Reading Order ב-PowerPoint לפני פרסום.',
  restrictedAccess:     'בדקו שהמסמך אינו מוגן בסיסמה ואינו מגביל קוראי מסך.',
  // DOCX-specific
  unclearHyperlinkText: 'הוסיפו טקסט ברור לקישורים — הימנעו מ"לחצו כאן" וביטויים דומים.',
  documentTitle:        'הוסיפו כותרת מסמך דרך File → Info → Properties → Title.',
  // XLSX-specific
  sheetName:            'שנו את שמות לשוניות הגיליון לשמות תיאוריים במקום ברירות המחדל כגון "Sheet1".',
}

export const OFFICE_REPORT_GROUPS = [
  { label: 'צבע וניגודיות',  keys: ['contrast'] },
  { label: 'מדיה ואיורים',   keys: ['missingAltText', 'mediaCaptions'] },
  { label: 'טבלאות',          keys: ['tableHeader', 'mergedCells'] },
  // PPTX structure keys + DOCX documentTitle + XLSX sheetName
  { label: 'מבנה מסמך',      keys: ['missingSlideTitle', 'duplicateSlideTitle', 'readingOrder', 'documentTitle', 'sheetName'] },
  // DOCX + XLSX hyperlink key
  { label: 'קישורים וניווט', keys: ['unclearHyperlinkText'] },
  { label: 'גישה למסמך',     keys: ['restrictedAccess'] },
]

/**
 * Returns true when a scan result should use Office Engine data as the primary source.
 */
export function isOfficeEngineResult(result) {
  return result?.engine === 'office-engine' && !!result?.officeLikeSummary
}

/**
 * Compute display-ready aggregate numbers from officeLikeSummary.
 * - totalFindings: sum of all positive item.count values
 * - issueTypes:    number of items with count > 0
 * - highCount:     sum of counts for HIGH_KEYS items that failed
 * - mediumCount:   sum of counts for MEDIUM_KEYS items that failed
 * - manualCount:   number of items with manual/partial/not_checked status
 */
export function buildOfficeEngineSummary(officeLikeSummary) {
  if (!officeLikeSummary) return null

  let totalFindings = 0
  let issueTypes    = 0
  let highCount     = 0
  let mediumCount   = 0
  let manualCount   = 0

  Object.entries(officeLikeSummary).forEach(([key, item]) => {
    const count  = item?.count  ?? 0
    const status = item?.status

    if (count > 0) { totalFindings += count; issueTypes++ }

    if (['manual', 'partial', 'not_checked'].includes(status)) {
      manualCount++
    } else if (HIGH_KEYS.includes(key) && status === 'failed' && count > 0) {
      highCount += count
    } else if (MEDIUM_KEYS.includes(key) && status === 'failed' && count > 0) {
      mediumCount += count
    }
  })

  return { totalFindings, issueTypes, highCount, mediumCount, manualCount }
}

/**
 * One shared score formula for all views (ResultsPage, ReportPage, PDF).
 * score = 100 - highCount*8 - mediumCount*5 - manualCount*3, clamped 0-100.
 */
export function calculateOfficeEngineScore(officeLikeSummary) {
  if (!officeLikeSummary) return 0
  const { highCount, mediumCount, manualCount } = buildOfficeEngineSummary(officeLikeSummary)
  return Math.max(0, Math.min(100,
    100 - highCount * 8 - mediumCount * 5 - manualCount * 3
  ))
}

/**
 * Build a flat array of grouped issue rows from officeLikeSummary.
 * One row per check — no repeated XML rows.
 * Items that are fully passed with count 0 are excluded.
 * Order follows OFFICE_REPORT_GROUPS.
 */
export function buildOfficeEngineIssueRows(officeLikeSummary) {
  if (!officeLikeSummary) return []

  const rows = []
  for (const group of OFFICE_REPORT_GROUPS) {
    for (const key of group.keys) {
      const item = officeLikeSummary[key]
      if (!item) continue
      const count    = item.count ?? 0
      const status   = item.status
      const isManual = ['manual', 'partial', 'not_checked'].includes(status)
      if (!isManual && status === 'passed' && count === 0) continue
      rows.push({
        key,
        label:          item.label,
        count,
        status,
        severity:       isManual ? 'manual' : (OFFICE_KEY_SEVERITY[key] ?? 'medium'),
        category:       group.label,
        recommendation: OFFICE_KEY_RECOMMENDATION[key] ?? '',
        note:           item.note ?? null,
      })
    }
  }
  return rows
}
