/**
 * Shared helpers for Office Engine results.
 * Used by ResultsPage, ReportPage, and any future export/print view.
 * All calculations are derived at runtime from officeLikeSummary — never hardcoded.
 */

// ── Occurrence location formatting ────────────────────────────────────────────
// Maps English PowerPoint/Word/Excel shape-type names to Hebrew display labels.
const SHAPE_EN_TO_HE = {
  'Picture':              'תמונה',
  'Freeform':             'אובייקט ציור',
  'Group':                'קבוצת אובייקטים',
  'Rectangle':            'צורה',
  'RoundedRectangle':     'צורה מעוגלת',
  'Oval':                 'עיגול',
  'Triangle':             'משולש',
  'TextBox':              'תיבת טקסט',
  'Table':                'טבלה',
  'Chart':                'תרשים',
  'SmartArt Graphic':     'SmartArt',
  'SmartArt':             'SmartArt',
  'Shape':                'צורה',
  'Line':                 'קו',
  'Connector':            'מחבר',
  'Diamond':              'יהלום',
  'Placeholder':          'מסגרת תוכן',
  'MediaObject':          'אובייקט מדיה',
  'Hyperlink':            'קישור',
  'Picture Placeholder':  'מסגרת תמונה',
}

// Extra hint shown inside an expanded occurrence for shape types that need explanation.
const SHAPE_TYPE_HINTS = {
  freeform: 'אובייקט ציור הוא צורה או רכיב גרפי שנוצר בתוך PowerPoint.',
  group:    'קבוצת אובייקטים היא כמה צורות/תמונות שמקובצות יחד.',
}

/**
 * Translate an English shape/object name from COM into a Hebrew display string.
 * "Freeform 4" → "אובייקט ציור 4",  "Picture 3" → "תמונה 3"
 */
function translateShapeName(name) {
  if (!name) return name
  // Strip technical row/col suffixes like "(Row 2 Col 3)"
  const cleaned = name.replace(/\s*\(Row\s+\d+\s+Col\s+\d+\).*$/, '').trim()
  // Try longest-match first so "SmartArt Graphic" beats "SmartArt"
  const keys = Object.keys(SHAPE_EN_TO_HE).sort((a, b) => b.length - a.length)
  for (const eng of keys) {
    const re = new RegExp(`^${eng.replace(/\s+/, '\\s+')}(\\s+\\d+)?$`, 'i')
    const m  = cleaned.match(re)
    if (m) return `${SHAPE_EN_TO_HE[eng]}${m[1] ?? ''}`
  }
  return cleaned
}

/**
 * Returns true when a location or name string contains garbled/corrupted text
 * (question-mark runs or Unicode replacement characters).
 * Used to decide whether to fall back to a sheet index.
 */
export function isCorruptedText(value) {
  if (!value || typeof value !== 'string') return true
  if (/\?{2,}/.test(value)) return true   // ????? pattern
  if (/�/.test(value)) return true   // Unicode replacement char (�)
  return false
}

/**
 * Convert a raw ASCII location string from the Office worker into a
 * user-friendly Hebrew display string.
 *
 * Input formats (from PS scripts):
 *   "Slide N - ShapeName"     → "שקופית N · translated"      (PPTX)
 *   "Slide N"                 → "שקופית N"
 *   "Page N - ObjectName"     → "עמוד N · translated"        (DOCX)
 *   "Entire document - ..."   → "כל המסמך"
 *   "Sheet N - Range A1:B1"  → "גיליון N · טווח A1:B1"      (XLSX — N is sheet index)
 *   "Sheet N - ObjectName"   → "גיליון N · translated"
 *   "Sheet N"                → "גיליון N"
 *
 * Corrupted text (?????, replacement chars) is detected and replaced
 * with a neutral fallback so the UI never shows garbage.
 */
export function formatOccurrenceLocation(loc) {
  if (!loc) return ''

  // Reject entirely corrupted strings before any parsing
  if (isCorruptedText(loc)) return 'מיקום לא זמין'

  // PPTX — "Slide N - ObjectName"
  let m = loc.match(/^Slide\s+(\d+)\s+-\s+(.+)$/)
  if (m) return `שקופית ${m[1]} · ${translateShapeName(m[2])}`

  // PPTX — "Slide N" only
  m = loc.match(/^Slide\s+(\d+)$/)
  if (m) return `שקופית ${m[1]}`

  // DOCX — "Page N - ObjectName"
  m = loc.match(/^Page\s+(\d+)\s+-\s+(.+)$/)
  if (m) return `עמוד ${m[1]} · ${translateShapeName(m[2])}`

  // DOCX — "Entire document ..."
  if (loc.startsWith('Entire document')) return 'כל המסמך'

  // XLSX — "Sheet N - Range A1:B1"  (N is always a sheet index number from the PS script)
  m = loc.match(/^Sheet\s+(\d+)\s+-\s+Range\s+(.+)$/)
  if (m) return `גיליון ${m[1]} · טווח ${m[2]}`

  // XLSX — "Sheet N - Range A1:B1"  (non-numeric sheet label, legacy)
  m = loc.match(/^Sheet\s+(.+?)\s+-\s+Range\s+(.+)$/)
  if (m) {
    const label = isCorruptedText(m[1]) ? '?' : m[1]
    return `גיליון ${label} · טווח ${m[2]}`
  }

  // XLSX — "Sheet N - ObjectName"
  m = loc.match(/^Sheet\s+(\d+)\s+-\s+(.+)$/)
  if (m) return `גיליון ${m[1]} · ${translateShapeName(m[2])}`

  // XLSX — "Sheet X - ObjectName"  (non-numeric, legacy)
  m = loc.match(/^Sheet\s+(.+?)\s+-\s+(.+)$/)
  if (m) {
    const label = isCorruptedText(m[1]) ? '?' : m[1]
    return `גיליון ${label} · ${translateShapeName(m[2])}`
  }

  // XLSX — "Sheet N" (index only)
  m = loc.match(/^Sheet\s+(\d+)$/)
  if (m) return `גיליון ${m[1]}`

  // XLSX — "Sheet X" (name, non-numeric)
  m = loc.match(/^Sheet\s+(.+)$/)
  if (m) return isCorruptedText(m[1]) ? 'גיליון' : `גיליון ${m[1]}`

  return isCorruptedText(loc) ? 'מיקום לא זמין' : loc
}

/**
 * Returns a helpful Hebrew hint for shape types that users might not recognise.
 * Keyed from the raw (untranslated) location string.
 */
export function getShapeTypeHint(loc) {
  if (!loc) return null
  const lower = loc.toLowerCase()
  if (lower.includes('freeform')) return SHAPE_TYPE_HINTS.freeform
  if (/\bgroup\b/.test(lower))   return SHAPE_TYPE_HINTS.group
  return null
}

// Severity classification used in buildOfficeEngineSummary counting
// PPTX keys: missingAltText, contrast, mediaCaptions, missingSlideTitle, duplicateSlideTitle, tableHeader, mergedCells
// DOCX keys: missingAltText, contrast, tableHeader, mergedCells, unclearHyperlinkText, documentTitle
// XLSX keys: missingAltText, contrast, tableHeader, mergedCells, unclearHyperlinkText, sheetName
const HIGH_KEYS   = ['missingAltText', 'contrast', 'mediaCaptions']
const MEDIUM_KEYS = [
  'missingSlideTitle', 'duplicateSlideTitle',
  'tableHeader', 'mergedCells',
  'unclearHyperlinkText', 'documentTitle',
  'sheetName', 'noHeadings',
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
  noHeadings:           'medium',
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
  // DOCX-specific
  noHeadings:           'החילו סגנונות כותרת מובנים של Word כמו Heading 1, Heading 2 ו-Heading 3 במקום עיצוב ידני.',
}

export const OFFICE_KEY_IMPACT = {
  missingAltText:       'משתמשי קוראי מסך לא יוכלו להבין את תוכן התמונה. עיוורים ולקויי ראייה יפספסו מידע חיוני המוצג באמצעות תמונות, תרשימים ואיורים.',
  contrast:             'טקסט עם ניגודיות נמוכה קשה לקריאה לאנשים עם לקויות ראייה, ראייה מוחלשת או עיוורון צבעים.',
  mediaCaptions:        'תוכן אודיו ווידאו ללא כתוביות אינו נגיש לחרשים, לכבדי שמיעה ולמשתמשים בסביבות ללא קול.',
  tableHeader:          'טבלה ללא שורת כותרת מונעת מקוראי מסך לקשר בין הנתונים לכותרות העמודות, ומקשה על הבנת ההקשר.',
  mergedCells:          'תאים ממוזגים משבשים את סדר הקריאה הלינארי של קוראי מסך ומקשים על ניווט ופרשנות הטבלה.',
  missingSlideTitle:    'קוראי מסך לא יכולים לנווט בין שקופיות ללא כותרת; משתמשים עם מוגבלות לא ידעו היכן הם נמצאים במצגת.',
  duplicateSlideTitle:  'כותרות כפולות מבלבלות משתמשי קוראי מסך שמנווטים לפי שמות שקופיות ברשימת תוכן.',
  readingOrder:         'סדר קריאה שגוי גורם לקוראי מסך להציג את תוכן השקופית בסדר לא הגיוני, מה שמוביל לאי-הבנה.',
  unclearHyperlinkText: 'קישורים עם טקסט כמו "לחצו כאן" אינם מתארים את היעד. משתמשי קוראי מסך המנווטים לפי רשימת קישורים לא ידעו לאן מוביל הקישור.',
  documentTitle:        'מסמך ללא כותרת בנכסי הקובץ קשה לזיהוי על ידי טכנולוגיות סיוע, דפדפנים ומנהלי קבצים.',
  noHeadings:           'מסמך ללא כותרות מונע מקוראי מסך לנווט בין חלקי התוכן ולהבין את מבנה המסמך; כל התוכן מוצג כגוש טקסט אחד רציף.',
  sheetName:            'שמות גיליונות כמו "Sheet1" אינם מתארים את תוכן הגיליון. משתמשי קוראי מסך לא ידעו לאיזה גיליון לעבור.',
  restrictedAccess:     'הגבלות מסמך עלולות למנוע מטכנולוגיות סיוע לגשת לתוכן, לקרוא אותו בקול או לנווט בו.',
}

export const OFFICE_KEY_WCAG = {
  missingAltText:       '1.1.1 Non-text Content · Level A',
  contrast:             '1.4.3 Contrast (Minimum) · Level AA',
  mediaCaptions:        '1.2.2 Captions (Prerecorded) · Level A',
  tableHeader:          '1.3.1 Info and Relationships · Level A',
  mergedCells:          '1.3.1 Info and Relationships · Level A',
  missingSlideTitle:    '2.4.2 Page Titled · Level A',
  duplicateSlideTitle:  '2.4.6 Headings and Labels · Level AA',
  readingOrder:         '1.3.2 Meaningful Sequence · Level A',
  unclearHyperlinkText: '2.4.4 Link Purpose (In Context) · Level A',
  documentTitle:        '2.4.2 Page Titled · Level A',
  noHeadings:           '2.4.6 Headings and Labels · Level AA',
  sheetName:            '2.4.6 Headings and Labels · Level AA',
  restrictedAccess:     '—',
}

export const OFFICE_KEY_FIX_STEPS = {
  missingAltText: [
    'לחצו לחיצה ימנית על התמונה ובחרו "Edit Alt Text" / "עריכת טקסט חלופי".',
    'הזינו תיאור קצר ומדויק של מה שמוצג בתמונה.',
    'אם התמונה דקורטיבית בלבד — סמנו "Mark as decorative" / "סמן כדקורטיבי".',
    'לחצו OK ושמרו את הקובץ.',
    'הריצו מחדש את בדיקת הנגישות לאימות התיקון.',
  ],
  contrast: [
    'בחרו את הטקסט בעל הניגודיות הנמוכה.',
    'שנו את צבע הטקסט לצבע כהה יותר (לדוגמה: שחור או אפור כהה על רקע בהיר).',
    'ודאו יחס ניגודיות של 4.5:1 לפחות לטקסט רגיל, 3:1 לטקסט גדול (מעל 18pt).',
    'השתמשו בכלי WebAIM Contrast Checker לבדיקת היחס.',
    'הימנעו משילובי צבעים בעייתיים כמו צהוב על לבן, אפור בהיר על לבן.',
  ],
  mediaCaptions: [
    'בחרו את קובץ האודיו/וידאו בשקופית.',
    'הוסיפו כתוביות דרך Insert → Captions (אם הגרסה תומכת בכך).',
    'לחלופין, צרו קובץ כתוביות SRT חיצוני ושמרו אותו לצד קובץ המצגת.',
    'הוסיפו תמלול טקסטואלי מלא בשקופית שאחרי אובייקט המדיה.',
  ],
  tableHeader: [
    'לחצו בתוך הטבלה.',
    'עברו ל-Table Design (PowerPoint/Word) או Table Tools → Design (Excel).',
    'סמנו את האפשרות "Header Row" / "שורת כותרת".',
    'ודאו שהשורה הראשונה מכילה כותרות תיאוריות לכל עמודה.',
    'שמרו ובדקו עם Accessibility Checker.',
  ],
  mergedCells: [
    'אתרו את התאים הממוזגים — Accessibility Checker מציין את מיקומם.',
    'בחרו את התא הממוזג.',
    'לחצו לחיצה ימנית → "Unmerge Cells" / "בטל מיזוג תאים".',
    'שקלו עיצוב חלופי שאינו מסתמך על מיזוג: שורות/עמודות קבוצות, צבע רקע, גבולות.',
    'ודאו שהנתונים עדיין קריאים לאחר ביטול המיזוג.',
  ],
  missingSlideTitle: [
    'עברו לשקופית ללא כותרת (מסומנת על ידי Accessibility Checker).',
    'אם אין מסגרת כותרת — הוסיפו: Insert → Text Box → Title placeholder.',
    'לחלופין, שנו את פריסת השקופית ל-Layout שכולל מסגרת כותרת.',
    'הזינו כותרת ייחודית ומתארת לכל שקופית.',
    'חזרו על הפעולה לכל שקופית שמסומנת כחסרת כותרת.',
  ],
  duplicateSlideTitle: [
    'אתרו את השקופיות עם הכותרת הכפולה.',
    'שנו כל כותרת כפולה לכותרת ייחודית שמתארת את תוכן השקופית הספציפית.',
    'השתמשו בחלונית Slide Panel לניווט מהיר וסקירת כל הכותרות.',
    'כותרות כמו "המשך", "עוד", "דוגמה" — הפכו לכותרות ספציפיות כמו "דוגמה: ניתוח נתונים".',
  ],
  readingOrder: [
    'פתחו את חלונית Selection Pane: Home → Arrange → Selection Pane.',
    'הסדר בחלונית מלמטה למעלה הוא סדר הקריאה בפועל.',
    'גררו אובייקטים לסדר הנכון (תוכן ראשי קודם, הערות אחרונות).',
    'ודאו שכל אובייקט בעל שם תיאורי (לחצו פעמיים על שמו בחלונית לשינוי).',
    'הריצו Accessibility Checker לאימות.',
  ],
  unclearHyperlinkText: [
    'בחרו את הקישור עם הטקסט הלא ברור.',
    'לחצו לחיצה ימנית → Edit Hyperlink / עריכת היפר-קישור.',
    'שנו את שדה "Text to display" לטקסט מתאר, לדוגמה: "מדריך נגישות דיגיטלית של W3C".',
    'הימנעו מ: "לחצו כאן", "קישור", "פה", "כאן", URL גולמי.',
    'שמרו ובדקו.',
  ],
  documentTitle: [
    'פתחו File → Info.',
    'לחצו על שדה "Title" תחת Properties → Add a title.',
    'הזינו כותרת מסמך ברורה ומתארת את תוכן הקובץ.',
    'שמרו את המסמך.',
    'הריצו Accessibility Checker לאימות.',
  ],
  noHeadings: [
    'סמנו את הטקסט המשמש ככותרת ראשית של סעיף.',
    'עברו ל-Home → Styles.',
    'בחרו "Heading 1" לכותרות ראשיות, "Heading 2" לכותרות משנה, "Heading 3" לרמה שלישית.',
    'אל תשתמשו בהגדלת גופן, הדגשה ידנית או צבע במקום סגנונות Heading.',
    'ודאו שישנה לפחות כותרת Heading 1 אחת במסמך.',
    'הריצו Accessibility Checker לאימות.',
  ],
  sheetName: [
    'לחצו לחיצה ימנית על לשונית הגיליון בתחתית חלון Excel.',
    'בחרו "Rename" / "שנה שם".',
    'הזינו שם תיאורי המשקף את תוכן הגיליון (לדוגמה: "מכירות Q1 2024", "נתוני עובדים").',
    'הימנעו מ: "Sheet1", "Sheet2", "גיליון1".',
    'חזרו על הפעולה לכל לשונית גיליון.',
  ],
  restrictedAccess: [
    'פתחו File → Info → Protect Document / Workbook / Presentation.',
    'בדקו אילו הגבלות הוגדרו (הצפנה, הגבלת עריכה, סיסמה).',
    'הסירו הגבלות מיותרות שאינן נדרשות לאבטחה.',
    'אם ההגבלות הכרחיות — ודאו שהן אינן חוסמות טכנולוגיות סיוע כמו NVDA, JAWS.',
    'שמרו ובדקו עם קורא מסך.',
  ],
}

export const OFFICE_REPORT_GROUPS = [
  { label: 'צבע וניגודיות',  keys: ['contrast'] },
  { label: 'מדיה ואיורים',   keys: ['missingAltText', 'mediaCaptions'] },
  { label: 'טבלאות',          keys: ['tableHeader', 'mergedCells'] },
  // PPTX structure keys + DOCX documentTitle + DOCX noHeadings + XLSX sheetName
  { label: 'מבנה מסמך',      keys: ['missingSlideTitle', 'duplicateSlideTitle', 'readingOrder', 'documentTitle', 'noHeadings', 'sheetName'] },
  // DOCX + XLSX hyperlink key
  { label: 'קישורים וניווט', keys: ['unclearHyperlinkText'] },
  { label: 'גישה למסמך',     keys: ['restrictedAccess'] },
]

/**
 * Returns true for generic "Checked by Microsoft … Accessibility Checker." notes
 * that come from the backend when a check passes.
 * These should be suppressed inside individual warning cards — the engine is already
 * shown once at the section level.
 */
export function isGenericCheckedByNote(note) {
  if (!note) return false
  return (
    /נבדק על ידי Microsoft .* Accessibility Checker/i.test(note) ||
    /Checked by Microsoft .* Accessibility Checker/i.test(note)
  )
}

/**
 * Returns true when a scan result should use Office Engine data as the primary source.
 */
export function isOfficeEngineResult(result) {
  return result?.engine === 'office-engine' && !!result?.officeLikeSummary
}

// Capped penalty model — prevents one heavily-occurring rule from collapsing the score to 0.
// Each failed rule contributes: min(count * perOcc, maxPerRule) to total penalty.
// Manual rules contribute a flat penalty only when count > 0 (something needs review).
const PENALTY = {
  high:   { perOcc: 2.0, maxPerRule: 35 },
  medium: { perOcc: 1.0, maxPerRule: 20 },
  low:    { perOcc: 0.5, maxPerRule: 10 },
  manual: 3,   // flat per rule when count > 0
}

/**
 * One shared score formula for all views (ResultsPage, ReportPage, PDF).
 * Uses a capped per-rule penalty so no single rule can alone collapse the score to 0.
 */
export function calculateOfficeEngineScore(officeLikeSummary) {
  if (!officeLikeSummary) return 0

  let totalPenalty = 0

  Object.entries(officeLikeSummary).forEach(([key, item]) => {
    const count  = item?.count  ?? 0
    const status = item?.status

    if (['manual', 'partial', 'not_checked'].includes(status)) {
      if (count > 0) totalPenalty += PENALTY.manual
    } else if (status === 'failed' && count > 0) {
      const sev = OFFICE_KEY_SEVERITY[key] ?? 'medium'
      const cfg = PENALTY[sev] ?? PENALTY.medium
      totalPenalty += Math.min(count * cfg.perOcc, cfg.maxPerRule)
    }
    // passed items with count 0 → no penalty
  })

  return Math.max(0, Math.min(100, Math.round(100 - totalPenalty)))
}

/**
 * Compute display-ready aggregate numbers from officeLikeSummary.
 * - totalFindings: sum of all positive item.count values
 * - issueTypes:    number of items with count > 0
 * - highCount:     sum of counts for high-severity failed items
 * - mediumCount:   sum of counts for medium-severity failed items
 * - manualCount:   number of manual/partial/not_checked items
 * - score:         calculateOfficeEngineScore(officeLikeSummary)
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

  const score = calculateOfficeEngineScore(officeLikeSummary)
  return { totalFindings, issueTypes, highCount, mediumCount, manualCount, score }
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
      const rawNote = item.note ?? null
      rows.push({
        key,
        label:          item.label,
        count,
        locations:      item.locations ?? [],
        status,
        severity:       isManual ? 'manual' : (OFFICE_KEY_SEVERITY[key] ?? 'medium'),
        category:       group.label,
        recommendation: OFFICE_KEY_RECOMMENDATION[key] ?? '',
        impact:         OFFICE_KEY_IMPACT[key]     ?? null,
        wcag:           OFFICE_KEY_WCAG[key]        ?? null,
        fixSteps:       OFFICE_KEY_FIX_STEPS[key]  ?? [],
        note:           isGenericCheckedByNote(rawNote) ? null : rawNote,
      })
    }
  }
  return rows
}
