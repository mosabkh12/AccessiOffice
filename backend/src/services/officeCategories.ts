/**
 * Categories aligned with Microsoft Office Accessibility Checker groups.
 * Hebrew labels for UI; English comments for reference.
 */

/** צבע וניגודיות — Color and Contrast */
export const CAT_COLOR_CONTRAST = 'צבע וניגודיות'

/** מדיה ואיורים — Media and Illustrations */
export const CAT_MEDIA = 'מדיה ואיורים'

/** טבלאות — Tables */
export const CAT_TABLES = 'טבלאות'

/** מבנה מסמך — Document Structure */
export const CAT_DOCUMENT_STRUCTURE = 'מבנה מסמך'

/** גישה למסמך — Document Access */
export const CAT_DOCUMENT_ACCESS = 'גישה למסמך'

/** Office-style issue titles (Hebrew) */
export const OFFICE_TITLES = {
  hardTextContrast: 'ניגודיות טקסט קשה לקריאה',
  missingAltText: 'חסר טקסט חלופי',
  missingTableHeader: 'חסרת שורת כותרת בטבלה',
  mergedOrSplitCells: 'שימוש בתאים ממוזגים או מפוצלים',
  noHeadingsInDocument: 'אין כותרות במסמך',
  visualHeadingNoStyle: 'טקסט שנראה ככותרת ללא סגנון כותרת',
  headingLevelSkip: 'מבנה כותרות לא תקין',
  unclearHyperlink: 'טקסט קישור לא ברור',
  restrictedAccess: 'גישה מוגבלת',
  slideWithoutTitle: 'שקופית ללא כותרת',
  duplicateSlideTitle: 'כותרת שקופית כפולה',
  readingOrderCheck: 'בדיקת סדר קריאה',
  decorativeCandidates: 'אובייקטים שניתן לסמן כדקורטיביים',
  smallFont: 'גודל גופן קטן מדי',
  emptyParagraphSpacing: 'שימוש ברווחים ריקים לעיצוב',
  longParagraph: 'פסקה ארוכה וקשה לקריאה',
  genericSheetName: 'שם גיליון לא תיאורי',
  columnsWithoutHeader: 'עמודות ללא כותרת ברורה',
  emptyCellsInRange: 'תאים ריקים בתוך טווח נתונים',
  limitedScan: 'ניתוח מבני מוגבל',
} as const
