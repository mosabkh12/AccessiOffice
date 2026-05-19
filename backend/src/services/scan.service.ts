import type { WcagRuleKey } from '../data/wcag.rules.js'
import { getWcagRule } from '../data/wcag.rules.js'
import type { FileType, ScanIssue, ScanResult, ScanType, Severity, UserType } from '../types/scan.types.js'

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

const DOCX_ISSUES: IssueTemplate[] = [
  {
    id: 'missing-alt-text',
    wcagKey: 'non-text-content',
    title: 'תמונה ללא טקסט חלופי',
    severity: 'High',
    category: 'תוכן',
    affectedUsers: ['משתמשים עיוורים', 'משתמשי קורא מסך'],
    impact: 'קוראי מסך אינם יכולים להבין את מטרת התמונה או את תוכנה בהקשר המסמך.',
    recommendation: 'הוסיפו טקסט חלופי ברור לכל תמונה, או סמנו תמונות דקורטיביות בהתאם.',
  },
  {
    id: 'incorrect-heading-structure',
    wcagKey: 'info-relationships',
    title: 'מבנה כותרות שגוי',
    severity: 'High',
    category: 'מבנה',
    affectedUsers: ['משתמשי קורא מסך', 'משתמשים עם לקויות קוגניטיביות'],
    impact: 'דילוג ברמות כותרות מקשה על ניווט והבנת מבנה המסמך.',
    recommendation: 'השתמשו בסגנונות כותרת של Word (כותרת 1, 2, 3) בסדר היררכי ללא דילוגים.',
  },
  {
    id: 'unclear-link-text',
    wcagKey: 'link-purpose',
    title: 'טקסט קישור לא ברור',
    severity: 'Medium',
    category: 'ניווט',
    affectedUsers: ['משתמשי קורא מסך'],
    impact: 'ביטויים כמו "לחץ כאן" אינם מסבירים לאן מוביל הקישור.',
    recommendation: 'השתמשו בטקסט קישור תיאורי, למשל: "הורדת מדריך הנגישות (PDF)".',
  },
  {
    id: 'table-without-header',
    wcagKey: 'info-relationships',
    title: 'טבלה ללא שורת כותרת',
    severity: 'Medium',
    category: 'טבלאות',
    affectedUsers: ['משתמשי קורא מסך'],
    impact: 'קשה לקשר בין נתונים לעמודות ולשורות ללא כותרות מוגדרות.',
    recommendation: 'סמנו את השורה הראשונה כשורת כותרת בהגדרות הטבלה.',
  },
  {
    id: 'low-color-contrast',
    wcagKey: 'contrast-minimum',
    title: 'ניגודיות צבעים נמוכה',
    severity: 'High',
    category: 'תצוגה',
    affectedUsers: ['משתמשים עם לקות ראייה', 'משתמשים בסביבה מוארת'],
    impact: 'טקסט בהיר על רקע בהיר אינו קריא עבור משתמשים רבים.',
    recommendation: 'הבטיחו יחס ניגודיות מינימלי של 4.5:1 לטקסט רגיל.',
  },
]

const PPTX_ISSUES: IssueTemplate[] = [
  {
    id: 'slide-without-title',
    wcagKey: 'page-titled-headings',
    title: 'שקופית ללא כותרת',
    severity: 'Medium',
    category: 'מבנה',
    affectedUsers: ['משתמשי קורא מסך'],
    impact: 'משתמשים אינם יכולים לזהות במהירות את נושא השקופית.',
    recommendation: 'הוסיפו כותרת ייחודית לכל שקופית באמצעות תיבת הכותרת.',
  },
  {
    id: 'image-without-alt-text',
    wcagKey: 'non-text-content',
    title: 'תמונה ללא טקסט חלופי',
    severity: 'High',
    category: 'תוכן',
    affectedUsers: ['משתמשים עיוורים', 'משתמשי קורא מסך'],
    impact: 'מידע ויזואלי בשקופית אינו זמין למשתמשי טכנולוגיות מסייעות.',
    recommendation: 'ערכו טקסט חלופי משמעותי לכל תמונה דרך "עריכת טקסט חלופי".',
  },
  {
    id: 'low-color-contrast',
    wcagKey: 'contrast-minimum',
    title: 'ניגודיות צבעים נמוכה',
    severity: 'High',
    category: 'תצוגה',
    affectedUsers: ['משתמשים עם לקות ראייה'],
    impact: 'טקסט על רקע דקורטיבי עלול שלא לעמוד בדרישות הניגודיות.',
    recommendation: 'השתמשו בצירופי צבעים בעלי ניגודיות גבוהה לכל הטקסט בשקופית.',
  },
  {
    id: 'reading-order-issue',
    wcagKey: 'meaningful-sequence',
    title: 'בעיית סדר קריאה',
    severity: 'Medium',
    category: 'מבנה',
    affectedUsers: ['משתמשי קורא מסך'],
    impact: 'התוכן עלול להיקרא בסדר שאינו תואם את הפריסה הוויזואלית.',
    recommendation: 'בדקו ותקנו את סדר הקריאה בחלונית "סדר קריאה" של PowerPoint.',
  },
  {
    id: 'small-font-size',
    wcagKey: 'resize-text',
    title: 'גודל גופן קטן מדי',
    severity: 'Low',
    category: 'תצוגה',
    affectedUsers: ['משתמשים עם לקות ראייה', 'אנשים מבוגרים'],
    impact: 'טקסט קטן מאוד קשה לקריאה גם כשהניגודיות תקינה.',
    recommendation: 'השתמשו בגודל מינימלי של 18 נקודות לגוף טקסט ו-24 לכותרות שקופית.',
  },
]

const XLSX_ISSUES: IssueTemplate[] = [
  {
    id: 'table-without-header',
    wcagKey: 'info-relationships',
    title: 'טבלה ללא שורת כותרת',
    severity: 'High',
    category: 'טבלאות',
    affectedUsers: ['משתמשי קורא מסך'],
    impact: 'נתונים בגיליון אינם ניתנים לניווט לפי שמות עמודות ושורות.',
    recommendation: 'הגדירו שורת כותרת וסמנו "הטבלה שלי כוללת כותרות".',
  },
  {
    id: 'merged-cells-layout',
    wcagKey: 'info-relationships',
    title: 'תאים ממוזגים לפריסה',
    severity: 'High',
    category: 'טבלאות',
    affectedUsers: ['משתמשי קורא מסך'],
    impact: 'תאים ממוזגים פוגעים בניווט טבלאי תקין בקוראי מסך.',
    recommendation: 'הימנעו ממיזוג תאים; השתמשו ביישור ובעיצוב במקום.',
  },
  {
    id: 'empty-cells-spacing',
    wcagKey: 'info-relationships',
    title: 'תאים ריקים לריווח ויזואלי',
    severity: 'Medium',
    category: 'מבנה',
    affectedUsers: ['משתמשי קורא מסך'],
    impact: 'תאים ריקים לריווח יוצרים בלבול בניווט ובהבנת הנתונים.',
    recommendation: 'השתמשו בריווח, רוחב עמודות ומבנה טבלה תקין במקום תאים ריקים.',
  },
  {
    id: 'sheet-without-title',
    wcagKey: 'headings-labels',
    title: 'גיליון ללא שם תיאורי',
    severity: 'Medium',
    category: 'מבנה',
    affectedUsers: ['משתמשי קורא מסך', 'משתמשים עם לקויות קוגניטיביות'],
    impact: 'שמות גנריים כמו "Sheet1" אינם מסבירים את תוכן הגיליון.',
    recommendation: 'שנו את שם הגיליון לשם תיאורי, למשל: "נתוני מכירות Q1".',
  },
  {
    id: 'low-contrast-formatting',
    wcagKey: 'contrast-minimum',
    title: 'ניגודיות נמוכה בעיצוב תאים',
    severity: 'High',
    category: 'תצוגה',
    affectedUsers: ['משתמשים עם לקות ראייה'],
    impact: 'תאים עם צבעי רקע וגופן חלשים אינם קריאים למשתמשים רבים.',
    recommendation: 'החילו צבעי מילוי וגופן העומדים ביחס ניגודיות WCAG AA.',
  },
]

const LOCATIONS: Record<FileType, string[]> = {
  docx: ['עמוד 1 / תמונה 2', 'עמוד 1 / כותרות', 'עמוד 4 / קישור', 'עמוד 3 / טבלה', 'עמוד 2 / גוף טקסט'],
  pptx: ['שקופית 2', 'שקופית 3 / תמונה 1', 'שקופית 5 / תיבת טקסט', 'שקופית 4 / תוכן', 'שקופית 6 / כותרת תחתונה'],
  xlsx: ['גיליון1 / A1:F20', 'גיליון1 / B5:D8', 'גיליון1 / שורות 10–15', 'גיליון2 (שם לשונית)', 'גיליון1 / C3:E12'],
}

function buildIssue(template: IssueTemplate, location: string): ScanIssue {
  const wcag = getWcagRule(template.wcagKey)
  return {
    id: template.id,
    title: template.title,
    severity: template.severity,
    category: template.category,
    wcagPrinciple: wcag.principle,
    wcagCriterion: wcag.criterion,
    wcagLevel: wcag.level,
    wcagExplanation: wcag.explanationHe,
    affectedUsers: template.affectedUsers,
    impact: template.impact,
    recommendation: template.recommendation,
    location,
  }
}

function buildIssues(fileType: FileType, templates: IssueTemplate[]): ScanIssue[] {
  const locations = LOCATIONS[fileType]
  return templates.map((t, i) => buildIssue(t, locations[i] ?? 'מסמך'))
}

function calculateScore(issues: ScanIssue[]): number {
  const weights = { High: 12, Medium: 7, Low: 3 } as const
  const penalty = issues.reduce((sum, issue) => sum + weights[issue.severity], 0)
  return Math.max(0, 100 - penalty)
}

function buildSummary(issues: ScanIssue[]) {
  return {
    totalIssues: issues.length,
    high: issues.filter((i) => i.severity === 'High').length,
    medium: issues.filter((i) => i.severity === 'Medium').length,
    low: issues.filter((i) => i.severity === 'Low').length,
  }
}

export const CRITICAL_ANALYSIS_HE =
  'סריקה אוטומטית מסייעת לזהות בעיות נגישות נפוצות, אך אינה מחליפה בדיקה אנושית מקצועית. יש לבדוק ידנית את איכות טקסט חלופי, הקשר תוכן, סדר קריאה מורכב והתאמה מלאה לתקן הישראלי (ת"י 5568) לפני פרסום המסמך.'

export function generateScanResult(
  fileName: string,
  fileType: FileType,
  userType: UserType,
  scanType: ScanType,
): ScanResult {
  let templates: IssueTemplate[]
  switch (fileType) {
    case 'docx':
      templates = DOCX_ISSUES
      break
    case 'pptx':
      templates = PPTX_ISSUES
      break
    case 'xlsx':
      templates = XLSX_ISSUES
      break
  }

  const issues = buildIssues(fileType, templates)

  return {
    fileName,
    fileType,
    userType,
    scanType,
    score: calculateScore(issues),
    summary: buildSummary(issues),
    issues,
    criticalAnalysis: CRITICAL_ANALYSIS_HE,
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
