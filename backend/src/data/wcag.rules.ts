/** Central WCAG 2.1 mapping – single source of truth for all scan issues */
export interface WcagRule {
  criterion: string
  principle: string
  level: string
  explanationHe: string
}

export const WCAG_RULES = {
  'non-text-content': {
    criterion: '1.1.1 Non-text Content',
    principle: 'Perceivable',
    level: 'A',
    explanationHe:
      'תוכן שאינו טקסטואלי צריך לכלול חלופה טקסטואלית שמשרתת מטרה מקבילה.',
  },
  'contrast-minimum': {
    criterion: '1.4.3 Contrast (Minimum)',
    principle: 'Perceivable',
    level: 'AA',
    explanationHe:
      'טקסט צריך להיות מוצג ביחס ניגודיות מספיק מול הרקע כדי לאפשר קריאה למשתמשים עם לקות ראייה.',
  },
  'meaningful-sequence': {
    criterion: '1.3.2 Meaningful Sequence',
    principle: 'Perceivable',
    level: 'A',
    explanationHe:
      'כאשר סדר הקריאה משפיע על משמעות התוכן, צריך להיות ניתן לקבוע את הסדר בצורה תוכנתית.',
  },
  'link-purpose': {
    criterion: '2.4.4 Link Purpose (In Context)',
    principle: 'Operable',
    level: 'A',
    explanationHe: 'מטרת הקישור צריכה להיות ברורה מתוך טקסט הקישור או מתוך ההקשר שלו.',
  },
  'page-titled-headings': {
    criterion: '2.4.2 Page Titled / 2.4.6 Headings and Labels',
    principle: 'Operable',
    level: 'A/AA',
    explanationHe: 'כותרות ברורות עוזרות למשתמש להבין את מבנה המסמך ולנווט בין חלקיו.',
  },
  'resize-text': {
    criterion: '1.4.4 Resize Text',
    principle: 'Perceivable',
    level: 'AA',
    explanationHe: 'יש לאפשר קריאה נוחה והגדלת טקסט בלי לאבד תוכן או פונקציונליות.',
  },
  'info-relationships': {
    criterion: '1.3.1 Info and Relationships',
    principle: 'Understandable',
    level: 'A',
    explanationHe:
      'מידע, מבנה והקשרים בין רכיבי התוכן צריכים להיות ניתנים לזיהוי תוכנתי על ידי טכנולוגיות מסייעות.',
  },
  'headings-labels': {
    criterion: '2.4.6 Headings and Labels',
    principle: 'Understandable',
    level: 'AA',
    explanationHe: 'כותרות ותוויות תיאוריות מסייעות להבנת מבנה המסמך ותוכנו.',
  },
  'reading-level': {
    criterion: '3.1.5 Reading Level',
    principle: 'Understandable',
    level: 'AAA',
    explanationHe: 'תוכן ארוך ומורכב מדי עלול להקשות על קריאה והבנה, במיוחד למשתמשים עם לקויות קוגניטיביות.',
  },
} as const satisfies Record<string, WcagRule>

export type WcagRuleKey = keyof typeof WCAG_RULES

export function getWcagRule(key: string): WcagRule {
  const rule = WCAG_RULES[key as WcagRuleKey]
  if (rule) return rule
  return {
    criterion: 'WCAG 2.1',
    principle: 'General',
    level: '—',
    explanationHe: 'לא נמצא מיפוי WCAG עבור בעיה זו.',
  }
}
