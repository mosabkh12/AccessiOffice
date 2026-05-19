import Button from '../components/Button.jsx'

const FILE_TYPES = [
  {
    icon: '📄',
    title: 'Word (.docx)',
    desc: 'בדיקת מסמכי טקסט, כותרות, קישורים וטבלאות.',
    checks: ['טקסט חלופי', 'מבנה כותרות', 'קישורים', 'ניגודיות'],
  },
  {
    icon: '📊',
    title: 'PowerPoint (.pptx)',
    desc: 'בדיקת מצגות, שקופיות, סדר קריאה וניגודיות.',
    checks: ['כותרות שקופית', 'תמונות', 'סדר קריאה', 'גודל גופן'],
  },
  {
    icon: '📈',
    title: 'Excel (.xlsx)',
    desc: 'בדיקת גיליונות, טבלאות ומבנה נתונים.',
    checks: ['שורת כותרת', 'תאים ממוזגים', 'שם גיליון', 'ניגודיות'],
  },
]

const STEPS = [
  { num: '1', label: 'העלאת קובץ' },
  { num: '2', label: 'בחירת סוג משתמש' },
  { num: '3', label: 'סריקת נגישות' },
  { num: '4', label: 'הפקת דוח' },
]

export default function HomePage() {
  return (
    <div className="page page-home">
      <section className="hero-panel card">
        <div className="hero-panel__content">
          <p className="eyebrow">מערכת בדיקת נגישות אקדמית</p>
          <h1>AccessiOffice</h1>
          <p className="hero-subtitle">בודק נגישות לקבצי Office</p>
          <p className="hero-description">
            AccessiOffice הוא אב־טיפוס Web לבדיקת נגישות במסמכי Word, PowerPoint ו־Excel.
            המערכת מזהה בעיות נפוצות, מסבירה את ההשפעה על משתמשים עם מוגבלות, ומפיקה דוח תיקון
            לפי WCAG 2.1 ות&quot;י 5568.
          </p>
          <div className="hero-actions">
            <Button to="/upload">התחלת סריקת נגישות</Button>
            <p className="hero-formats">תומך בקבצי docx, pptx, xlsx</p>
          </div>
        </div>
      </section>

      <section className="section-block">
        <h2 className="section-title">איך זה עובד</h2>
        <ol className="how-steps">
          {STEPS.map((s) => (
            <li key={s.num} className="how-step">
              <span className="how-step__num">{s.num}</span>
              <span>{s.label}</span>
            </li>
          ))}
        </ol>
      </section>

      <section className="section-block">
        <h2 className="section-title">סוגי קבצים נתמכים</h2>
        <div className="file-type-grid">
          {FILE_TYPES.map((ft) => (
            <article key={ft.title} className="file-type-card card">
              <span className="file-type-card__icon" aria-hidden="true">{ft.icon}</span>
              <h3>{ft.title}</h3>
              <p>{ft.desc}</p>
              <ul className="check-tags">
                {ft.checks.map((c) => (
                  <li key={c}>{c}</li>
                ))}
              </ul>
            </article>
          ))}
        </div>
      </section>

      <section className="section-block">
        <article className="standards-card card">
          <h2 className="section-title">תקנים ודוחות</h2>
          <ul className="standards-list">
            <li><strong>WCAG 2.1</strong> — קריטריונים בינלאומיים לנגישות תוכן</li>
            <li><strong>ת&quot;י 5568</strong> — תקן הנגישות הישראלי</li>
            <li><strong>דוח מותאם</strong> — לפי סוג משתמש: מחבר, בודק או מוסד לימודי</li>
          </ul>
        </article>
      </section>
    </div>
  )
}
