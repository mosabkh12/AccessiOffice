import { NavLink } from 'react-router-dom'
import { useScan } from '../context/ScanContext.jsx'

const NAV = [
  { to: '/', label: 'בית', end: true },
  { to: '/upload', label: 'סריקה' },
  { to: '/results', label: 'תוצאות', needsResults: true },
  { to: '/report', label: 'דוח', needsResults: true },
]

export default function Layout({ children }) {
  const { scanData } = useScan()
  const hasResults = Boolean(scanData?.results)

  return (
    <div className="layout">
      <header className="layout-header">
        <div className="layout-header-inner">
          <NavLink to="/" className="brand">
            <span className="logo-icon" aria-hidden="true">♿</span>
            <span className="logo-text">
              <strong>AccessiOffice</strong>
              <span>בודק נגישות לקבצי Office</span>
            </span>
          </NavLink>
          <nav className="layout-nav" aria-label="ניווט ראשי">
            {NAV.map(({ to, label, end, needsResults }) => {
              const disabled = needsResults && !hasResults
              if (disabled) {
                return (
                  <span key={to} className="nav-link nav-link--disabled" aria-disabled="true">
                    {label}
                  </span>
                )
              }
              return (
                <NavLink
                  key={to}
                  to={to}
                  end={end}
                  className={({ isActive }) => `nav-link${isActive ? ' nav-link--active' : ''}`}
                >
                  {label}
                </NavLink>
              )
            })}
          </nav>
        </div>
      </header>

      <main className="layout-main">{children}</main>

      <footer className="layout-footer">
        <p>AccessiOffice © 2026</p>
        <p>בדיקת נגישות לפי WCAG 2.1 ות&quot;י 5568</p>
      </footer>
    </div>
  )
}
