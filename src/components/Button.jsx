import { Link } from 'react-router-dom'

export default function Button({ children, to, onClick, variant = 'primary', disabled, type = 'button' }) {
  const className = `btn btn-${variant}`

  if (to) {
    return (
      <Link to={to} className={className}>
        {children}
      </Link>
    )
  }

  return (
    <button type={type} className={className} onClick={onClick} disabled={disabled}>
      {children}
    </button>
  )
}
