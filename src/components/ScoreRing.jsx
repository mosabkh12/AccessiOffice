export default function ScoreRing({ score }) {
  const radius = 52
  const circumference = 2 * Math.PI * radius
  const offset = circumference - (score / 100) * circumference

  const color = score >= 85 ? '#38a169' : score >= 65 ? '#d69e2e' : '#c53030'

  return (
    <div className="score-ring" role="img" aria-label={`ציון נגישות: ${score} מתוך 100`}>
      <svg width="120" height="120" viewBox="0 0 120 120">
        <circle cx="60" cy="60" r={radius} fill="none" stroke="#e2e8f0" strokeWidth="10" />
        <circle
          cx="60"
          cy="60"
          r={radius}
          fill="none"
          stroke={color}
          strokeWidth="10"
          strokeDasharray={circumference}
          strokeDashoffset={offset}
          strokeLinecap="round"
        />
      </svg>
      <span className="score-ring-value">{score}</span>
    </div>
  )
}
