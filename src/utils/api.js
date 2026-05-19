const API_URL = import.meta.env.VITE_API_URL || 'http://localhost:4000'

const USER_TYPE_TO_API = {
  author: 'document-author',
  auditor: 'accessibility-auditor',
  lecturer: 'lecturer-institution',
}

const USER_TYPE_FROM_API = Object.fromEntries(
  Object.entries(USER_TYPE_TO_API).map(([k, v]) => [v, k]),
)

const FILE_LABELS = { docx: 'Word', pptx: 'PowerPoint', xlsx: 'Excel' }

const FALLBACK_NAMES = {
  docx: 'uploaded-document.docx',
  pptx: 'uploaded-presentation.pptx',
  xlsx: 'uploaded-spreadsheet.xlsx',
}

export function getFileType(filename) {
  const ext = filename?.split('.').pop()?.toLowerCase()
  if (['docx', 'pptx', 'xlsx'].includes(ext)) return ext
  return null
}

/** Client-side safe display name before/after API */
export function safeFileName(name, fileType) {
  const fallback = FALLBACK_NAMES[fileType] || 'uploaded-file'
  if (!name?.trim()) return fallback
  const multiplyCount = (name.match(/×/g) || []).length
  if (multiplyCount >= 2 && !/[\u0590-\u05FF]/.test(name)) return fallback
  if (/Ã.|Â.|×[§©ª«¬]/.test(name)) return fallback
  return name.trim()
}

export async function submitScan({ file, userType, scanType }) {
  const formData = new FormData()
  formData.append('file', file, file.name)
  formData.append('displayFileName', file.name)
  formData.append('userType', USER_TYPE_TO_API[userType] ?? userType)
  formData.append('scanType', scanType)

  const response = await fetch(`${API_URL}/api/scan`, {
    method: 'POST',
    body: formData,
  })

  const data = await response.json().catch(() => ({}))

  if (!response.ok) {
    throw new Error(data.error || 'הסריקה נכשלה. ודאו שהשרת פועל ונסו שוב.')
  }

  return normalizeScanResponse(data, file.name)
}

function normalizeIssue(raw) {
  const severity = String(raw?.severity || 'Medium')
  const severityKey = severity.toLowerCase()

  const affectedUsers = Array.isArray(raw?.affectedUsers)
    ? raw.affectedUsers
    : raw?.affectedUsers
      ? [String(raw.affectedUsers)]
      : ['לא צוין']

  const wcagCriterion = raw?.wcagCriterion || '—'
  const wcagLevel = raw?.wcagLevel || '—'
  const wcagPrinciple = raw?.wcagPrinciple || '—'

  return {
    id: raw?.id || `issue-${Math.random().toString(36).slice(2, 9)}`,
    title: raw?.title || 'בעיית נגישות לא מזוהה',
    severity: severityKey,
    severityLabel: severityKey === 'high' ? 'גבוהה' : severityKey === 'medium' ? 'בינונית' : 'נמוכה',
    category: raw?.category || 'כללי',
    wcagCriterion,
    wcagPrinciple,
    wcagLevel,
    wcagExplanation: raw?.wcagExplanation || 'לא זמין הסבר WCAG לבעיה זו.',
    wcagDisplay: `${wcagCriterion} (${wcagLevel})`,
    affectedUsers,
    affectedUsersText: affectedUsers.join(' · '),
    impact: raw?.impact || 'לא צוינה השפעה.',
    recommendation: raw?.recommendation || 'לא צוינה המלצה.',
    location: raw?.location || '—',
  }
}

export function normalizeScanResponse(api, clientFileName) {
  const userType = USER_TYPE_FROM_API[api.userType] ?? api.userType
  const fileType = api.fileType || getFileType(clientFileName) || 'docx'
  const fileName = safeFileName(api.fileName || clientFileName, fileType)
  const fileLabel = FILE_LABELS[fileType] ?? 'Office'

  const issues = (api.issues || []).map(normalizeIssue)
  const summary = api.summary || { totalIssues: issues.length, high: 0, medium: 0, low: 0 }

  let summaryText
  if (api.score >= 85) {
    summaryText = `קובץ ${fileLabel} עומד ברוב דרישות הנגישות, אך נמצאו ${summary.totalIssues} בעיות לטיפול.`
  } else if (api.score >= 65) {
    summaryText = `קובץ ${fileLabel} דורש שיפורים משמעותיים. נמצאו ${summary.totalIssues} בעיות.`
  } else {
    summaryText = `קובץ ${fileLabel} אינו עומד בדרישות נגישות בסיסיות. יש לתקן ${summary.totalIssues} בעיות לפני פרסום.`
  }

  const criticalAnalysis =
    api.criticalAnalysis ||
    'סריקה אוטומטית מסייעת לזהות בעיות נפוצות, אך אינה מחליפה בדיקה אנושית מקצועית.'

  return {
    fileName,
    fileType,
    userType,
    scanType: api.scanType,
    scannedAt: new Date().toLocaleDateString('he-IL', {
      year: 'numeric',
      month: 'long',
      day: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
    }),
    score: typeof api.score === 'number' ? api.score : 0,
    totalIssues: summary.totalIssues ?? issues.length,
    severityCounts: {
      high: summary.high ?? 0,
      medium: summary.medium ?? 0,
      low: summary.low ?? 0,
    },
    summary: summaryText,
    issues,
    recommendations: [...new Set(issues.map((i) => i.recommendation).filter(Boolean))],
    criticalAnalysis,
    limitations: [
      criticalAnalysis,
      'הכלי מנתח מבנה כללי ואינו קורא את כל תוכן הקובץ בפועל.',
      'בדיקת איכות טקסט חלופי והקשר תוכן דורשת סקירה אנושית.',
    ],
  }
}
