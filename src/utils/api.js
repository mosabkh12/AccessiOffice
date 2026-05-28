const API_URL = import.meta.env.VITE_API_URL || 'http://localhost:4000'
const DEBUG_SCAN = import.meta.env.VITE_ACCESSIOFFICE_DEBUG_SCAN === 'true'

const USER_TYPE_TO_API = {
  author: 'document-author',
  auditor: 'accessibility-auditor',
  lecturer: 'lecturer-institution',
}

const USER_TYPE_FROM_API = Object.fromEntries(
  Object.entries(USER_TYPE_TO_API).map(([k, v]) => [v, k]),
)

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
  // occurrenceCount: prefer explicit field, fall back to locations array length, then 1
  const locations = Array.isArray(raw?.locations) ? raw.locations : null
  const occurrenceCount =
    typeof raw?.occurrenceCount === 'number'
      ? raw.occurrenceCount
      : locations
        ? locations.length
        : 1

  return {
    id: raw?.id || `issue-${Math.random().toString(36).slice(2, 9)}`,
    wcagKey: raw?.wcagKey,
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
    occurrenceCount,
    locations,
    confidence: raw?.confidence,
  }
}

export function normalizeScanResponse(api, clientFileName) {
  const userType = USER_TYPE_FROM_API[api.userType] ?? api.userType
  const fileType = api.fileType || getFileType(clientFileName) || 'docx'
  const fileName = safeFileName(api.fileName || clientFileName, fileType)

  const issues = (api.issues || []).map(normalizeIssue)
  const quickFix = (api.quickFix || []).map(normalizeIssue)
  const summary = api.summary || {}
  const scanDiagnostics = api.scanDiagnostics || api.diagnostics || {}
  const checkStatuses = Array.isArray(api.checkStatuses) ? api.checkStatuses : []
  const manualReviewChecks = Array.isArray(api.manualReviewChecks)
    ? api.manualReviewChecks
    : checkStatuses.filter((c) => ['manual', 'partial', 'not_checked'].includes(c?.status))
  const officeLikeSummary    = api.officeLikeSummary    ?? null
  const fileHash             = api.fileHash             ?? null
  const engine               = api.engine               ?? 'xml'
  const workerError          = api.workerError          ?? null
  const workerScannerVersion = api.workerScannerVersion ?? null
  const powerpointWorkerRaw  = api.powerpointWorkerRaw  ?? null

  // Resolve occurrence-aware counts with backward-compatible fallbacks
  const totalOccurrences  = summary.totalOccurrences  ?? issues.reduce((s, i) => s + (i.occurrenceCount ?? 1), 0)
  const totalIssueTypes   = summary.totalIssueTypes   ?? issues.length
  const highOccurrences   = summary.highOccurrences   ?? issues.filter(i => i.severity === 'high').reduce((s, i) => s + (i.occurrenceCount ?? 1), 0)
  const mediumOccurrences = summary.mediumOccurrences ?? issues.filter(i => i.severity === 'medium').reduce((s, i) => s + (i.occurrenceCount ?? 1), 0)
  const lowOccurrences    = summary.lowOccurrences    ?? issues.filter(i => i.severity === 'low').reduce((s, i) => s + (i.occurrenceCount ?? 1), 0)
  const quickFixOccurrences = summary.quickFixOccurrences ?? quickFix.reduce((s, i) => s + (i.occurrenceCount ?? 1), 0)

  // Summary text using occurrence counts
  let summaryText
  if (totalOccurrences === 0) {
    summaryText = 'לא נמצאו בעיות נגישות משמעותיות בסריקה האוטומטית. מומלץ לבצע בדיקה ידנית לפני פרסום.'
  } else {
    summaryText = `הקובץ כולל ${totalOccurrences} ממצאי נגישות ב-${totalIssueTypes} סוגי בעיות. מומלץ לתקן לפני פרסום.`
    if (quickFixOccurrences > 0) {
      summaryText += ` בנוסף נמצאו ${quickFixOccurrences} הצעות Quick Fix לאובייקטים דקורטיביים.`
    }
  }

  const criticalAnalysis =
    api.criticalAnalysis ||
    'סריקה אוטומטית מסייעת לזהות בעיות נפוצות, אך אינה מחליפה בדיקה אנושית מקצועית.'

  // ── Debug: single source of truth for all UI counts ─────────────────────
  const _byTitle = [
    ...issues.map(i => ({
      title: i.title,
      occurrenceCount: i.occurrenceCount,
      severity: i.severity,
      locationsCount: i.locations?.length ?? 0,
      isQuickFix: false,
    })),
    ...quickFix.map(i => ({
      title: i.title,
      occurrenceCount: i.occurrenceCount,
      severity: i.severity,
      locationsCount: i.locations?.length ?? 0,
      isQuickFix: true,
    })),
  ]
  if (DEBUG_SCAN) {
    console.log('[SCAN RESPONSE] Data sources:', {
      fileHash,
      engine,
      workerError,
      workerScannerVersion,
      sectionA_source: officeLikeSummary ? `officeLikeSummary (engine: ${engine})` : 'AssistantPanel standard rendering',
      sectionB_source: officeLikeSummary ? `issues[${issues.length}] + quickFix[${quickFix.length}] (detailed WCAG findings)` : 'N/A',
      officeLikeSummary: officeLikeSummary
        ? Object.entries(officeLikeSummary).map(([k, v]) => `${k}: ${v.status} (${v.count})`)
        : null,
    })
    console.log('[NORMALIZED ISSUE COUNTS]', {
      totalIssueTypes,
      totalOccurrences,
      highOccurrences,
      mediumOccurrences,
      lowOccurrences,
      quickFixOccurrences,
      byTitle: _byTitle,
    })
    console.log('[DASHBOARD COUNTS]', { totalOccurrences, totalIssueTypes, highOccurrences, mediumOccurrences, quickFixOccurrences })
    console.log('[ASSISTANT PANEL COUNTS]', _byTitle.map(i => `${i.title}: ${i.occurrenceCount}${i.isQuickFix ? ' (QF)' : ''}`))
    console.log('[RESULTS TABLE COUNTS]', issues.map(i => `${i.title}: ${i.occurrenceCount}`))
    console.log('[REPORT COUNTS]', { totalOccurrences, issues: issues.length, quickFix: quickFix.length })
    console.log('[SCAN DIAGNOSTICS]', {
      receivedScannerVersion: api.scannerVersion,
      normalizedScannerVersion: api.scannerVersion,
      scanDiagnostics,
      checkStatuses,
      manualReviewChecks,
    })
  }

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
    // Occurrence-aware counts (what the dashboard shows)
    totalOccurrences,
    totalIssueTypes,
    highOccurrences,
    mediumOccurrences,
    lowOccurrences,
    quickFixOccurrences,
    quickFixCount: quickFix.length,
    // Backward-compatible aliases
    totalIssues: totalOccurrences,
    severityCounts: { high: highOccurrences, medium: mediumOccurrences, low: lowOccurrences },
    summary: summaryText,
    issues,
    quickFix,
    scannerVersion: api.scannerVersion,
    scanDiagnostics,
    diagnostics: scanDiagnostics,
    checkStatuses,
    manualReviewChecks,
    officeLikeSummary,
    fileHash,
    engine,
    workerError,
    workerScannerVersion,
    powerpointWorkerRaw,
    recommendations: [...new Set(issues.map((i) => i.recommendation).filter(Boolean))],
    criticalAnalysis,
    limitations: [
      criticalAnalysis,
      'הכלי מנתח מבנה כללי ואינו קורא את כל תוכן הקובץ בפועל.',
      'בדיקת איכות טקסט חלופי והקשר תוכן דורשת סקירה אנושית.',
    ],
  }
}
