import { statSync } from 'fs'
import { analyzePptx } from '../src/services/pptxScanner.service.ts'
import { buildScanIssue, buildSummary, dedupeDrafts } from '../src/services/scannerShared.ts'

const file = process.argv[2]
if (!file) {
  console.error('Usage: npx tsx scripts/scan-pptx-quiet.mjs <path.pptx>')
  process.exit(1)
}

const captured = { office: null, comparison: null }
const origLog = console.log
console.log = (...args) => {
  const tag = typeof args[0] === 'string' ? args[0] : ''
  if (tag === '[OFFICE-LIKE COUNTS]') captured.office = args[1]
  if (tag === '[POWERPOINT COMPARISON SUMMARY]') captured.comparison = args[1]
}

const size = statSync(file).size
const r = analyzePptx(file, size)
const issues = dedupeDrafts(r.drafts).map(buildScanIssue)
const summary = buildSummary(issues)
origLog(
  JSON.stringify(
    { file, size, office: captured.office, comparison: captured.comparison, signals: r.signals, summary },
    null,
    2,
  ),
)
