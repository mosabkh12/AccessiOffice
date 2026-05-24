import type { FileType } from '../types/scan.types.js'
import { analyzeDocx } from './docxScanner.service.js'
import { analyzePptx } from './pptxScanner.service.js'
import { analyzeXlsx } from './xlsxScanner.service.js'
import {
  buildScanIssue,
  calculateScore,
  dedupeDrafts,
  type ExtractedSignals,
  type IssueDraft,
} from './scannerShared.js'
import type { ScanIssue } from '../types/scan.types.js'

export interface OfficeScanResult {
  issues: ScanIssue[]
  signals: ExtractedSignals
}

export async function scanOfficeFile(
  filePath: string,
  fileType: FileType,
  fileSize: number,
): Promise<OfficeScanResult> {
  let drafts: IssueDraft[] = []
  let signals: ExtractedSignals = { fileType, fileSize }

  switch (fileType) {
    case 'docx': {
      const r = analyzeDocx(filePath, fileSize)
      drafts = r.drafts
      signals = r.signals
      break
    }
    case 'pptx': {
      const r = analyzePptx(filePath, fileSize)
      drafts = r.drafts
      signals = r.signals
      break
    }
    case 'xlsx': {
      const r = await analyzeXlsx(filePath, fileSize)
      drafts = r.drafts
      signals = r.signals
      break
    }
  }

  const issues = dedupeDrafts(drafts).map(buildScanIssue)
  return { issues, signals }
}

export function logScanSummary(
  fileName: string,
  signals: ExtractedSignals,
  issueCount: number,
  score: number,
): void {
  const base = {
    fileName,
    fileType: signals.fileType,
    fileSize: signals.fileSize,
    slideCount: signals.slideCount,
    paragraphCount: signals.paragraphCount,
    sheetCount: signals.sheetCount,
    imageCount: signals.imageCount,
    tableCount: signals.tableCount,
    hyperlinkCount: signals.hyperlinkCount,
    mergedCellCount: signals.mergedCellCount,
    missingTitleSlides: signals.missingTitleSlides,
    issueCount,
    score,
  }

  if (signals.fileType === 'pptx') {
    console.log('[AccessiOffice PPTX scan]', {
      ...base,
      titleCount: signals.titleCount,
      duplicateTitleSlides: signals.duplicateTitleSlides,
      imageLikeObjectCount: signals.imageLikeObjectCount,
      missingAltCount: signals.missingAltCount,
      decorativeCandidateCount: signals.decorativeCandidateCount,
      readingOrderRiskSlides: signals.readingOrderRiskSlides,
    })
    return
  }

  console.log('[AccessiOffice scan]', base)
}

export { calculateScore, buildScanIssue, dedupeDrafts }
