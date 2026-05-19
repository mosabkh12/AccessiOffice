import type { FileType } from '../types/scan.types.js'

const FALLBACK_BY_TYPE: Record<FileType, string> = {
  docx: 'uploaded-document.docx',
  pptx: 'uploaded-presentation.pptx',
  xlsx: 'uploaded-spreadsheet.xlsx',
}

/** Detect common UTF-8 misread as Latin-1 (e.g. Hebrew showing as ×¤×¨) */
function looksCorrupted(name: string): boolean {
  if (!name) return true
  const multiplySigns = (name.match(/×/g) || []).length
  if (multiplySigns >= 2 && !/[\u0590-\u05FF]/.test(name)) return true
  if (/Ã.|Â.|Ø.|Ù.|×[§©ª«¬\u00A0-\u00BF]/.test(name)) return true
  return false
}

function tryFixLatin1AsUtf8(name: string): string {
  try {
    const fixed = Buffer.from(name, 'latin1').toString('utf8')
    if (fixed && !looksCorrupted(fixed)) return fixed
  } catch {
    /* ignore */
  }
  return name
}

function isUsableName(name: string): boolean {
  const trimmed = name?.trim()
  if (!trimmed || trimmed.length > 255) return false
  if (looksCorrupted(trimmed)) return false
  return true
}

/**
 * Resolve a safe display file name for API responses.
 * Prefers client-sent UTF-8 name, then decoded originalname, then fallback.
 */
export function resolveDisplayFileName(
  originalName: string | undefined,
  clientFileName: string | undefined,
  fileType: FileType,
): string {
  const fallback = FALLBACK_BY_TYPE[fileType]
  const candidates = [
    clientFileName?.trim(),
    originalName ? tryFixLatin1AsUtf8(originalName) : undefined,
    originalName?.trim(),
  ]

  for (const name of candidates) {
    if (name && isUsableName(name)) return name
  }

  return fallback
}
