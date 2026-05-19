import type { ExtractedSignals, IssueDraft } from './scannerShared.js'
import {
  CAT_COLOR_CONTRAST,
  CAT_DOCUMENT_ACCESS,
  CAT_DOCUMENT_STRUCTURE,
  CAT_MEDIA,
  CAT_TABLES,
  OFFICE_TITLES,
} from './officeCategories.js'
import {
  attrFromTag,
  contrastRatio,
  isMissingAltText,
  isUnclearLink,
  readZip,
  readZipEntryText,
  wordCount,
} from './ooxml.utils.js'

const PARAS_PER_PAGE_FALLBACK = 38

interface ParaInfo {
  index: number
  text: string
  style: string
  page: number
  xml: string
}

interface DocxImageInfo {
  index: number
  missingAlt: boolean
  paraIndex: number
  page: number
}

interface EmptySequence {
  startPara: number
  endPara: number
  page: number
  context: string
}

function extractParagraphs(xml: string): string[] {
  return xml.split(/<w:p[\s>]/i).slice(1).map((c) => `<w:p ${c}`)
}

function paragraphText(pXml: string): string {
  return [...pXml.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1]).join('').trim()
}

function paragraphStyle(pXml: string): string {
  return pXml.match(/<w:pStyle w:val="([^"]+)"/i)?.[1] ?? ''
}

/** Headings only from w:pStyle — never from bold/size alone. */
function headingLevelFromStyle(style: string): number | null {
  if (!style.trim()) return null
  const headingMatch = style.match(/^Heading(\d+)$/i) || style.match(/^כותרת\s*(\d+)$/i)
  if (headingMatch) return parseInt(headingMatch[1], 10)
  if (/^(Title|כותרת)$/i.test(style.trim())) return 1
  return null
}

function hasPageBreak(pXml: string): boolean {
  return (
    /<w:lastRenderedPageBreak\b/i.test(pXml) ||
    /<w:br[^>]*w:type="page"/i.test(pXml) ||
    /<w:br[^>]*type="page"/i.test(pXml)
  )
}

function buildParagraphs(docXml: string): ParaInfo[] {
  const chunks = extractParagraphs(docXml)
  let page = 1
  const paras: ParaInfo[] = []

  chunks.forEach((pXml, i) => {
    if (i > 0 && hasPageBreak(chunks[i - 1])) page++
    if (hasPageBreak(pXml)) page++
    paras.push({
      index: i + 1,
      text: paragraphText(pXml),
      style: paragraphStyle(pXml),
      page,
      xml: pXml,
    })
  })

  return paras
}

function estimatePageFallback(paraIndex: number): number {
  return Math.max(1, Math.ceil(paraIndex / PARAS_PER_PAGE_FALLBACK))
}

function truncateSnippet(text: string, max = 72): string {
  const t = text.replace(/\s+/g, ' ').trim()
  if (t.length <= max) return t
  return `${t.slice(0, max)}…`
}

function nearbyContext(paras: ParaInfo[], paraIndex: number): string {
  const idx = paraIndex - 1
  for (let d = 0; d <= 4; d++) {
    const before = paras[idx - d]
    if (before?.text && before.text.length > 1) return truncateSnippet(before.text)
    const after = paras[idx + d]
    if (after?.text && after.text.length > 1) return truncateSnippet(after.text)
  }
  return ''
}

function formatLocation(page: number, parts: string[]): string {
  return [`עמוד ${page}`, ...parts].join(' · ')
}

/** Visual-only cues (bold + large font, outline) — not w:pStyle. */
function looksVisuallyLikeHeading(pXml: string, text: string, style: string): boolean {
  if (headingLevelFromStyle(style)) return false
  if (!text || text.length < 2 || wordCount(text) > 22) return false

  const hasBold = /<w:b\b[^>]*\/>|<w:b\b[^>]*><\/w:b>|<w:b\s*\/>/i.test(pXml)
  const sizes = [...pXml.matchAll(/<w:sz[^>]*w:val="(\d+)"/gi)].map((m) => parseInt(m[1], 10))
  const maxSz = sizes.length ? Math.max(...sizes) : 0
  const hasOutline = /<w:outlineLvl\b/i.test(pXml)

  if (hasOutline) return true
  if (maxSz >= 40) return true
  if (hasBold && maxSz >= 28) return true
  return false
}

function paraIndexAtOffset(docXml: string, offset: number): number {
  let count = 0
  for (const m of docXml.matchAll(/<w:p[\s>]/gi)) {
    if ((m.index ?? 0) < offset) count++
    else break
  }
  return Math.max(1, count)
}

function extractDocxImages(docXml: string, paras: ParaInfo[]): DocxImageInfo[] {
  const images: DocxImageInfo[] = []
  const docPrTags = [
    ...docXml.matchAll(/<(?:wp:)?docPr\b[^>]*\/?>/gi),
    ...docXml.matchAll(/<(?:wp:)?docPr\b[\s\S]*?<\/(?:wp:)?docPr>/gi),
  ]
  const seen = new Set<number>()

  for (const m of docPrTags) {
    const tag = m[0]
    const offset = m.index ?? 0
    if (seen.has(offset)) continue
    seen.add(offset)

    const descr = attrFromTag(tag, 'descr')
    const title = attrFromTag(tag, 'title')
    const name = attrFromTag(tag, 'name')
    const paraIndex = paraIndexAtOffset(docXml, offset)
    const para = paras[paraIndex - 1]
    const page = para?.page ?? estimatePageFallback(paraIndex)

    images.push({
      index: images.length + 1,
      missingAlt: isMissingAltText(descr, title, name),
      paraIndex,
      page,
    })
  }

  if (images.length === 0) {
    const drawingChunks = [
      ...(docXml.matchAll(/<w:drawing[\s\S]*?<\/w:drawing>/gi) ?? []),
      ...(docXml.matchAll(/<pic:pic[\s\S]*?<\/pic:pic>/gi) ?? []),
    ]
    drawingChunks.forEach((m) => {
      const chunk = m[0]
      const offset = m.index ?? 0
      const descr = chunk.match(/descr="([^"]*)"/i)?.[1]
      const titleAttr = chunk.match(/title="([^"]*)"/i)?.[1]
      const name = chunk.match(/name="([^"]*)"/i)?.[1]
      const paraIndex = paraIndexAtOffset(docXml, offset)
      const para = paras[paraIndex - 1]
      images.push({
        index: images.length + 1,
        missingAlt: isMissingAltText(descr, titleAttr, name),
        paraIndex,
        page: para?.page ?? estimatePageFallback(paraIndex),
      })
    })
  }

  return images
}

function findEmptySequences(paras: ParaInfo[], minRun = 3): EmptySequence[] {
  const sequences: EmptySequence[] = []
  let runStart = -1

  paras.forEach((p, i) => {
    const isEmpty = !p.text || p.text.length < 2
    if (isEmpty && runStart < 0) runStart = i
    if (!isEmpty && runStart >= 0) {
      const len = i - runStart
      if (len >= minRun) {
        const startPara = paras[runStart]
        sequences.push({
          startPara: startPara.index,
          endPara: paras[i - 1].index,
          page: startPara.page,
          context: nearbyContext(paras, startPara.index),
        })
      }
      runStart = -1
    }
  })

  if (runStart >= 0 && paras.length - runStart >= minRun) {
    const startPara = paras[runStart]
    sequences.push({
      startPara: startPara.index,
      endPara: paras[paras.length - 1].index,
      page: startPara.page,
      context: nearbyContext(paras, startPara.index),
    })
  }

  return sequences
}

function detectRestrictedAccess(zip: ReturnType<typeof readZip>, drafts: IssueDraft[]): void {
  const settings = readZipEntryText(zip, 'word/settings.xml')
  if (!settings) return

  const hasProtection =
    /<w:documentProtection\b/i.test(settings) ||
    /<w:writeProtection\b/i.test(settings) ||
    /w:enforcement="1"/i.test(settings)

  if (!hasProtection) return

  const editAttr = settings.match(/<w:documentProtection[^>]*w:edit="([^"]*)"/i)?.[1]
  drafts.push({
    id: 'restricted-access',
    wcagKey: 'info-relationships',
    title: OFFICE_TITLES.restrictedAccess,
    severity: 'High',
    category: CAT_DOCUMENT_ACCESS,
    affectedUsers: ['כל המשתמשים', 'משתמשי קורא מסך'],
    impact: `המסמך מוגן${editAttr ? ` (מצב: ${editAttr})` : ''}. גישה מוגבלת עלולה למנוע שימוש בטכנולוגיות מסייעות.`,
    recommendation: 'ב-Word: קובץ → מידע על המסמך → הגנת מסמך — הסירו הגנה מיותרת לפני פרסום.',
    location: 'הגדרות המסמך (word/settings.xml)',
    confidence: 'High',
  })
}

function detectWordTableMerges(
  docXml: string,
  paras: ParaInfo[],
  drafts: IssueDraft[],
): void {
  const tables = docXml.split(/<w:tbl[\s>]/i).slice(1)
  tables.forEach((chunk, i) => {
    const tableXml = `<w:tbl ${chunk}`
    const hasMerge =
      /<w:gridSpan\b[^>]*w:val="(?!1")/i.test(tableXml) ||
      /<w:vMerge\b/i.test(tableXml) ||
      /<w:hMerge\b/i.test(tableXml)
    if (!hasMerge) return

    const tableOffset = docXml.indexOf(chunk)
    const paraIdx = paraIndexAtOffset(docXml, tableOffset)
    const page = paras[paraIdx - 1]?.page ?? estimatePageFallback(paraIdx)
    drafts.push({
      id: `table-merged-${i + 1}`,
      wcagKey: 'info-relationships',
      title: OFFICE_TITLES.mergedOrSplitCells,
      severity: 'High',
      category: CAT_TABLES,
      affectedUsers: ['משתמשי קורא מסך'],
      impact: `בטבלה ${i + 1} זוהו תאים ממוזגים או מפוצלים (gridSpan/vMerge/hMerge).`,
      recommendation: 'בטלו מיזוג/פיצול תאים והשתמשו בכותרות טבלה תקינות.',
      location: formatLocation(page, [`טבלה ${i + 1}`, `פסקה ${paraIdx}`]),
      confidence: 'High',
    })
  })
}

function detectDocxContrast(docXml: string, paras: ParaInfo[], drafts: IssueDraft[]): void {
  let failures = 0
  let firstPara = 1
  let firstPage = 1

  for (const m of docXml.matchAll(/<w:r[\s\S]*?<\/w:r>/gi)) {
    const run = m[0]
    const fg = run.match(/<w:color[^>]*w:val="([A-Fa-f0-9]{6})"/i)?.[1]
    const shd = run.match(/<w:shd[^>]*w:fill="([A-Fa-f0-9]{6})"/i)?.[1]
    if (fg && shd) {
      const ratio = contrastRatio(fg, shd)
      if (ratio != null && ratio < 4.5) {
        failures++
        const offset = m.index ?? 0
        firstPara = paraIndexAtOffset(docXml, offset)
        firstPage = paras[firstPara - 1]?.page ?? estimatePageFallback(firstPara)
      }
    }
  }

  if (failures >= 1) {
    drafts.push({
      id: 'low-contrast-docx',
      wcagKey: 'contrast-minimum',
      title: OFFICE_TITLES.hardTextContrast,
      severity: 'High',
      category: CAT_COLOR_CONTRAST,
      affectedUsers: ['משתמשים עם לקות ראייה'],
      impact: `זוהו ${failures} מופעי טקסט עם ניגודיות נמוכה מ-4.5:1.`,
      recommendation: 'שפרו ניגודיות בין צבע גופן לרקע (מילוי פסקה).',
      location: formatLocation(firstPage, [`פסקה ${firstPara}`]),
      confidence: 'Low',
    })
  }
}

export function analyzeDocx(
  filePath: string,
  fileSize: number,
): { drafts: IssueDraft[]; signals: ExtractedSignals } {
  const zip = readZip(filePath)
  const docXml = readZipEntryText(zip, 'word/document.xml')
  const drafts: IssueDraft[] = []
  const signals: ExtractedSignals = {
    fileType: 'docx',
    fileSize,
    paragraphCount: 0,
    imageCount: 0,
    tableCount: 0,
    hyperlinkCount: 0,
    emptyParagraphCount: 0,
  }

  if (!docXml) return { drafts, signals }

  const paras = buildParagraphs(docXml)
  signals.paragraphCount = paras.length

  const headingOccurrences: { level: number; paraIndex: number; page: number }[] = []
  const visualHeadingParas: ParaInfo[] = []
  let substantiveParas = 0
  let emptyCount = 0

  paras.forEach((p) => {
    const lvl = headingLevelFromStyle(p.style)
    if (lvl) headingOccurrences.push({ level: lvl, paraIndex: p.index, page: p.page })

    if (looksVisuallyLikeHeading(p.xml, p.text, p.style)) {
      visualHeadingParas.push(p)
    }

    if (!p.text || p.text.length < 2) {
      emptyCount++
    } else {
      substantiveParas++
    }

    const words = wordCount(p.text)
    if (words > 120) {
      const ctx = nearbyContext(paras, p.index)
      drafts.push({
        id: `long-paragraph-${p.index}`,
        wcagKey: 'reading-level',
        title: OFFICE_TITLES.longParagraph,
        severity: 'Low',
        category: CAT_DOCUMENT_STRUCTURE,
        affectedUsers: ['משתמשים עם לקויות קוגניטיביות', 'משתמשי קורא מסך'],
        impact: `פסקה ${p.index} מכילה כ-${words} מילים.${ctx ? ` טקסט סמוך: "${ctx}".` : ''}`,
        recommendation: 'פצלו לפסקאות קצרות עם כותרות משנה.',
        location: formatLocation(p.page, [`פסקה ${p.index}`]),
        confidence: 'High',
      })
    }
  })
  signals.emptyParagraphCount = emptyCount

  visualHeadingParas.forEach((p) => {
    const ctx = nearbyContext(paras, p.index)
    const snippet = truncateSnippet(p.text)
    drafts.push({
      id: `visual-heading-p${p.index}`,
      wcagKey: 'headings-labels',
      title: OFFICE_TITLES.visualHeadingNoStyle,
      severity: 'Medium',
      category: CAT_DOCUMENT_STRUCTURE,
      affectedUsers: ['משתמשי קורא מסך', 'משתמשים עם לקויות קוגניטיביות'],
      impact: `פסקה ${p.index} מעוצבת ככותרת (גופן גדול/מודגש) אך ללא סגנון Heading.${snippet ? ` טקסט: "${snippet}".` : ''}${ctx && ctx !== snippet ? ` הקשר: "${ctx}".` : ''}`,
      recommendation: 'החילו סגנון כותרת (Heading 1–3) במקום עיצוב ידני.',
      location: formatLocation(p.page, [`פסקה ${p.index}`, snippet ? `ליד: "${snippet}"` : ''].filter(Boolean)),
      confidence: 'High',
    })
  })

  const images = extractDocxImages(docXml, paras)
  signals.imageCount = images.length
  images.forEach((img) => {
    if (!img.missingAlt) return
    const ctx = nearbyContext(paras, img.paraIndex)
    drafts.push({
      id: `missing-alt-text-${img.index}`,
      wcagKey: 'non-text-content',
      title: OFFICE_TITLES.missingAltText,
      severity: 'High',
      category: CAT_MEDIA,
      affectedUsers: ['משתמשים עיוורים', 'משתמשי קורא מסך'],
      impact: `תמונה ${img.index} ללא טקסט חלופי (wp:docPr descr/title).${ctx ? ` טקסט סמוך: "${ctx}".` : ''}`,
      recommendation: 'ב-Word: לחצו ימני על התמונה → תמונה → טקסט חלופי.',
      location: formatLocation(img.page, [`תמונה ${img.index}`, `פסקה ${img.paraIndex}`]),
      confidence: 'High',
    })
  })

  const headingLevels = headingOccurrences.map((h) => h.level)

  if (substantiveParas >= 5 && headingLevels.length === 0) {
    const sampleParas = visualHeadingParas.slice(0, 3).map((p) => p.index)
    const firstPage = visualHeadingParas[0]?.page ?? paras[0]?.page ?? 1
    drafts.push({
      id: 'incorrect-heading-structure',
      wcagKey: 'headings-labels',
      title: OFFICE_TITLES.noHeadingsInDocument,
      severity: 'High',
      category: CAT_DOCUMENT_STRUCTURE,
      affectedUsers: ['משתמשי קורא מסך', 'משתמשים עם לקויות קוגניטיביות'],
      impact: `זוהו ${substantiveParas} פסקאות עם תוכן אך אין אף סגנון כותרת (w:pStyle Heading).${visualHeadingParas.length ? ` זוהו ${visualHeadingParas.length} פסקאות שנראות ככותרת ללא סגנון.` : ''}`,
      recommendation: 'החילו סגנונות כותרת מובנים (Heading 1–3) ב-Word.',
      location:
        sampleParas.length > 0
          ? formatLocation(firstPage, [`פסקאות ${sampleParas.join(', ')}`])
          : formatLocation(paras[0]?.page ?? 1, ['ללא סגנונות כותרת במסמך']),
      confidence: 'Medium',
    })
  } else if (headingOccurrences.length >= 2) {
    for (let i = 1; i < headingOccurrences.length; i++) {
      const prev = headingOccurrences[i - 1]
      const cur = headingOccurrences[i]
      if (cur.level - prev.level > 1) {
        drafts.push({
          id: `heading-skip-${cur.paraIndex}`,
          wcagKey: 'info-relationships',
          title: OFFICE_TITLES.headingLevelSkip,
          severity: 'High',
          category: CAT_DOCUMENT_STRUCTURE,
          affectedUsers: ['משתמשי קורא מסך'],
          impact: `דילוג ברמת כותרת מ-${prev.level} ל-${cur.level} בפסקה ${cur.paraIndex}.`,
          recommendation: 'הימנעו מדילוג ברמות כותרת.',
          location: formatLocation(cur.page, [`פסקה ${cur.paraIndex}`, `כותרת ${cur.level}`]),
          confidence: 'High',
        })
        break
      }
    }
  }

  const tableCount = (docXml.match(/<w:tbl[\s>]/gi) ?? []).length
  signals.tableCount = tableCount
  if (tableCount > 0) {
    const tables = docXml.split(/<w:tbl[\s>]/i).slice(1)
    tables.forEach((chunk, i) => {
      const tableXml = `<w:tbl ${chunk}`
      const firstRow = tableXml.match(/<w:tr[\s\S]*?<\/w:tr>/i)?.[0] ?? ''
      const hasTblHeader = /<w:tblHeader\s*\/>/i.test(firstRow)
      const cells = [...firstRow.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1].trim())
      const hasContent = cells.some((c) => c.length > 0)
      const looksLikeHeader =
        hasTblHeader || (cells.length >= 2 && cells.every((c) => c.length > 0 && c.length < 60))
      if (hasContent && !looksLikeHeader) {
        const tableOffset = docXml.indexOf(chunk)
        const paraIdx = paraIndexAtOffset(docXml, tableOffset)
        const page = paras[paraIdx - 1]?.page ?? estimatePageFallback(paraIdx)
        drafts.push({
          id: `table-header-${i + 1}`,
          wcagKey: 'info-relationships',
          title: OFFICE_TITLES.missingTableHeader,
          severity: 'Medium',
          category: CAT_TABLES,
          affectedUsers: ['משתמשי קורא מסך'],
          impact: `בטבלה ${i + 1} ליד פסקה ${paraIdx} לא זוהתה שורת כותרת.`,
          recommendation: 'סמנו שורת כותרת בטבלה ב-Word.',
          location: formatLocation(page, [`טבלה ${i + 1}`, `פסקה ${paraIdx}`]),
          confidence: 'High',
        })
      }
    })
  }

  let hyperlinkCount = 0
  paras.forEach((p) => {
    if (!/<w:hyperlink/i.test(p.xml)) return
    hyperlinkCount++
    if (isUnclearLink(p.text)) {
      const ctx = nearbyContext(paras, p.index)
      drafts.push({
        id: `unclear-link-p${p.index}`,
        wcagKey: 'link-purpose',
        title: OFFICE_TITLES.unclearHyperlink,
        severity: 'Medium',
        category: CAT_DOCUMENT_STRUCTURE,
        affectedUsers: ['משתמשי קורא מסך'],
        impact: `בפסקה ${p.index} זוהה קישור: "${p.text}".${ctx ? ` הקשר: "${ctx}".` : ''}`,
        recommendation: 'השתמשו בטקסט קישור תיאורי.',
        location: formatLocation(p.page, [`פסקה ${p.index}`]),
        confidence: 'High',
      })
    }
  })
  signals.hyperlinkCount = hyperlinkCount

  const emptySequences = findEmptySequences(paras)
  emptySequences.forEach((seq, i) => {
    const count = seq.endPara - seq.startPara + 1
    drafts.push({
      id: `empty-paragraphs-${seq.startPara}-${seq.endPara}`,
      wcagKey: 'info-relationships',
      title: OFFICE_TITLES.emptyParagraphSpacing,
      severity: 'Medium',
      category: CAT_DOCUMENT_STRUCTURE,
      affectedUsers: ['משתמשי קורא מסך'],
      impact: `רצף של ${count} פסקאות ריקות (פסקאות ${seq.startPara}–${seq.endPara}).${seq.context ? ` טקסט סמוך: "${seq.context}".` : ''}`,
      recommendation: 'השתמשו במרווח לפני/אחרי פסקה במקום פסקאות ריקות.',
      location: formatLocation(seq.page, [
        `פסקאות ${seq.startPara}–${seq.endPara}`,
        seq.context ? `ליד: "${seq.context}"` : '',
      ].filter(Boolean)),
      confidence: 'High',
    })
  })

  if (emptySequences.length === 0 && emptyCount >= 8 && paras.length >= 10) {
    drafts.push({
      id: 'empty-paragraphs-scattered',
      wcagKey: 'info-relationships',
      title: OFFICE_TITLES.emptyParagraphSpacing,
      severity: 'Medium',
      category: CAT_DOCUMENT_STRUCTURE,
      affectedUsers: ['משתמשי קורא מסך'],
      impact: `זוהו ${emptyCount} פסקאות ריקות מפוזרות במסמך.`,
      recommendation: 'השתמשו במרווח לפני/אחרי פסקה.',
      location: formatLocation(paras[0]?.page ?? 1, [`${emptyCount} פסקאות ריקות`]),
      confidence: 'High',
    })
  }

  detectWordTableMerges(docXml, paras, drafts)
  detectDocxContrast(docXml, paras, drafts)
  detectRestrictedAccess(zip, drafts)

  return { drafts, signals }
}
