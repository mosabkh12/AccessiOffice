import type { ExtractedSignals, IssueDraft } from './issue.builder.js'
import {
  collectText,
  contrastRatio,
  isUnclearLink,
  listZipEntries,
  parseXml,
  readZip,
  readZipEntryText,
} from './ooxml.utils.js'

function slideHasTitle(xml: string, parsed: unknown): boolean {
  if (/type="(?:title|ctrTitle)"/i.test(xml)) {
    const chunk = xml.match(/type="(?:title|ctrTitle)"[\s\S]{0,2000}/i)?.[0]
    if (chunk) {
      const textInChunk = [...chunk.matchAll(/<a:t>([^<]*)<\/a:t>/g)]
        .map((m) => m[1])
        .join('')
        .trim()
      if (textInChunk.length > 2) return true
    }
  }
  const tree = parsed as Record<string, unknown>
  const sld = tree?.sld as Record<string, unknown> | undefined
  const spTree = (sld?.cSld as Record<string, unknown> | undefined)?.spTree
  const shapes = (spTree as Record<string, unknown> | undefined)?.sp
  const list = Array.isArray(shapes) ? shapes : shapes ? [shapes] : []
  for (const sp of list) {
    const spObj = sp as Record<string, unknown>
    const ph = (spObj.nvSpPr as Record<string, unknown> | undefined)?.nvPr as Record<string, unknown> | undefined
    const phType = String((ph?.ph as Record<string, unknown> | undefined)?.['@_type'] ?? '')
    if (phType.includes('title') || phType === 'ctrTitle') {
      if (collectText(spObj.txBody).trim().length > 2) return true
    }
  }
  return false
}

function safeParseXml(xml: string): unknown {
  try {
    return parseXml(xml)
  } catch {
    return null
  }
}

function analyzeSlide(slideNum: number, xml: string) {
  const parsed = safeParseXml(xml)
  const picChunks = [...xml.matchAll(/<pic[\s\S]*?<\/pic>/gi)].map((m) => m[0])
  let imagesMissingAlt = 0
  for (const picChunk of picChunks) {
    const descr = picChunk.match(/descr="([^"]*)"/i)?.[1]?.trim()
    if (!descr || descr.length < 2) imagesMissingAlt++
  }
  const imageCount = picChunks.length
  const textShapes = (xml.match(/<sp[\s>]/gi) ?? []).length
  const unclearLinks: string[] = []
  for (const m of xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)) {
    if (isUnclearLink(m[1])) unclearLinks.push(m[1].trim())
  }
  let smallFontRuns = 0
  for (const sm of xml.matchAll(/sz="(\d+)"/g)) {
    const sz = parseInt(sm[1], 10)
    if (sz > 0 && sz < 1400) smallFontRuns++
  }
  let lowContrastHints = 0
  const fills = [...xml.matchAll(/srgbClr val="([A-Fa-f0-9]{6})"/g)].map((m) => m[1])
  for (let i = 0; i < fills.length - 1; i++) {
    const ratio = contrastRatio(fills[i], fills[i + 1])
    if (ratio != null && ratio < 4.5) lowContrastHints++
  }
  return {
    slideNum,
    hasTitle: slideHasTitle(xml, parsed),
    imageCount,
    imagesMissingAlt,
    textShapes,
    unclearLinks: [...new Set(unclearLinks)],
    smallFontRuns,
    lowContrastHints,
  }
}

export function analyzePptx(filePath: string, fileSize: number): { drafts: IssueDraft[]; signals: ExtractedSignals } {
  const zip = readZip(filePath)
  const slidePaths = listZipEntries(zip, /^ppt\/slides\/slide\d+\.xml$/i)
  const drafts: IssueDraft[] = []
  const missingTitleSlides: number[] = []
  let totalImages = 0

  const signals: ExtractedSignals = {
    fileType: 'pptx',
    fileSize,
    slideCount: slidePaths.length,
    imageCount: 0,
    missingTitleSlides: [],
  }

  for (const slidePath of slidePaths) {
    const xml = readZipEntryText(zip, slidePath)
    if (!xml) continue
    const slideNum = parseInt(slidePath.match(/slide(\d+)\.xml/i)?.[1] ?? '0', 10)
    const loc = `שקופית ${slideNum}`
    const s = analyzeSlide(slideNum, xml)
    totalImages += s.imageCount

    if (!s.hasTitle) {
      missingTitleSlides.push(slideNum)
      drafts.push({
        id: `slide-without-title-${slideNum}`,
        wcagKey: 'page-titled-headings',
        title: 'שקופית ללא כותרת',
        severity: 'Medium',
        category: 'מבנה',
        affectedUsers: ['משתמשי קורא מסך'],
        impact: `בשקופית ${slideNum} לא זוהתה כותרת.`,
        recommendation: 'הוסיפו כותרת ייחודית לשקופית.',
        location: loc,
      })
    }

    if (s.imageCount > 0 && s.imagesMissingAlt > 0) {
      drafts.push({
        id: `image-without-alt-${slideNum}`,
        wcagKey: 'non-text-content',
        title: 'תמונה ללא טקסט חלופי',
        severity: 'High',
        category: 'תוכן',
        affectedUsers: ['משתמשים עיוורים', 'משתמשי קורא מסך'],
        impact: `בשקופית ${slideNum}: ${s.imagesMissingAlt} תמונות ללא טקסט חלופי.`,
        recommendation: 'הוסיפו תיאור alt לכל תמונה.',
        location: loc,
      })
    }

    if (s.textShapes >= 5) {
      drafts.push({
        id: `reading-order-${slideNum}`,
        wcagKey: 'meaningful-sequence',
        title: 'בעיית סדר קריאה',
        severity: 'Medium',
        category: 'מבנה',
        affectedUsers: ['משתמשי קורא מסך'],
        impact: `בשקופית ${slideNum} זוהו ${s.textShapes} אזורי טקסט – סיכון לסדר קריאה שגוי.`,
        recommendation: 'בדקו סדר קריאה ב-PowerPoint.',
        location: loc,
      })
    }

    for (const link of s.unclearLinks) {
      drafts.push({
        id: `unclear-link-${slideNum}-${link.slice(0, 6)}`,
        wcagKey: 'link-purpose',
        title: 'קישור לא ברור',
        severity: 'Medium',
        category: 'ניווט',
        affectedUsers: ['משתמשי קורא מסך'],
        impact: `בשקופית ${slideNum}: "${link}".`,
        recommendation: 'החליפו בטקסט קישור תיאורי.',
        location: loc,
      })
    }

    if (s.smallFontRuns >= 2) {
      drafts.push({
        id: `small-font-${slideNum}`,
        wcagKey: 'resize-text',
        title: 'גודל גופן קטן מדי',
        severity: 'Low',
        category: 'תצוגה',
        affectedUsers: ['משתמשים עם לקות ראייה'],
        impact: `בשקופית ${slideNum} זוהה גופן קטן (${s.smallFontRuns} מופעים).`,
        recommendation: 'הגדילו גופן ל-18pt לפחות.',
        location: loc,
      })
    }

    if (s.lowContrastHints > 0) {
      drafts.push({
        id: `low-contrast-${slideNum}`,
        wcagKey: 'contrast-minimum',
        title: 'ניגודיות צבעים נמוכה',
        severity: 'High',
        category: 'תצוגה',
        affectedUsers: ['משתמשים עם לקות ראייה'],
        impact: `בשקופית ${slideNum} זוהו צירופי צבע בעלי ניגודיות חלשה אפשרית.`,
        recommendation: 'שפרו ניגודיות טקסט/רקע.',
        location: loc,
      })
    }
  }

  signals.imageCount = totalImages
  signals.missingTitleSlides = missingTitleSlides
  return { drafts, signals }
}
