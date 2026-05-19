import type { ExtractedSignals, IssueDraft } from './issue.builder.js'
import { isUnclearLink, readZip, readZipEntryText } from './ooxml.utils.js'

function extractParagraphs(xml: string): string[] {
  return xml.split(/<w:p[\s>]/i).slice(1).map((c) => `<w:p ${c}`)
}

function paragraphText(pXml: string): string {
  return [...pXml.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1]).join('').trim()
}

function paragraphStyle(pXml: string): string {
  return pXml.match(/<w:pStyle w:val="([^"]+)"/i)?.[1] ?? ''
}

function headingLevel(style: string): number | null {
  const m = style.match(/Heading(\d+)/i) || style.match(/כותרת\s*(\d+)/i)
  return m ? parseInt(m[1], 10) : null
}

function countImages(xml: string): { total: number; missingAlt: number } {
  const chunks = [
    ...(xml.match(/<w:drawing[\s\S]*?<\/w:drawing>/gi) ?? []),
    ...(xml.match(/<wp:inline[\s\S]*?<\/wp:inline>/gi) ?? []),
    ...(xml.match(/<pic:pic[\s\S]*?<\/pic:pic>/gi) ?? []),
  ]
  let missingAlt = 0
  for (const chunk of chunks) {
    const descr = chunk.match(/descr="([^"]*)"/i)?.[1]?.trim()
    const title = chunk.match(/title="([^"]*)"/i)?.[1]?.trim()
    if (!descr && !title) missingAlt++
  }
  return { total: chunks.length, missingAlt }
}

export function analyzeDocx(filePath: string, fileSize: number): { drafts: IssueDraft[]; signals: ExtractedSignals } {
  const zip = readZip(filePath)
  const docXml = readZipEntryText(zip, 'word/document.xml')
  const drafts: IssueDraft[] = []
  const signals: ExtractedSignals = { fileType: 'docx', fileSize, paragraphCount: 0, imageCount: 0, tableCount: 0, hyperlinkCount: 0, emptyParagraphCount: 0 }

  if (!docXml) return { drafts, signals }

  const paras = extractParagraphs(docXml)
  signals.paragraphCount = paras.length

  const headingLevels: number[] = []
  let emptyCount = 0
  let substantiveParas = 0

  paras.forEach((p, idx) => {
    const text = paragraphText(p)
    const style = paragraphStyle(p)
    const lvl = headingLevel(style)
    if (lvl) headingLevels.push(lvl)
    if (!text || text.length < 2) emptyCount++
    else substantiveParas++

    if (text.length > 900) {
      drafts.push({
        id: `long-paragraph-${idx + 1}`,
        wcagKey: 'reading-level',
        title: 'פסקאות ארוכות וקשות לקריאה',
        severity: 'Low',
        category: 'קריאות',
        affectedUsers: ['משתמשים עם לקויות קוגניטיביות', 'משתמשי קורא מסך'],
        impact: `פסקה ${idx + 1} ארוכה מאוד (כ-${text.length} תווים).`,
        recommendation: 'פצלו לפסקאות קצרות עם כותרות משנה.',
        location: `פסקה ${idx + 1}`,
      })
    }
  })
  signals.emptyParagraphCount = emptyCount

  const { total: imageCount, missingAlt } = countImages(docXml)
  signals.imageCount = imageCount
  if (imageCount > 0 && missingAlt > 0) {
    drafts.push({
      id: 'missing-alt-text',
      wcagKey: 'non-text-content',
      title: 'תמונה ללא טקסט חלופי',
      severity: 'High',
      category: 'תוכן',
      affectedUsers: ['משתמשים עיוורים', 'משתמשי קורא מסך'],
      impact: `זוהו ${missingAlt} מתוך ${imageCount} תמונות ללא תיאור חלופי.`,
      recommendation: 'הוסיפו טקסט חלופי משמעותי לכל תמונה.',
      location: missingAlt === 1 ? 'תמונה 1' : `${missingAlt} תמונות במסמך`,
    })
  }

  if (substantiveParas >= 5 && headingLevels.length === 0) {
    drafts.push({
      id: 'incorrect-heading-structure',
      wcagKey: 'headings-labels',
      title: 'מבנה כותרות לא תקין',
      severity: 'High',
      category: 'מבנה',
      affectedUsers: ['משתמשי קורא מסך', 'משתמשים עם לקויות קוגניטיביות'],
      impact: `זוהו ${substantiveParas} פסקאות עם תוכן אך ללא סגנונות כותרת.`,
      recommendation: 'הגדירו כותרות 1–3 בסדר היררכי.',
      location: 'מבנה המסמך',
    })
  } else if (headingLevels.length >= 2) {
    for (let i = 1; i < headingLevels.length; i++) {
      if (headingLevels[i] - headingLevels[i - 1] > 1) {
        drafts.push({
          id: 'heading-skip',
          wcagKey: 'info-relationships',
          title: 'מבנה כותרות לא תקין',
          severity: 'High',
          category: 'מבנה',
          affectedUsers: ['משתמשי קורא מסך'],
          impact: `דילוג ברמת כותרת מ-${headingLevels[i - 1]} ל-${headingLevels[i]}.`,
          recommendation: 'הימנעו מדילוג ברמות כותרת.',
          location: `כותרת ${i + 1}`,
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
      const firstRow = (`<w:tbl ${chunk}`).match(/<w:tr[\s\S]*?<\/w:tr>/i)?.[0] ?? ''
      const hasHeader = /<w:tblHeader\s*\/>/i.test(firstRow) || /<w:b\s*\/>/i.test(firstRow)
      const cells = [...firstRow.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1].trim())
      const hasContent = cells.some((c) => c.length > 0)
      if (hasContent && !hasHeader) {
        drafts.push({
          id: `table-header-${i + 1}`,
          wcagKey: 'info-relationships',
          title: 'טבלה ללא כותרות ברורות',
          severity: 'Medium',
          category: 'טבלאות',
          affectedUsers: ['משתמשי קורא מסך'],
          impact: `בטבלה ${i + 1} לא זוהתה שורת כותרת ברורה.`,
          recommendation: 'סמנו שורת כותרת בטבלה.',
          location: `טבלה ${i + 1}`,
        })
      }
    })
  }

  let hyperlinkCount = 0
  paras.forEach((p, idx) => {
    if (!/<w:hyperlink/i.test(p)) return
    hyperlinkCount++
    const texts = [...p.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1]).join('').trim()
    if (isUnclearLink(texts)) {
      drafts.push({
        id: `unclear-link-p${idx + 1}`,
        wcagKey: 'link-purpose',
        title: 'קישור לא ברור',
        severity: 'Medium',
        category: 'ניווט',
        affectedUsers: ['משתמשי קורא מסך'],
        impact: `בפסקה ${idx + 1} זוהה קישור: "${texts}".`,
        recommendation: 'השתמשו בטקסט קישור תיאורי.',
        location: `פסקה ${idx + 1}`,
      })
    }
  })
  signals.hyperlinkCount = hyperlinkCount

  if (emptyCount >= 8 && paras.length >= 10) {
    drafts.push({
      id: 'empty-paragraphs',
      wcagKey: 'info-relationships',
      title: 'שימוש ברווחים ריקים לעיצוב',
      severity: 'Medium',
      category: 'מבנה',
      affectedUsers: ['משתמשי קורא מסך'],
      impact: `זוהו ${emptyCount} פסקאות ריקות לריווח.`,
      recommendation: 'השתמשו במרווח לפני/אחרי פסקה.',
      location: 'מסמך Word',
    })
  }

  return { drafts, signals }
}
