import type { ExtractedSignals, IssueDraft } from './issue.builder.js'
import {
  collectText,
  isUnclearLink,
  listZipEntries,
  parseXml,
  readZip,
  readZipEntryText,
} from './ooxml.utils.js'

// ── Utility helpers ───────────────────────────────────────────────────────────

function toArr<T>(x: T | T[] | null | undefined): T[] {
  if (x == null) return []
  return Array.isArray(x) ? x : [x]
}

function attr(node: unknown, name: string): string {
  if (!node || typeof node !== 'object') return ''
  return String((node as Record<string, unknown>)[`@_${name}`] ?? '')
}

// Generic shape names that are NOT meaningful alt text
const GENERIC_NAME_RE =
  /^(picture|image|graphic|freeform|autoshape|rectangle|square|oval|ellipse|circle|shape|textbox|text\s*box|group|icon|chart|object|connector|line|curve|arc|arrow|smartart|diagram|video|audio|media|content|placeholder|תמונה|צורה|אובייקט|קבוצה|גרף|תרשים|וידאו|שמע|מדיה|אייקון)\s*\d*\.?$/i

function isMeaningfulAlt(text: string): boolean {
  const t = text.trim()
  return t.length >= 2 && !GENERIC_NAME_RE.test(t)
}

function objectHasAltText(title: string, descr: string): boolean {
  return isMeaningfulAlt(title) || isMeaningfulAlt(descr)
}

// ── cNvPr extraction ──────────────────────────────────────────────────────────

function extractCNvPr(node: Record<string, unknown>) {
  const holders = ['nvSpPr', 'nvPicPr', 'nvGraphicFramePr', 'nvGrpSpPr', 'nvCxnSpPr']
  for (const h of holders) {
    const nvPr = node[h] as Record<string, unknown> | undefined
    if (!nvPr) continue
    const cNvPr = nvPr['cNvPr'] as Record<string, unknown> | undefined
    if (cNvPr) {
      return {
        id: attr(cNvPr, 'id'),
        name: attr(cNvPr, 'name'),
        title: attr(cNvPr, 'title'),
        descr: attr(cNvPr, 'descr'),
      }
    }
  }
  return { id: '', name: '', title: '', descr: '' }
}

// ── Placeholder detection ─────────────────────────────────────────────────────

function getPlaceholderType(sp: Record<string, unknown>): string | null {
  const nvSpPr = sp['nvSpPr'] as Record<string, unknown> | undefined
  if (!nvSpPr) return null
  const nvPr = nvSpPr['nvPr'] as Record<string, unknown> | undefined
  if (!nvPr) return null
  const ph = nvPr['ph']
  if (ph === undefined || ph === null) return null
  if (typeof ph !== 'object' || Array.isArray(ph)) return 'body'
  return attr(ph as Record<string, unknown>, 'type') || 'body'
}

function isTitlePh(phType: string | null): boolean {
  return phType === 'title' || phType === 'ctrTitle'
}

// ── Shape fill classification ─────────────────────────────────────────────────

type FillKind = 'blip' | 'gradient' | 'solid' | 'pattern' | 'line' | 'none'

function getSpFillKind(sp: Record<string, unknown>): FillKind {
  const spPr = sp['spPr'] as Record<string, unknown> | undefined
  if (!spPr) return 'none'

  if ('blipFill' in spPr) return 'blip'
  if ('gradFill' in spPr) return 'gradient'
  if ('solidFill' in spPr) return 'solid'
  if ('pattFill' in spPr) return 'pattern'

  // Detect line/connector geometry
  const prstGeom = spPr['prstGeom'] as Record<string, unknown> | undefined
  if (prstGeom) {
    const prst = attr(prstGeom, 'prst').toLowerCase()
    if (prst === 'line' || prst === 'lineInv' || prst === 'bentConnector2' || prst === 'curvedConnector2') return 'line'
  }
  // Presence of <a:ln> (border/stroke) without fill on a simple shape → treat as line
  if ('ln' in spPr && !('solidFill' in spPr) && !('gradFill' in spPr)) return 'line'

  return 'none'
}

// ── GraphicFrame type ─────────────────────────────────────────────────────────

type GfType = 'chart' | 'table' | 'smartart' | 'icon' | 'media' | 'other'

function getGfType(gf: Record<string, unknown>): GfType {
  const graphic = gf['graphic'] as Record<string, unknown> | undefined
  const gData = graphic?.['graphicData'] as Record<string, unknown> | undefined
  if (!gData) return 'other'
  const uri = attr(gData, 'uri').toLowerCase()
  if (uri.includes('chart')) return 'chart'
  if (uri.includes('/table') || 'tbl' in gData) return 'table'
  if (uri.includes('diagram')) return 'smartart'
  if (uri.includes('svg') || uri.includes('icon')) return 'icon'
  if (uri.includes('media') || uri.includes('video') || uri.includes('audio')) return 'media'
  const keys = Object.keys(gData)
  if (keys.some((k) => k.includes('chart'))) return 'chart'
  if (keys.some((k) => k === 'tbl')) return 'table'
  if (keys.some((k) => k.includes('dgm'))) return 'smartart'
  return 'other'
}

function tableHasHeaderRow(gf: Record<string, unknown>): boolean {
  const gData = (gf['graphic'] as Record<string, unknown> | undefined)?.['graphicData'] as
    | Record<string, unknown>
    | undefined
  const tbl = gData?.['tbl'] as Record<string, unknown> | undefined
  if (!tbl) return false
  const tblPr = tbl['tblPr'] as Record<string, unknown> | undefined
  if (!tblPr) return false
  const tblLook = tblPr['tblLook'] as Record<string, unknown> | undefined
  if (!tblLook) return attr(tblPr, 'firstRow') === '1'
  if (attr(tblLook, 'firstRow') === '1') return true
  const val = attr(tblLook, 'val')
  return val ? (parseInt(val, 16) & 0x0020) !== 0 : false
}

// ── Object model ──────────────────────────────────────────────────────────────

type ObjCategory =
  | 'placeholder'
  | 'text-box'
  | 'picture'
  | 'image-filled-shape'
  | 'visual-shape'    // gradFill/solidFill custom shape — visual content, needs alt text
  | 'chart'
  | 'table'
  | 'smartart'
  | 'icon'
  | 'media'
  | 'connector'
  | 'simple-shape'
  | 'group'
  | 'other'

interface SlideObj {
  category: ObjCategory
  cNvPr: ReturnType<typeof extractCNvPr>
  phType: string | null
  text: string
  needsAltText: boolean
  hasAltText: boolean
  isDecorativeCandidate: boolean
  ignore: boolean
  ignoreReason: string
  tableHasHeader?: boolean
  isGroup: boolean
  childCount: number
}

// ── Shape classifier ──────────────────────────────────────────────────────────

function classifySp(sp: Record<string, unknown>): SlideObj {
  const cNvPr = extractCNvPr(sp)
  const phType = getPlaceholderType(sp)
  const text = collectText(sp['txBody']).trim()
  const hasMeaningfulText = text.length > 2

  // Placeholders are always ignored for alt text (they read their own content)
  if (phType !== null) {
    return {
      category: 'placeholder', cNvPr, phType, text,
      needsAltText: false, hasAltText: false,
      isDecorativeCandidate: false, ignore: true,
      ignoreReason: `placeholder type="${phType}"`,
      isGroup: false, childCount: 0,
    }
  }

  const fillKind = getSpFillKind(sp)

  // Image fill → always needs alt text
  if (fillKind === 'blip') {
    const hasAlt = objectHasAltText(cNvPr.title, cNvPr.descr)
    return {
      category: 'image-filled-shape', cNvPr, phType: null, text,
      needsAltText: true, hasAltText: hasAlt,
      isDecorativeCandidate: false, ignore: false, ignoreReason: '',
      isGroup: false, childCount: 0,
    }
  }

  // Gradient or pattern fill custom shape — visual graphic that might convey information
  // WCAG 1.1.1: non-text content needs a text alternative or must be marked decorative.
  // These are flagged as needing alt text; the user decides alt text or decorative.
  if (fillKind === 'gradient' || fillKind === 'pattern') {
    const hasAlt = objectHasAltText(cNvPr.title, cNvPr.descr)
    // If it also has readable text, the text makes it accessible
    if (hasMeaningfulText) {
      return {
        category: 'text-box', cNvPr, phType: null, text,
        needsAltText: false, hasAltText: false,
        isDecorativeCandidate: false, ignore: true,
        ignoreReason: 'gradient shape with readable text',
        isGroup: false, childCount: 0,
      }
    }
    return {
      category: 'visual-shape', cNvPr, phType: null, text: '',
      needsAltText: true, hasAltText: hasAlt,
      isDecorativeCandidate: false, ignore: false, ignoreReason: '',
      isGroup: false, childCount: 0,
    }
  }

  // Solid fill custom shape without text → decorative Quick Fix candidate
  // These are purely-colored shapes used as visual accents
  if (fillKind === 'solid') {
    if (hasMeaningfulText) {
      return {
        category: 'text-box', cNvPr, phType: null, text,
        needsAltText: false, hasAltText: false,
        isDecorativeCandidate: false, ignore: true,
        ignoreReason: 'solid-fill shape with readable text',
        isGroup: false, childCount: 0,
      }
    }
    return {
      category: 'simple-shape', cNvPr, phType: null, text: '',
      needsAltText: false, hasAltText: false,
      isDecorativeCandidate: true, ignore: false, ignoreReason: '',
      isGroup: false, childCount: 0,
    }
  }

  // Line / connector geometry → decorative
  if (fillKind === 'line') {
    return {
      category: 'connector', cNvPr, phType: null, text: '',
      needsAltText: false, hasAltText: false,
      isDecorativeCandidate: true, ignore: false, ignoreReason: '',
      isGroup: false, childCount: 0,
    }
  }

  // No meaningful fill. Check for readable text.
  if (hasMeaningfulText) {
    return {
      category: 'text-box', cNvPr, phType: null, text,
      needsAltText: false, hasAltText: false,
      isDecorativeCandidate: false, ignore: true,
      ignoreReason: 'text box with readable text',
      isGroup: false, childCount: 0,
    }
  }

  // Empty shape with no fill, no image, no text → decorative candidate
  return {
    category: 'simple-shape', cNvPr, phType: null, text: '',
    needsAltText: false, hasAltText: false,
    isDecorativeCandidate: true, ignore: false, ignoreReason: '',
    isGroup: false, childCount: 0,
  }
}

function classifyPic(pic: Record<string, unknown>): SlideObj {
  const cNvPr = extractCNvPr(pic)
  return {
    category: 'picture', cNvPr, phType: null, text: '',
    needsAltText: true, hasAltText: objectHasAltText(cNvPr.title, cNvPr.descr),
    isDecorativeCandidate: false, ignore: false, ignoreReason: '',
    isGroup: false, childCount: 0,
  }
}

function classifyGf(gf: Record<string, unknown>): SlideObj {
  const cNvPr = extractCNvPr(gf)
  const gfType = getGfType(gf)

  if (gfType === 'table') {
    return {
      category: 'table', cNvPr, phType: null, text: '',
      needsAltText: false, hasAltText: false,
      isDecorativeCandidate: false, ignore: false, ignoreReason: '',
      tableHasHeader: tableHasHeaderRow(gf),
      isGroup: false, childCount: 0,
    }
  }

  return {
    category: gfType === 'other' ? 'other' : (gfType as ObjCategory),
    cNvPr, phType: null, text: '',
    needsAltText: true, hasAltText: objectHasAltText(cNvPr.title, cNvPr.descr),
    isDecorativeCandidate: false, ignore: false, ignoreReason: '',
    isGroup: false, childCount: 0,
  }
}

function classifyCxn(cxn: Record<string, unknown>): SlideObj {
  return {
    category: 'connector', cNvPr: extractCNvPr(cxn), phType: null, text: '',
    needsAltText: false, hasAltText: false,
    isDecorativeCandidate: true, ignore: false, ignoreReason: '',
    isGroup: false, childCount: 0,
  }
}

// ── Slide analysis ─────────────────────────────────────────────────────────────

interface SlideAnalysis {
  slideNum: number
  hasTitle: boolean
  titleText: string
  objects: SlideObj[]
  missingAltObjs: SlideObj[]
  decorativeObjs: SlideObj[]
  tablesTotal: number
  tablesMissingHeader: number
  readingOrderRisk: boolean
  unclearLinks: string[]
}

function emptySlide(slideNum: number): SlideAnalysis {
  return {
    slideNum, hasTitle: false, titleText: '', objects: [],
    missingAltObjs: [], decorativeObjs: [],
    tablesTotal: 0, tablesMissingHeader: 0,
    readingOrderRisk: false, unclearLinks: [],
  }
}

function analyzeSlide(slideNum: number, xml: string): SlideAnalysis {
  let parsed: Record<string, unknown>
  try {
    parsed = parseXml<Record<string, unknown>>(xml)
  } catch {
    return emptySlide(slideNum)
  }

  const sld = parsed['sld'] as Record<string, unknown> | undefined
  const spTree = (sld?.['cSld'] as Record<string, unknown> | undefined)?.['spTree'] as
    | Record<string, unknown>
    | undefined
  if (!spTree) return emptySlide(slideNum)

  const objects: SlideObj[] = []

  // ── Process top-level shapes ───────────────────────────────────────────────
  for (const sp of toArr(spTree['sp'] as unknown) as Record<string, unknown>[]) objects.push(classifySp(sp))
  for (const pic of toArr(spTree['pic'] as unknown) as Record<string, unknown>[]) objects.push(classifyPic(pic))
  for (const gf of toArr(spTree['graphicFrame'] as unknown) as Record<string, unknown>[]) objects.push(classifyGf(gf))
  for (const cxn of toArr(spTree['cxnSp'] as unknown) as Record<string, unknown>[]) objects.push(classifyCxn(cxn))

  // ── Process group children individually (groups are transparent) ───────────
  // PowerPoint groups are a container; the individual shapes inside are the
  // actual reading objects. We classify each child sp/pic/gf directly.
  // This avoids double-counting group + children, and correctly surfaces
  // visual shapes (blipFill, gradFill) and decorative shapes (solidFill, line).
  let groupChildrenSkipped = 0
  for (const grp of toArr(spTree['grpSp'] as unknown) as Record<string, unknown>[]) {
    const childSps = toArr(grp['sp'] as unknown) as Record<string, unknown>[]
    const childPics = toArr(grp['pic'] as unknown) as Record<string, unknown>[]
    const childGfs = toArr(grp['graphicFrame'] as unknown) as Record<string, unknown>[]
    const childCxns = toArr(grp['cxnSp'] as unknown) as Record<string, unknown>[]

    for (const sp of childSps) objects.push(classifySp(sp))
    for (const pic of childPics) objects.push(classifyPic(pic))
    for (const gf of childGfs) objects.push(classifyGf(gf))
    for (const cxn of childCxns) objects.push(classifyCxn(cxn))

    // Note: nested grpSp inside grpSp are skipped (very rare, usually decorative frames)
    groupChildrenSkipped += toArr(grp['grpSp'] as unknown).length
  }

  // ── Title detection ────────────────────────────────────────────────────────
  let hasTitle = false
  let titleText = ''
  for (const obj of objects) {
    if (obj.category === 'placeholder' && isTitlePh(obj.phType) && obj.text.length > 2) {
      hasTitle = true
      titleText = obj.text
      break
    }
  }

  const missingAltObjs = objects.filter((o) => !o.ignore && o.needsAltText && !o.hasAltText)
  const decorativeObjs = objects.filter((o) => !o.ignore && o.isDecorativeCandidate)

  const tables = objects.filter((o) => o.category === 'table')
  const tablesTotal = tables.length
  const tablesMissingHeader = tables.filter((o) => !o.tableHasHeader).length

  // Reading order risk:
  // - 4+ selectable non-ignored reading objects with at least one visual element
  // - OR any slide with 5+ visual/mixed objects and no placeholder structure
  const selectable = objects.filter((o) => !o.ignore)
  const hasVisuals = selectable.some((o) =>
    ['picture', 'image-filled-shape', 'visual-shape', 'chart', 'smartart', 'icon', 'group', 'media'].includes(o.category),
  )
  const hasPlaceholders = objects.some((o) => o.category === 'placeholder')
  // Slides with no placeholders and many visual objects always need reading-order verification
  const readingOrderRisk = selectable.length >= 4 && hasVisuals && !hasPlaceholders

  const unclearLinks: string[] = []
  for (const m of xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)) {
    const t = m[1].trim()
    if (t && isUnclearLink(t) && !unclearLinks.includes(t)) unclearLinks.push(t)
  }

  return {
    slideNum, hasTitle, titleText, objects,
    missingAltObjs, decorativeObjs,
    tablesTotal, tablesMissingHeader,
    readingOrderRisk, unclearLinks,
  }
}

// ── Title normalization for duplicate detection ───────────────────────────────

function normTitle(t: string): string {
  return t.trim().replace(/\s+/g, ' ').replace(/[.!?؟،,;:]+$/, '').toLowerCase()
}

// ── Main export ───────────────────────────────────────────────────────────────

export function analyzePptx(
  filePath: string,
  fileSize: number,
): { drafts: IssueDraft[]; signals: ExtractedSignals } {
  const zip = readZip(filePath)
  const slidePaths = listZipEntries(zip, /^ppt\/slides\/slide\d+\.xml$/i)

  const slides: SlideAnalysis[] = []
  const missingTitleSlides: number[] = []
  let totalImages = 0
  let totalMissingAlt = 0
  let totalDecorativeCandidates = 0
  let totalReadingOrderRisk = 0
  let totalTablesMissingHeader = 0

  for (const slidePath of slidePaths) {
    const xml = readZipEntryText(zip, slidePath)
    if (!xml) continue
    const slideNum = parseInt(slidePath.match(/slide(\d+)\.xml/i)?.[1] ?? '0', 10)
    const s = analyzeSlide(slideNum, xml)
    slides.push(s)

    if (!s.hasTitle) missingTitleSlides.push(slideNum)
    totalImages += s.objects.filter(
      (o) => o.category === 'picture' || o.category === 'image-filled-shape',
    ).length
    totalMissingAlt += s.missingAltObjs.length
    totalDecorativeCandidates += s.decorativeObjs.length
    if (s.readingOrderRisk) totalReadingOrderRisk++
    totalTablesMissingHeader += s.tablesMissingHeader
  }

  // Duplicate slide title detection
  const titleMap = new Map<string, number[]>()
  for (const s of slides) {
    if (s.hasTitle && s.titleText) {
      const norm = normTitle(s.titleText)
      if (!titleMap.has(norm)) titleMap.set(norm, [])
      titleMap.get(norm)!.push(s.slideNum)
    }
  }
  const duplicateTitles = [...titleMap.entries()].filter(([, nums]) => nums.length > 1)

  // ── Build issues ─────────────────────────────────────────────────────────────

  const drafts: IssueDraft[] = []

  // 1. Missing alt text — one issue for the whole file with occurrence count
  if (totalMissingAlt > 0) {
    const affectedSlides = slides.filter((s) => s.missingAltObjs.length > 0)
    drafts.push({
      id: 'missing-alt-pptx',
      wcagKey: 'non-text-content',
      title: 'חסר טקסט חלופי',
      severity: 'High',
      category: 'מדיה ואיורים',
      affectedUsers: ['משתמשים עיוורים', 'משתמשי קורא מסך'],
      impact: `${totalMissingAlt} אובייקט${totalMissingAlt !== 1 ? 'ים' : ''} (תמונות, גרפים, צורות ויזואליות) ללא טקסט חלופי.`,
      recommendation: 'הוסיפו טקסט חלופי תיאורי לכל תמונה, גרף ואיור דרך חלונית הנגישות של PowerPoint.',
      location: `שקופיות ${affectedSlides.map((s) => s.slideNum).join(', ')}`,
      occurrenceCount: totalMissingAlt,
      locations: affectedSlides.map((s) => `שקופית ${s.slideNum} — ${s.missingAltObjs.length} אובייקטים`),
    })
  }

  // 2. Missing slide title — ONE merged issue with occurrence count + locations
  if (missingTitleSlides.length > 0) {
    const locs = missingTitleSlides.map((n) => `שקופית ${n}`)
    drafts.push({
      id: 'missing-title-pptx',
      wcagKey: 'page-titled-headings',
      title: 'שקופית ללא כותרת',
      severity: 'Medium',
      category: 'מבנה מסמך',
      affectedUsers: ['משתמשי קורא מסך'],
      impact: `${missingTitleSlides.length} שקופיות ללא כותרת המוגדרת כ-Placeholder של כותרת.`,
      recommendation: 'הוסיפו כותרת ייחודית לכל שקופית באמצעות תיבת הכותרת הייעודית, לא תיבת טקסט רגילה.',
      location: locs.join(', '),
      occurrenceCount: missingTitleSlides.length,
      locations: locs,
    })
  }

  // 3. Duplicate slide title
  for (const [norm, slideNums] of duplicateTitles) {
    drafts.push({
      id: `duplicate-title-${norm.slice(0, 20).replace(/\s/g, '-')}`,
      wcagKey: 'page-titled-headings',
      title: 'כותרת שקופית כפולה',
      severity: 'Medium',
      category: 'מבנה מסמך',
      affectedUsers: ['משתמשי קורא מסך'],
      impact: `הכותרת "${norm}" מופיעה בשקופיות ${slideNums.join(', ')} – קשה לנווט בין שקופיות.`,
      recommendation: 'הקצו כותרת ייחודית לכל שקופית.',
      location: `שקופיות ${slideNums.join(', ')}`,
      occurrenceCount: slideNums.length,
      locations: slideNums.map((n) => `שקופית ${n}`),
    })
  }

  // 4. Reading order — ONE merged issue with occurrence count (heuristic)
  const readingOrderSlides = slides.filter((s) => s.readingOrderRisk).slice(0, 2)
  if (readingOrderSlides.length > 0) {
    const locs = readingOrderSlides.map(
      (s) => `שקופית ${s.slideNum} — ${s.objects.filter((o) => !o.ignore).length} אובייקטים`,
    )
    drafts.push({
      id: 'reading-order-pptx',
      wcagKey: 'meaningful-sequence',
      title: 'בדיקת סדר קריאה',
      severity: 'Medium',
      category: 'מבנה מסמך',
      affectedUsers: ['משתמשי קורא מסך'],
      impact: `${readingOrderSlides.length} שקופיות עם אובייקטים ויזואליים מרובים ללא מבנה Placeholder – יש לאמת סדר קריאה.`,
      recommendation: 'פתחו חלונית "סדר בחירה" ב-PowerPoint ובדקו את הסדר הלוגי של כל אובייקט.',
      location: readingOrderSlides.map((s) => `שקופית ${s.slideNum}`).join(', '),
      occurrenceCount: readingOrderSlides.length,
      locations: locs,
    })
  }

  // 5. Tables without header — one per slide (typically rare in PPTX)
  for (const s of slides) {
    if (s.tablesMissingHeader > 0) {
      drafts.push({
        id: `table-no-header-slide-${s.slideNum}`,
        wcagKey: 'info-relationships',
        title: 'טבלה ללא שורת כותרות',
        severity: 'Medium',
        category: 'טבלאות',
        affectedUsers: ['משתמשי קורא מסך'],
        impact: `שקופית ${s.slideNum}: ${s.tablesMissingHeader} טבלה${s.tablesMissingHeader > 1 ? 'ות' : ''} ללא שורת כותרת מסומנת.`,
        recommendation: 'סמנו "שורת כותרת" בעיצוב הטבלה → עיצוב → אפשרויות סגנון → שורת כותרת.',
        location: `שקופית ${s.slideNum}`,
        occurrenceCount: s.tablesMissingHeader,
      })
    }
  }

  // 6. Unclear links
  const seenLinks = new Set<string>()
  for (const s of slides) {
    for (const link of s.unclearLinks) {
      const key = `${s.slideNum}-${link}`
      if (!seenLinks.has(key)) {
        seenLinks.add(key)
        drafts.push({
          id: `unclear-link-${s.slideNum}-${Buffer.from(link).toString('base64').slice(0, 8)}`,
          wcagKey: 'link-purpose',
          title: 'קישור לא ברור',
          severity: 'Medium',
          category: 'ניווט',
          affectedUsers: ['משתמשי קורא מסך'],
          impact: `שקופית ${s.slideNum}: קישור עם טקסט "${link}" אינו מתאר את יעד הקישור.`,
          recommendation: 'החליפו בטקסט קישור תיאורי שמסביר לאן הקישור מוביל.',
          location: `שקופית ${s.slideNum}`,
          occurrenceCount: 1,
        })
      }
    }
  }

  // 7. Quick Fix: decorative candidates — one issue with occurrence count
  if (totalDecorativeCandidates > 0) {
    const affectedSlides = slides.filter((s) => s.decorativeObjs.length > 0)
    drafts.push({
      id: 'decorative-candidates-pptx',
      wcagKey: 'non-text-content',
      title: 'אובייקטים שניתן לסמן כדקורטיביים',
      severity: 'Low',
      category: 'Quick Fix',
      affectedUsers: ['משתמשי קורא מסך'],
      impact: `${totalDecorativeCandidates} צורות (קווים, מילוי צבע אחיד, תיבות ריקות) שאינן נושאות תוכן מהותי – קוראי מסך קוראים אותן באופן מיותר.`,
      recommendation: 'סמנו אובייקטים דקורטיביים כ"דקורטיבי" בחלונית הנגישות של PowerPoint: לחץ ימני על הצורה → עריכת טקסט חלופי → סמן "דקורטיבי".',
      location: `שקופיות ${affectedSlides.map((s) => s.slideNum).join(', ')}`,
      occurrenceCount: totalDecorativeCandidates,
      locations: affectedSlides.map((s) => `שקופית ${s.slideNum} — ${s.decorativeObjs.length} צורות`),
      isQuickFix: true,
    })
  }

  // ── Debug logging ─────────────────────────────────────────────────────────────

  const officeCounts = {
    missingAltText: totalMissingAlt,
    quickFixDecorative: totalDecorativeCandidates,
    missingSlideTitle: missingTitleSlides.length,
    duplicateSlideTitle: duplicateTitles.length,
    readingOrder: totalReadingOrderRisk,
    contrastIssues: 0,
    tableIssues: totalTablesMissingHeader,
    mainIssueTotal: drafts.filter((d) => !d.isQuickFix).length,
    quickFixTotal: drafts.filter((d) => d.isQuickFix).length,
  }
  console.log('[OFFICE-LIKE COUNTS]', JSON.stringify(officeCounts, null, 2))

  const perSlideSummary = slides.map((s) => {
    const ignored = s.objects.filter((o) => o.ignore)
    const selectable = s.objects.filter((o) => !o.ignore)
    return {
      slide: s.slideNum,
      topLevelReadingObjects: selectable.length,
      ignoredPlaceholders: ignored.filter((o) => o.category === 'placeholder').length,
      ignoredTextBoxes: ignored.filter((o) => o.category === 'text-box').length,
      imageObjects: s.objects.filter((o) => o.category === 'picture' || o.category === 'image-filled-shape').length,
      visualShapeObjects: s.objects.filter((o) => o.category === 'visual-shape').length,
      missingAltText: s.missingAltObjs.length,
      decorativeCandidates: s.decorativeObjs.length,
      tablesTotal: s.tablesTotal,
      tablesMissingHeader: s.tablesMissingHeader,
      readingOrderRisk: s.readingOrderRisk,
    }
  })
  console.log('[OBJECT CLASSIFICATION SUMMARY]', JSON.stringify(perSlideSummary, null, 2))

  const signals: ExtractedSignals = {
    fileType: 'pptx',
    fileSize,
    slideCount: slides.length,
    imageCount: totalImages,
    missingTitleSlides,
    quickFixCount: totalDecorativeCandidates,
  }

  return { drafts, signals }
}
