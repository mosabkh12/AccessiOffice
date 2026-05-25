import type { ExtractedSignals, IssueDraft } from './scannerShared.js'
import {
  CAT_COLOR_CONTRAST,
  CAT_DOCUMENT_STRUCTURE,
  CAT_MEDIA,
  OFFICE_TITLES,
} from './officeCategories.js'
import {
  isGenericAltText,
  isMissingAltText,
  isUnclearLink,
  listZipEntries,
  parseXml,
  ptFromSz,
  readZip,
  readZipEntryText,
  wordCount,
} from './ooxml.utils.js'
import type AdmZip from 'adm-zip'
import { snippet, type AccessibilityState, type ObjectClassification } from './pptxDiagnostics.js'

export const PPTX_SCANNER_BUILD = 'DEBUG-PPTX-V7-OFFICE-COUNTS'

const DECORATIVE_SCORE_THRESHOLD = 4
const DECORATIVE_SCORE_STRICT = 5
const SMALL_SHAPE_AREA = 800_000_000_000
const THIN_SHAPE_EMU = 180_000

const BOILERPLATE_TITLE_PATTERNS = [
  /^click to edit master title style$/i,
  /^click to edit master subtitle style$/i,
  /^click to edit master text styles/i,
  /^title text$/i,
  /^subtitle$/i,
  /^שקופית\s*\d+$/i,
  /^slide\s*\d+$/i,
]

const BOILERPLATE_BODY_PATTERNS = [
  /click to edit master/i,
  /second level/i,
  /third level/i,
  /fourth level/i,
  /^‹#›$/i,
  /^<\d+>$/i,
]

const SHAPE_TAGS = ['grpSp', 'graphicFrame', 'pic', 'cxnSp', 'sp'] as const
type ShapeTag = (typeof SHAPE_TAGS)[number]

interface RawShapeElement {
  tag: ShapeTag
  xml: string
  inGroup: boolean
  fromLayout: boolean
  groupDepth: number
}

type ObjectKind =
  | 'picture'
  | 'shape'
  | 'freeform'
  | 'autoshape'
  | 'textBox'
  | 'connector'
  | 'group'
  | 'graphicFrame'
  | 'chart'
  | 'smartArt'
  | 'media'
  | 'other'

interface InventoryObject {
  slideNum: number
  index: number
  tag: ShapeTag
  xmlTag: string
  objectKind: ObjectKind
  cnvId: string
  descr?: string
  title?: string
  name?: string
  hasBlip: boolean
  hasGraphicData: boolean
  graphicUri?: string
  hasHyperlink: boolean
  relationshipId?: string
  relationshipTarget?: string
  hasText: boolean
  textSnippet: string
  isPlaceholder: boolean
  placeholderType?: string
  isTitlePlaceholder: boolean
  isBodyPlaceholder: boolean
  isFooterPlaceholder: boolean
  isSlideNumberPlaceholder: boolean
  isDatePlaceholder: boolean
  isContentPlaceholder: boolean
  isCustomShape: boolean
  hasSolidFill: boolean
  hasNoFill: boolean
  hasOutline: boolean
  isMarkedDecorative: boolean
  isGenericName: boolean
  isGroup: boolean
  inGroup: boolean
  groupDepth: number
  fromLayout: boolean
  x: number
  y: number
  cx: number
  cy: number
  hidden: boolean
  accessibilityState: AccessibilityState
  classifiedAs: ObjectClassification
  decorativeScore: number
  meaningfulVisualScore: number
  reason: string
  countedAs: 'object' | 'group-container-skipped'
}

interface SlideTitleInfo {
  slideNum: number
  hasTitlePlaceholder: boolean
  titlePlaceholdersFound: number
  titleTexts: string[]
  hiddenTitleTexts: string[]
  chosenTitle: string
  normalizedTitle: string
  titleSource: string
  layoutType?: string
  isMissingTitle: boolean
  missingTitleReason: string
}

interface SlideSlideSummary {
  slide: number
  chosenTitle: string
  missingTitle: boolean
  duplicateTitle: boolean
  visualObjectsCount: number
  customShapesWithText: number
  imageLikeObjects: number
  missingAltObjects: number
  decorativeCandidates: number
  readingOrderScore: number
  readingOrderRisk: boolean
}

function attrFromChunk(tag: string, name: string): string | undefined {
  return tag.match(new RegExp(name + '="([^"]*)"', 'i'))?.[1]
}

function chunkText(chunk: string): string {
  const parts = [...chunk.matchAll(/<a:t[^>]*>([^<]*)<\/a:t>/g)].map((m) => m[1])
  return parts.join('').trim()
}

function cleanTitleText(raw: string): string {
  return raw
    .replace(/[a-z]{2}-IL\d+/gi, '')
    .replace(/[a-z]{2}-[A-Z]{2}\d+/gi, '')
    .replace(/\s+/g, ' ')
    .trim()
}

function normalizeTitle(text: string): string {
  return cleanTitleText(text)
    .normalize('NFKC')
    .replace(/[\u200B-\u200D\uFEFF\u00AD]/g, '')
    .replace(/\u00a0/g, ' ')
    .replace(/[.,;:!?״׳]+$/u, '')
    .replace(/[^\p{L}\p{N}\s]/gu, '')
    .toLowerCase()
}

function normalizeTitleForDuplicate(text: string): string {
  return cleanTitleText(text)
    .normalize('NFKC')
    .replace(/[\u200B-\u200D\uFEFF\u00AD]/g, '')
    .replace(/\u00a0/g, ' ')
    .replace(/[.,;:!?״׳]+$/u, '')
    .toLowerCase()
}

function hashChunk(s: string): string {
  let h = 0
  for (let i = 0; i < Math.min(s.length, 200); i++) h = (h * 31 + s.charCodeAt(i)) | 0
  return Math.abs(h).toString(36)
}

function extractBalancedTag(xml: string, startIndex: number): string | null {
  const slice = xml.slice(startIndex)
  const openMatch = slice.match(/^<((?:p:|pic:)?(\w+))\b/i)
  if (!openMatch) return null
  const localName = openMatch[2]
  const openTagRe = new RegExp(`<(?:p:|pic:)?${localName}\\b`, 'gi')
  const closeTagRe = new RegExp(`</(?:p:|pic:)?${localName}>`, 'gi')

  let depth = 0
  let searchStart = startIndex
  while (searchStart < xml.length) {
    const tail = xml.slice(searchStart)
    openTagRe.lastIndex = 0
    closeTagRe.lastIndex = 0
    const openM = openTagRe.exec(tail)
    const closeM = closeTagRe.exec(tail)
    if (!openM && !closeM) break
    const openPos = openM ? openM.index : Number.POSITIVE_INFINITY
    const closePos = closeM ? closeM.index : Number.POSITIVE_INFINITY
    if (openPos < closePos) {
      depth++
      searchStart += openM!.index + openM![0].length
    } else {
      depth--
      const end = searchStart + closeM!.index + closeM![0].length
      if (depth === 0) return xml.slice(startIndex, end)
      searchStart = end
    }
  }
  return null
}

function findAllTaggedChunks(xml: string, localName: string): string[] {
  const chunks: string[] = []
  let pos = 0
  const openRe = new RegExp(`<(?:p:|pic:)?${localName}\\b`, 'gi')
  while (pos < xml.length) {
    openRe.lastIndex = pos
    const m = openRe.exec(xml)
    if (!m) break
    const chunk = extractBalancedTag(xml, m.index)
    if (chunk) {
      chunks.push(chunk)
      pos = m.index + chunk.length
    } else {
      pos = m.index + 1
    }
  }
  return chunks
}

function isBoilerplateTitleText(text: string): boolean {
  const t = cleanTitleText(text)
  if (!t) return true
  return BOILERPLATE_TITLE_PATTERNS.some((p) => p.test(t))
}

function isMasterBoilerplateBodyText(text: string): boolean {
  const t = text.trim()
  if (!t) return false
  return BOILERPLATE_BODY_PATTERNS.some((p) => p.test(t))
}

function isListOrBodyContent(text: string): boolean {
  if (/\d+\.\s/.test(text)) return true
  if ((text.match(/\d+\./g) ?? []).length >= 2) return true
  return wordCount(text) > 12
}

function isMasterChromeName(name?: string): boolean {
  if (!name) return false
  return /placeholder|footer|date|slide number|content placeholder|subtitle placeholder/i.test(name)
}

function getPlaceholderInfo(chunk: string): { isPlaceholder: boolean; type?: string } {
  const m = chunk.match(/<p:ph\b[^>]*\btype="([^"]*)"/i)
  if (m) return { isPlaceholder: true, type: m[1] }
  const nvPh = chunk.match(/<p:nvSpPr[\s\S]*?<p:ph\b[^>]*\btype="([^"]*)"/i)
  if (nvPh) return { isPlaceholder: true, type: nvPh[1] }
  const name = parseCnvFromChunk(chunk).name ?? ''
  if (isMasterChromeName(name)) {
    if (/title/i.test(name)) return { isPlaceholder: true, type: 'title' }
    if (/content|body|subtitle/i.test(name)) return { isPlaceholder: true, type: 'body' }
    if (/slide number/i.test(name)) return { isPlaceholder: true, type: 'sldNum' }
    if (/date/i.test(name)) return { isPlaceholder: true, type: 'dt' }
    if (/footer/i.test(name)) return { isPlaceholder: true, type: 'ftr' }
    return { isPlaceholder: true, type: 'obj' }
  }
  return { isPlaceholder: false }
}

function hasVisibleOutline(chunkXml: string): boolean {
  const lnMatch = chunkXml.match(/<a:ln\b[^>]*\bw="(\d+)"/i)
  if (!lnMatch) return false
  if (parseInt(lnMatch[1], 10) < 10_000) return false
  const lnSection = chunkXml.match(/<a:ln\b[\s\S]*?<\/a:ln>/i)?.[0] ?? ''
  return /<a:solidFill|<a:gradFill|<a:schemeClr/i.test(lnSection) || !/<a:noFill\s*\/>/.test(lnSection)
}

function isVisualBannerShape(obj: InventoryObject, fullText: string, chunkXml: string): boolean {
  if (!obj.hasText || isMasterBoilerplateBodyText(fullText)) return false
  if (isListOrBodyContent(fullText)) return false
  const words = wordCount(fullText)
  if (words > 12) return false
  if (obj.y > 5_500_000) return false
  const isTopBand =
    obj.hasSolidFill && !obj.hasNoFill && obj.y < 2_500_000 && obj.cy >= 350_000
  const isWideTitleCard =
    obj.hasSolidFill && !obj.hasNoFill && obj.cx >= 5_000_000 && obj.cy >= 700_000 && words <= 10
  const isRoundRectBanner =
    obj.hasSolidFill &&
    /<a:prstGeom\b[^>]*\bprst="roundRect"/i.test(chunkXml) &&
    words <= 10 &&
    obj.cy >= 400_000
  const isOutlinedTextCard =
    obj.hasNoFill &&
    hasVisibleOutline(chunkXml) &&
    obj.cx >= 5_000_000 &&
    obj.cy >= 400_000 &&
    words <= 10
  const isTxBoxVisualCard =
    /<p:cNvSpPr[^>]*\btxBox="1"/i.test(chunkXml) &&
    obj.hasNoFill &&
    obj.cx >= 5_000_000 &&
    obj.cy >= 400_000 &&
    words <= 10 &&
    /<a:ln\b[^>]*\bw="(?:1[2-9]|[2-9]\d)\d{3,}"/i.test(chunkXml)
  return isTopBand || isWideTitleCard || isRoundRectBanner || isOutlinedTextCard || isTxBoxVisualCard
}

function isTitlePhType(type?: string): boolean {
  return type === 'title' || type === 'ctrTitle'
}

function isBodyPhType(type?: string): boolean {
  return type === 'body' || type === 'subTitle' || type === 'obj' || type === 'clipArt' || type === 'media'
}

function isFooterPhType(type?: string): boolean {
  return type === 'ftr' || type === 'dt'
}

function isSlideNumberPhType(type?: string): boolean {
  return type === 'sldNum'
}

function mapAccessibilityToLegacy(state: AccessibilityState): ObjectClassification {
  switch (state) {
    case 'needsAltText':
      return 'missing-alt'
    case 'decorativeCandidate':
      return 'decorative-candidate'
    case 'alreadyAccessible':
      return 'has-alt'
    case 'ignoreFromAccessibility':
    default:
      return 'skipped-not-visual'
  }
}

function detectObjectKind(tag: ShapeTag, chunkXml: string, name?: string): ObjectKind {
  if (tag === 'pic') return 'picture'
  if (tag === 'cxnSp') return 'connector'
  if (tag === 'graphicFrame') {
    const uri = chunkXml.match(/<a:graphicData[^>]*\buri="([^"]+)"/i)?.[1] ?? ''
    if (/chart/i.test(uri)) return 'chart'
    if (/diagram/i.test(uri)) return 'smartArt'
    if (/video|movie|media/i.test(uri)) return 'media'
    return 'graphicFrame'
  }
  if (/<a:custGeom\b/i.test(chunkXml)) return 'freeform'
  if (/<p:cNvSpPr[^>]*\btxBox="1"/i.test(chunkXml) || /textbox/i.test(name ?? '')) return 'textBox'
  if (/<a:prstGeom\b/i.test(chunkXml)) return 'autoshape'
  return 'shape'
}

function hasGraphicDataUri(chunk: string): { has: boolean; uri?: string } {
  const uri = chunk.match(/<a:graphicData[^>]*\buri="([^"]+)"/i)?.[1]
  return { has: Boolean(uri), uri }
}

function hasHyperlinkInChunk(chunk: string): boolean {
  return /<a:hlinkClick\b/i.test(chunk) || /<p:hlinkClick\b/i.test(chunk)
}

function shapeArea(obj: InventoryObject): number {
  return obj.cx > 0 && obj.cy > 0 ? obj.cx * obj.cy : 0
}

function isThinOrSmall(obj: InventoryObject): boolean {
  if (obj.cx > 0 && obj.cy > 0 && (obj.cx < THIN_SHAPE_EMU || obj.cy < THIN_SHAPE_EMU)) return true
  const area = shapeArea(obj)
  return area > 0 && area < SMALL_SHAPE_AREA
}

function isNearSlideEdge(obj: InventoryObject): boolean {
  return obj.x < 400_000 || obj.y < 400_000
}

function buildShapeSignature(obj: InventoryObject, chunkXml: string): string {
  const prst = chunkXml.match(/<a:prstGeom\b[^>]*\bprst="([^"]+)"/i)?.[1] ?? ''
  return `${obj.name ?? ''}|${prst}|${obj.cx}|${obj.cy}|${obj.fromLayout ? 'L' : 'S'}`
}

function scoreDecorative(
  obj: InventoryObject,
  chunkXml: string,
  fullText: string,
  repeatCount: number,
): number {
  let score = 0
  if (obj.tag === 'cxnSp') score += 4
  if (obj.objectKind === 'connector') score += 3
  if (!obj.hasText && !obj.hasBlip && !obj.hasGraphicData) score += 2
  if ((obj.hasSolidFill || obj.hasOutline) && !obj.hasText && !obj.hasBlip) score += 2
  if (isThinOrSmall(obj)) score += 2
  if (obj.objectKind === 'freeform') score += 2
  if (obj.objectKind === 'autoshape' && !obj.hasText) score += 1
  if (/<a:prstGeom\b[^>]*\bprst="(rect|ellipse|roundRect|triangle|rtTriangle|line)"/i.test(chunkXml) && !obj.hasText) {
    score += 1
  }
  if (obj.fromLayout && !obj.hasBlip) score += 1
  if (repeatCount >= 3) score += 1
  if (repeatCount >= 5) score += 2
  if (isNearSlideEdge(obj)) score += 1
  if (/connector|straight|line|curve/i.test(obj.name ?? '')) score += 2
  if (obj.hasText && wordCount(fullText) <= 2 && !obj.hasBlip && isThinOrSmall(obj)) score += 1
  if (obj.name && /rounded rectangle/i.test(obj.name) && !obj.hasText) score += 2
  if (obj.hidden) score -= 5
  return score
}

function scoreMeaningfulVisual(
  obj: InventoryObject,
  chunkXml: string,
  fullText: string,
): number {
  let score = 0
  if (obj.tag === 'pic' || obj.objectKind === 'picture') score += 5
  if (obj.tag === 'graphicFrame' || obj.hasGraphicData) score += 5
  if (/chart|diagram|video|ink|oleObject|picture/i.test(obj.graphicUri ?? chunkXml)) score += 4
  if (obj.hasBlip) score += 4
  if (obj.hasHyperlink) score += 1
  if (obj.hasText && !obj.isPlaceholder && (obj.hasSolidFill || obj.hasOutline)) score += 2
  if (
    obj.isCustomShape &&
    !obj.isTitlePlaceholder &&
    !obj.isBodyPlaceholder &&
    isVisualBannerShape(obj, fullText, chunkXml)
  ) {
    score += 3
  }
  if (shapeArea(obj) >= 2_000_000_000_000) score += 1
  if (obj.hasText && wordCount(fullText) <= 8 && obj.hasSolidFill && isVisualBannerShape(obj, fullText, chunkXml)) score += 2
  if (obj.inGroup && (obj.hasBlip || obj.hasSolidFill)) score += 1
  if (obj.objectKind === 'smartArt' || obj.objectKind === 'chart' || obj.objectKind === 'media') score += 5
  return score
}

function shouldIgnoreFromAccessibility(
  obj: InventoryObject,
  fullText: string,
): { ignore: boolean; reason: string } {
  if (obj.hidden) return { ignore: true, reason: 'hidden or off-slide object' }
  if (obj.isTitlePlaceholder && !obj.hasBlip) {
    return { ignore: true, reason: 'title placeholder' }
  }
  if (obj.isBodyPlaceholder && obj.hasText && !obj.hasBlip) {
    return { ignore: true, reason: 'body/content placeholder with readable text' }
  }
  if (obj.isFooterPlaceholder || obj.isSlideNumberPlaceholder || obj.isDatePlaceholder) {
    return { ignore: true, reason: 'footer/date/slide-number placeholder' }
  }
  if (obj.isPlaceholder && obj.hasText && !obj.hasBlip && !obj.hasSolidFill && !obj.hasOutline) {
    return { ignore: true, reason: 'text placeholder carries content' }
  }
  if (obj.fromLayout && obj.hasText && isMasterBoilerplateBodyText(fullText) && !obj.hasBlip) {
    return { ignore: false, reason: '' }
  }
  if (isMasterChromeName(obj.name) && (obj.isPlaceholder || isMasterBoilerplateBodyText(fullText))) {
    return { ignore: true, reason: 'master chrome placeholder' }
  }
  if (obj.hasText && isMasterBoilerplateBodyText(fullText) && !obj.fromLayout) {
    return { ignore: true, reason: 'master default placeholder text' }
  }
  if (obj.countedAs === 'group-container-skipped') {
    return { ignore: true, reason: 'group container — children counted separately' }
  }
  return { ignore: false, reason: '' }
}

/** Shape text is exposed to AT via a:t — not a visual-only object (Reading Order Pane text object). */
function isTextReadableObject(
  obj: InventoryObject,
  fullText: string,
  chunkXml: string,
): boolean {
  if (!obj.hasText || obj.hasBlip || obj.hasGraphicData) return false
  if (obj.isTitlePlaceholder || obj.isBodyPlaceholder || obj.isFooterPlaceholder) return true
  if (obj.isSlideNumberPlaceholder || obj.isDatePlaceholder) return true
  if (obj.objectKind === 'textBox') return true
  if (obj.isPlaceholder && obj.hasText) return true
  return false
}

function isPictureLikeObject(obj: InventoryObject): boolean {
  return (
    obj.tag === 'pic' ||
    obj.hasBlip ||
    obj.hasGraphicData ||
    obj.objectKind === 'chart' ||
    obj.objectKind === 'smartArt' ||
    obj.objectKind === 'media' ||
    obj.objectKind === 'picture'
  )
}

function classifyAccessibilityObject(
  obj: InventoryObject,
  chunkXml: string,
  fullText: string,
  repeatCount: number,
): { state: AccessibilityState; reason: string; decorativeScore: number; meaningfulVisualScore: number } {
  const decorativeScore = scoreDecorative(obj, chunkXml, fullText, repeatCount)
  const meaningfulVisualScore = scoreMeaningfulVisual(obj, chunkXml, fullText)

  if (obj.isMarkedDecorative) {
    return {
      state: 'alreadyAccessible',
      reason: 'marked decorative in XML',
      decorativeScore,
      meaningfulVisualScore,
    }
  }

  const ignoreCheck = shouldIgnoreFromAccessibility(obj, fullText)
  if (ignoreCheck.ignore) {
    return {
      state: 'ignoreFromAccessibility',
      reason: ignoreCheck.reason,
      decorativeScore,
      meaningfulVisualScore,
    }
  }

  if (!isMissingAltText(obj.descr, obj.title, obj.name)) {
    return {
      state: 'alreadyAccessible',
      reason: 'meaningful cNvPr descr/title',
      decorativeScore,
      meaningfulVisualScore,
    }
  }

  if (isTextReadableObject(obj, fullText, chunkXml)) {
    return {
      state: 'ignoreFromAccessibility',
      reason: 'text object — readable from shape text in reading order',
      decorativeScore,
      meaningfulVisualScore,
    }
  }

  if (obj.isGroup && obj.hasBlip && !obj.hasText) {
    return {
      state: 'needsAltText',
      reason: 'grouped visual unit with image content without meaningful alt',
      decorativeScore,
      meaningfulVisualScore,
    }
  }

  if (obj.tag === 'cxnSp' || (obj.objectKind === 'connector' && !obj.hasText && !obj.hasBlip)) {
    return {
      state: 'decorativeCandidate',
      reason: `connector/line ornament (decorativeScore=${decorativeScore})`,
      decorativeScore,
      meaningfulVisualScore,
    }
  }

  if (
    decorativeScore >= DECORATIVE_SCORE_STRICT &&
    !obj.hasText &&
    !obj.hasBlip &&
    !obj.hasGraphicData &&
    !isPictureLikeObject(obj) &&
    (obj.fromLayout || obj.objectKind === 'connector')
  ) {
    return {
      state: 'decorativeCandidate',
      reason: `decorativeScore=${decorativeScore} (layout ornament/theme shape)`,
      decorativeScore,
      meaningfulVisualScore,
    }
  }

  if (
    decorativeScore >= DECORATIVE_SCORE_THRESHOLD &&
    !obj.hasText &&
    !obj.hasBlip &&
    !obj.hasGraphicData &&
    obj.fromLayout &&
    repeatCount >= 4 &&
    (obj.objectKind === 'autoshape' || obj.objectKind === 'freeform')
  ) {
    return {
      state: 'decorativeCandidate',
      reason: `decorativeScore=${decorativeScore} (layout accent)`,
      decorativeScore,
      meaningfulVisualScore,
    }
  }

  if (isPictureLikeObject(obj)) {
    return {
      state: 'needsAltText',
      reason: 'picture/chart/media/image without meaningful alt',
      decorativeScore,
      meaningfulVisualScore,
    }
  }

  // Plain geometric shapes (autoshapes, freeforms, connectors) without image content
  // are NOT flagged by Office's accessibility checker — only pictures/charts/media need alt text.
  return {
    state: 'ignoreFromAccessibility',
    reason: 'shape without image content — not flagged by Office checker',
    decorativeScore,
    meaningfulVisualScore,
  }
}

function isMarkedDecorativeXml(chunk: string): boolean {
  if (/decorative[^>]*val="1"/i.test(chunk)) return true
  if (/\bdescr="(?:Decorative|דקורטיבי)"/i.test(chunk)) return true
  if (/<adec:decorative\b[^>]*\bval="1"/i.test(chunk)) return true
  return false
}

function hasImageContent(chunk: string): boolean {
  return (
    /<a:blip\b/i.test(chunk) ||
    /<(?:p:)?blipFill\b/i.test(chunk) ||
    /blipFill/i.test(chunk) ||
    /<(?:p:|pic:)?pic\b/i.test(chunk) ||
    /graphicData[^>]*uri="[^"]*picture/i.test(chunk)
  )
}

function hasSolidFill(chunk: string): boolean {
  return /<a:solidFill\b/i.test(chunk) || /<a:gradFill\b/i.test(chunk)
}

function hasNoFill(chunk: string): boolean {
  return /<a:noFill\b/i.test(chunk)
}

function hasOutline(chunk: string): boolean {
  return /<a:ln\b/i.test(chunk)
}

function parseCnvFromChunk(chunk: string): { descr?: string; title?: string; name?: string; id: string } {
  const cNvPr =
    chunk.match(/<(?:p:)?cNvPr\b[^>]*\/?>/i)?.[0] ??
    chunk.match(/<(?:p:)?nvPicPr[\s\S]*?<(?:p:)?cNvPr\b[^>]*\/?>/i)?.[0] ??
    ''
  return {
    id: attrFromChunk(cNvPr, 'id') ?? `anon-${hashChunk(chunk)}`,
    descr: attrFromChunk(cNvPr, 'descr') ?? chunk.match(/\bdescr="([^"]*)"/i)?.[1],
    title: attrFromChunk(cNvPr, 'title') ?? chunk.match(/\btitle="([^"]*)"/i)?.[1],
    name: attrFromChunk(cNvPr, 'name') ?? chunk.match(/\bname="([^"]*)"/i)?.[1],
  }
}

function parsePlacement(chunk: string): { x: number; y: number; cx: number; cy: number } {
  const off =
    chunk.match(/<a:off\b[^>]*\bx="(\d+)"[^>]*\by="(\d+)"/i) ??
    chunk.match(/<a:off\b[^>]*\by="(\d+)"[^>]*\bx="(\d+)"/i)
  const ext =
    chunk.match(/<a:ext\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/i) ??
    chunk.match(/<a:ext\b[^>]*\bcy="(\d+)"[^>]*\bcx="(\d+)"/i)
  let x = 0
  let y = 0
  let cx = 0
  let cy = 0
  if (off) {
    if (off[0].includes('x="')) {
      x = parseInt(off[1], 10)
      y = parseInt(off[2], 10)
    } else {
      y = parseInt(off[1], 10)
      x = parseInt(off[2], 10)
    }
  }
  if (ext) {
    cx = parseInt(ext[1], 10)
    cy = parseInt(ext[2], 10)
  }
  return { x, y, cx, cy }
}

function isHiddenShape(chunk: string, x: number, y: number): boolean {
  if (/\bhidden="1"/i.test(chunk)) return true
  if (x < -50000 || y < -50000) return true
  return false
}

function getPhKey(chunk: string): string | null {
  const type = chunk.match(/<p:ph\b[^>]*\btype="([^"]*)"/i)?.[1]
  const idx = chunk.match(/<p:ph\b[^>]*\bidx="([^"]*)"/i)?.[1] ?? ''
  if (type) return `${type}:${idx}`
  return null
}

function getBlipEmbedId(chunk: string): string | undefined {
  return chunk.match(/<a:blip\b[^>]*\br:embed="([^"]+)"/i)?.[1]
}

function getSlideRelMap(zip: AdmZip, slidePath: string): Map<string, string> {
  const map = new Map<string, string>()
  const slideName = slidePath.split('/').pop() ?? ''
  const rels = readZipEntryText(zip, `ppt/slides/_rels/${slideName}.rels`)
  if (!rels) return map
  for (const m of rels.matchAll(/<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bTarget="([^"]+)"/gi)) {
    map.set(m[1], m[2])
  }
  return map
}

function getLayoutXml(zip: AdmZip, slidePath: string): string | null {
  const slideName = slidePath.split('/').pop() ?? ''
  const rels = readZipEntryText(zip, `ppt/slides/_rels/${slideName}.rels`)
  const target = rels?.match(/Target="(\.\.\/slideLayouts\/[^"]+)"/i)?.[1]
  if (!target) return null
  return readZipEntryText(zip, target.replace(/^\.\.\//, 'ppt/'))
}

function getLayoutType(layoutXml: string | null): string | undefined {
  return layoutXml?.match(/<p:sldLayout\b[^>]*\btype="([^"]+)"/i)?.[1]
}

function shapeContainedInGroup(chunk: string, groups: string[]): boolean {
  return groups.some((g) => g.length > chunk.length && g.includes(chunk))
}

/** Aggregate visual/text signals from all shapes inside a p:grpSp (Reading Order group unit). */
function aggregateGroupContent(grpXml: string): {
  hasBlip: boolean
  hasGraphicData: boolean
  graphicUri?: string
  hasText: boolean
  text: string
  hasHyperlink: boolean
  childVisualCount: number
} {
  let hasBlip = false
  let hasGraphicData = false
  let graphicUri: string | undefined
  let hasText = false
  let text = ''
  let hasHyperlink = false
  let childVisualCount = 0
  for (const tag of ['pic', 'graphicFrame', 'sp', 'cxnSp'] as const) {
    for (const chunk of findAllTaggedChunks(grpXml, tag)) {
      if (hasImageContent(chunk)) {
        hasBlip = true
        childVisualCount++
      }
      const g = hasGraphicDataUri(chunk)
      if (g.has) {
        hasGraphicData = true
        graphicUri = g.uri
        childVisualCount++
      }
      const t = chunkText(chunk)
      if (t) {
        hasText = true
        text += (text ? ' ' : '') + t
      }
      if (hasHyperlinkInChunk(chunk)) hasHyperlink = true
    }
  }
  return { hasBlip, hasGraphicData, graphicUri, hasText, text, hasHyperlink, childVisualCount }
}

/** Collect selectable shapes from slide + layout (layout fills placeholders not overridden on slide). */
function collectAllShapeElements(slideXml: string, layoutXml: string | null): RawShapeElement[] {
  const collectFrom = (xml: string, fromLayout: boolean): RawShapeElement[] => {
    const groups = findAllTaggedChunks(xml, 'grpSp')
    const items: RawShapeElement[] = []
    for (const tag of SHAPE_TAGS) {
      if (tag === 'grpSp') continue
      for (const chunk of findAllTaggedChunks(xml, tag)) {
        const inGroup = shapeContainedInGroup(chunk, groups)
        items.push({ tag, xml: chunk, inGroup, fromLayout, groupDepth: inGroup ? 1 : 0 })
      }
    }
    for (const chunk of findAllTaggedChunks(xml, 'grpSp')) {
      if (!shapeContainedInGroup(chunk, groups)) {
        items.push({ tag: 'grpSp', xml: chunk, inGroup: false, fromLayout, groupDepth: 0 })
      }
    }
    return items
  }

  const slideShapes = collectFrom(slideXml, false)
  const slidePhKeys = new Set(slideShapes.map((s) => getPhKey(s.xml)).filter(Boolean) as string[])
  const layoutShapes = layoutXml
    ? collectFrom(layoutXml, true).filter((s) => {
        const phKey = getPhKey(s.xml)
        return !(phKey && slidePhKeys.has(phKey))
      })
    : []

  const seen = new Set<string>()
  let out: RawShapeElement[] = []
  for (const s of [...slideShapes, ...layoutShapes]) {
    const id = parseCnvFromChunk(s.xml).id
    if (seen.has(id)) {
      if (s.fromLayout) continue
      out = out.filter((o) => parseCnvFromChunk(o.xml).id !== id || o.fromLayout)
    }
    seen.add(id)
    out.push(s)
  }
  return out
}

function buildInventoryObject(
  slideNum: number,
  index: number,
  raw: RawShapeElement,
  relMap: Map<string, string>,
  repeatCount: number,
): InventoryObject {
  const cnv = parseCnvFromChunk(raw.xml)
  const ph = getPlaceholderInfo(raw.xml)
  const groupAgg = raw.tag === 'grpSp' ? aggregateGroupContent(raw.xml) : null
  const text = groupAgg?.text ?? chunkText(raw.xml)
  const { x, y, cx, cy } = parsePlacement(raw.xml)
  const embedId = getBlipEmbedId(raw.xml)
  const graphic = groupAgg
    ? { has: groupAgg.hasGraphicData, uri: groupAgg.graphicUri }
    : hasGraphicDataUri(raw.xml)
  const isCustom =
    raw.tag === 'sp' &&
    !ph.isPlaceholder &&
    (hasSolidFill(raw.xml) || hasOutline(raw.xml)) &&
    !hasImageContent(raw.xml)
  const objectKind = raw.tag === 'grpSp' ? 'group' : detectObjectKind(raw.tag, raw.xml, cnv.name)

  const obj: InventoryObject = {
    slideNum,
    index,
    tag: raw.tag,
    xmlTag: raw.tag,
    objectKind,
    cnvId: cnv.id,
    descr: cnv.descr,
    title: cnv.title,
    name: cnv.name,
    hasBlip: groupAgg?.hasBlip ?? hasImageContent(raw.xml),
    hasGraphicData: graphic.has,
    graphicUri: graphic.uri,
    hasHyperlink: groupAgg?.hasHyperlink ?? hasHyperlinkInChunk(raw.xml),
    relationshipId: embedId,
    relationshipTarget: embedId ? relMap.get(embedId) : undefined,
    hasText: text.length > 0,
    textSnippet: snippet(text, 40),
    isPlaceholder: ph.isPlaceholder,
    placeholderType: ph.type,
    isTitlePlaceholder: isTitlePhType(ph.type),
    isBodyPlaceholder: isBodyPhType(ph.type),
    isFooterPlaceholder: isFooterPhType(ph.type),
    isSlideNumberPlaceholder: isSlideNumberPhType(ph.type) || /slide number/i.test(cnv.name ?? ''),
    isDatePlaceholder: ph.type === 'dt' || /date placeholder/i.test(cnv.name ?? ''),
    isContentPlaceholder: ph.type === 'obj' || ph.type === 'clipArt' || ph.type === 'media',
    isCustomShape: isCustom,
    hasSolidFill: hasSolidFill(raw.xml),
    hasNoFill: hasNoFill(raw.xml),
    hasOutline: hasOutline(raw.xml),
    isMarkedDecorative: isMarkedDecorativeXml(raw.xml),
    isGenericName: isGenericAltText(cnv.name),
    isGroup: raw.tag === 'grpSp',
    inGroup: raw.inGroup,
    groupDepth: raw.groupDepth,
    fromLayout: raw.fromLayout,
    x,
    y,
    cx,
    cy,
    hidden: isHiddenShape(raw.xml, x, y),
    accessibilityState: 'ignoreFromAccessibility',
    classifiedAs: 'not-counted',
    decorativeScore: 0,
    meaningfulVisualScore: 0,
    reason: '',
    countedAs: 'object',
  }

  const fullText = chunkText(raw.xml)
  const result = classifyAccessibilityObject(obj, raw.xml, fullText, repeatCount)
  obj.accessibilityState = result.state
  obj.decorativeScore = result.decorativeScore
  obj.meaningfulVisualScore = result.meaningfulVisualScore
  obj.reason = result.reason
  obj.classifiedAs =
    result.state === 'alreadyAccessible' && obj.isMarkedDecorative
      ? 'decorative-marked'
      : mapAccessibilityToLegacy(result.state)
  return obj
}

/** Map grouped child to its outermost top-level p:grpSp chunk (if any). */
function findOutermostGroupChunk(
  chunk: string,
  topLevelGroups: { xml: string; id: string }[],
): string | null {
  let best: string | null = null
  for (const g of topLevelGroups) {
    if (g.xml.length > chunk.length && g.xml.includes(chunk)) {
      if (!best || g.xml.length < best.length) best = g.xml
    }
  }
  return best
}

/** Reading Order Pane: top-level objects + one entry per top-level group (not every child). */
function filterReadingOrderObjects(
  built: InventoryObject[],
  raws: RawShapeElement[],
  slideNum: number,
  relMap: Map<string, string>,
  signatureRepeat: Map<string, number>,
): { inventory: InventoryObject[]; groupChildrenSkipped: number } {
  const topLevelGroups = raws
    .filter((r) => r.tag === 'grpSp' && !r.inGroup)
    .map((r) => ({ xml: r.xml, id: parseCnvFromChunk(r.xml).id }))

  const groupXmlToObject = new Map<string, InventoryObject>()
  let groupChildrenSkipped = 0

  for (let i = 0; i < built.length; i++) {
    const obj = built[i]
    const raw = raws[i]
    if (!obj.inGroup) continue
    const outerXml = findOutermostGroupChunk(raw.xml, topLevelGroups)
    if (!outerXml) {
      groupChildrenSkipped++
      continue
    }
    if (!groupXmlToObject.has(outerXml)) {
      const gi = topLevelGroups.findIndex((g) => g.xml === outerXml)
      const gRaw: RawShapeElement = {
        tag: 'grpSp',
        xml: outerXml,
        inGroup: false,
        fromLayout: raw.fromLayout,
        groupDepth: 0,
      }
      groupXmlToObject.set(
        outerXml,
        buildInventoryObject(slideNum, 9000 + gi, gRaw, relMap, signatureRepeat.get(buildShapeSignatureFromRaw(gRaw)) ?? 0),
      )
    }
    groupChildrenSkipped++
  }

  const topLevelSingles = built.filter((o, i) => !raws[i].inGroup && raws[i].tag !== 'grpSp')
  // Only include top-level grpSp from raws if they are NOT already represented in groupXmlToObject
  // (groups with inGroup children are already synthesized there — including them again causes duplication)
  const topLevelGrpFromRaws = built.filter(
    (o, i) => raws[i].tag === 'grpSp' && !raws[i].inGroup && !groupXmlToObject.has(raws[i].xml),
  )
  const merged = [...topLevelSingles, ...topLevelGrpFromRaws, ...groupXmlToObject.values()]
  const inventory = merged.map((o, i) => ({ ...o, index: i + 1 }))
  return { inventory, groupChildrenSkipped }
}

function buildShapeSignatureFromRaw(raw: RawShapeElement): string {
  const cnv = parseCnvFromChunk(raw.xml)
  const pl = parsePlacement(raw.xml)
  const prst = raw.xml.match(/<a:prstGeom\b[^>]*\bprst="([^"]+)"/i)?.[1] ?? ''
  return `${cnv.name ?? ''}|${prst}|${pl.cx}|${pl.cy}|${raw.fromLayout ? 'L' : 'S'}`
}

function buildSlideInventory(
  slideNum: number,
  slideXml: string,
  layoutXml: string | null,
  relMap: Map<string, string>,
  signatureRepeat: Map<string, number>,
): { inventory: InventoryObject[]; groupChildrenSkipped: number } {
  const raws = collectAllShapeElements(slideXml, layoutXml)
  const built = raws.map((raw, i) => {
    const sig = buildShapeSignatureFromRaw(raw)
    return buildInventoryObject(slideNum, i + 1, raw, relMap, signatureRepeat.get(sig) ?? 0)
  })
  return filterReadingOrderObjects(built, raws, slideNum, relMap, signatureRepeat)
}

function buildDeckSignatureRepeat(slidePaths: string[], zip: AdmZip): Map<string, number> {
  const counts = new Map<string, number>()
  for (const slidePath of slidePaths) {
    const slideXml = readZipEntryText(zip, slidePath)
    if (!slideXml) continue
    const layoutXml = getLayoutXml(zip, slidePath)
    for (const raw of collectAllShapeElements(slideXml, layoutXml)) {
      const cnv = parseCnvFromChunk(raw.xml)
      const pl = parsePlacement(raw.xml)
      const sig = `${cnv.name ?? ''}|${raw.xml.match(/<a:prstGeom\b[^>]*\bprst="([^"]+)"/i)?.[1] ?? ''}|${pl.cx}|${pl.cy}|${raw.fromLayout ? 'L' : 'S'}`
      counts.set(sig, (counts.get(sig) ?? 0) + 1)
    }
  }
  return counts
}

interface SlideClassificationSummary {
  slide: number
  topLevelReadingObjects: number
  ignoredPlaceholders: number
  textObjects: number
  pictureObjects: number
  groupedObjectsCounted: number
  groupChildrenSkipped: number
  missingAltText: number
  decorativeCandidates: number
  contrastIssues: number
  readingOrderRisk: boolean
  ignored: number
  alreadyAccessible: number
  byType: Record<string, number>
}

function summarizeSlideClassification(
  slideNum: number,
  inventory: InventoryObject[],
  groupChildrenSkipped: number,
  readingOrderRisk: boolean,
): SlideClassificationSummary {
  const byType: Record<string, number> = {}
  let ignored = 0
  let needsAltText = 0
  let decorativeCandidates = 0
  let alreadyAccessible = 0
  let ignoredPlaceholders = 0
  let textObjects = 0
  let pictureObjects = 0
  for (const o of inventory) {
    byType[o.objectKind] = (byType[o.objectKind] ?? 0) + 1
    if (isPictureLikeObject(o)) pictureObjects++
    if (o.hasText && o.accessibilityState === 'ignoreFromAccessibility') textObjects++
    if (
      o.isPlaceholder ||
      o.isTitlePlaceholder ||
      o.isBodyPlaceholder ||
      o.isFooterPlaceholder
    ) {
      if (o.accessibilityState === 'ignoreFromAccessibility') ignoredPlaceholders++
    }
    switch (o.accessibilityState) {
      case 'ignoreFromAccessibility':
        ignored++
        break
      case 'needsAltText':
        needsAltText++
        break
      case 'decorativeCandidate':
        decorativeCandidates++
        break
      case 'alreadyAccessible':
        alreadyAccessible++
        break
    }
  }
  return {
    slide: slideNum,
    topLevelReadingObjects: inventory.length,
    ignoredPlaceholders,
    textObjects,
    pictureObjects,
    groupedObjectsCounted: inventory.filter((o) => o.isGroup || o.tag === 'grpSp').length,
    groupChildrenSkipped,
    missingAltText: needsAltText,
    decorativeCandidates,
    contrastIssues: 0,
    readingOrderRisk,
    ignored,
    alreadyAccessible,
    byType,
  }
}

function logObjectInventory(slideNum: number, inventory: InventoryObject[]): void {
  console.log('[PPTX OBJECT INVENTORY]', {
    slideNumber: slideNum,
    objectCount: inventory.length,
    objects: inventory.map((o) => ({
      objectIndex: o.index,
      xmlTag: o.xmlTag,
      objectType: o.objectKind,
      name: o.name ? snippet(o.name, 30) : undefined,
      hasBlip: o.hasBlip,
      hasGraphicData: o.hasGraphicData,
      hasText: o.hasText,
      fromLayout: o.fromLayout,
      inGroup: o.inGroup,
      accessibilityState: o.accessibilityState,
      decorativeScore: o.decorativeScore,
      meaningfulVisualScore: o.meaningfulVisualScore,
      reason: o.reason,
    })),
  })
}

function extractSlideTitle(
  slideNum: number,
  slideXml: string,
  layoutXml: string | null,
): SlideTitleInfo {
  const titleTexts: string[] = []
  const hiddenTitleTexts: string[] = []
  let titlePlaceholdersFound = 0
  let hasTitlePlaceholder = false

  const scanXml = (xml: string, source: string) => {
    for (const sp of findAllTaggedChunks(xml, 'sp')) {
      const ph = getPlaceholderInfo(sp)
      if (!isTitlePhType(ph.type)) continue
      titlePlaceholdersFound++
      hasTitlePlaceholder = true
      const t = cleanTitleText(chunkText(sp))
      if (!t) continue
      const pl = parsePlacement(sp)
      const hidden = isHiddenShape(sp, pl.x, pl.y)
      if (hidden) hiddenTitleTexts.push(t)
      else titleTexts.push(t)
    }
  }

  scanXml(slideXml, 'slide')
  if (layoutXml) scanXml(layoutXml, 'layout')

  const cSldName = cleanTitleText(slideXml.match(/<(?:p:)?cSld\b[^>]*\bname="([^"]*)"/i)?.[1] ?? '')
  const layoutType = getLayoutType(layoutXml)
  const isTitleSlideLayout = layoutType === 'title'

  const meaningfulFromPh = [...titleTexts, ...hiddenTitleTexts].filter((t) => !isBoilerplateTitleText(t))
  const cSldCandidate =
    cSldName && !/^slide\s*\d+$/i.test(cSldName) && !/^שקופית\s*\d+$/i.test(cSldName) && !isBoilerplateTitleText(cSldName)
      ? cSldName
      : ''
  const chosenTitle =
    [...meaningfulFromPh, cSldCandidate].filter(Boolean).sort((a, b) => b.length - a.length)[0] ?? ''
  const normalizedTitle = normalizeTitle(chosenTitle)
  const onlyBoilerplateInPh =
    titleTexts.length + hiddenTitleTexts.length > 0 &&
    meaningfulFromPh.length === 0

  let isMissingTitle = false
  let missingTitleReason = ''

  if (normalizedTitle.length > 0) {
    isMissingTitle = false
    missingTitleReason = 'meaningful title text in placeholder or metadata'
  } else if (slideNum === 1 && hasTitlePlaceholder) {
    isMissingTitle = false
    missingTitleReason = 'opening slide has title placeholder structure (PowerPoint title-slide semantics)'
  } else if (isTitleSlideLayout && titlePlaceholdersFound > 0) {
    isMissingTitle = false
    missingTitleReason = 'title slide layout with title placeholder present'
  } else if (onlyBoilerplateInPh) {
    isMissingTitle = true
    missingTitleReason = 'title placeholder contains only default placeholder text (e.g. Title Text)'
  } else if (hasTitlePlaceholder) {
    isMissingTitle = true
    missingTitleReason = 'title/ctrTitle placeholder exists but is empty'
  } else if (!hasTitlePlaceholder && !isTitleSlideLayout) {
    isMissingTitle = true
    missingTitleReason = 'no title placeholder on slide or layout'
  } else {
    isMissingTitle = false
    missingTitleReason = 'title slide layout without required separate title text'
  }

  const titleSource =
    titleTexts.length > 0
      ? 'title-placeholder'
      : hiddenTitleTexts.length > 0
        ? 'hidden-title-placeholder'
        : cSldName
          ? 'cSld-name'
          : 'none'

  const debug = {
    slideNumber: slideNum,
    titlePlaceholdersFound,
    titleTexts,
    hiddenTitleTexts,
    candidateTitleShapes: titlePlaceholdersFound,
    chosenTitle: snippet(chosenTitle, 40),
    normalizedTitle: snippet(normalizedTitle, 40),
    layoutType,
    missingTitle: isMissingTitle,
    reason: missingTitleReason,
  }
  console.log('[PPTX TITLE DEBUG]', debug)

  return {
    slideNum,
    hasTitlePlaceholder,
    titlePlaceholdersFound,
    titleTexts,
    hiddenTitleTexts,
    chosenTitle,
    normalizedTitle,
    titleSource,
    layoutType,
    isMissingTitle,
    missingTitleReason,
  }
}

interface ReadingOrderResult {
  score: number
  risk: boolean
  objectCount: number
  textBoxCount: number
  pictureCount: number
  customShapeCount: number
  groupCount: number
  inversions: number
}

function computeReadingOrderScore(
  slideXml: string,
  inventory: InventoryObject[],
  titleInfo: SlideTitleInfo,
): ReadingOrderResult {
  const content = inventory.filter(
    (o) =>
      !o.fromLayout &&
      o.accessibilityState !== 'ignoreFromAccessibility' &&
      o.accessibilityState !== 'alreadyAccessible' &&
      (o.hasText || o.hasBlip || o.isCustomShape || o.tag === 'pic' || o.hasGraphicData),
  )

  const textBoxCount = content.filter(
    (o) => o.hasText && !o.isTitlePlaceholder && !o.isBodyPlaceholder,
  ).length
  const pictureCount = content.filter((o) => o.hasBlip || o.tag === 'pic').length
  const customShapeCount = content.filter((o) => o.isCustomShape && o.hasText).length
  const groupCount = (slideXml.match(/<(?:p:)?grpSp\b/gi) ?? []).length
  const bodyPhCount = content.filter((o) => o.isBodyPlaceholder && o.hasText).length

  const placed = content.map((o, docIndex) => ({ docIndex, x: o.x, y: o.y }))
  let inversions = 0
  if (placed.length >= 3) {
    const visual = [...placed].sort((a, b) => a.y - b.y || a.x - b.x)
    const order = new Map(visual.map((p, i) => [p.docIndex, i]))
    for (let i = 0; i < placed.length; i++) {
      for (let j = i + 1; j < placed.length; j++) {
        if ((order.get(placed[i].docIndex) ?? 0) > (order.get(placed[j].docIndex) ?? 0)) inversions++
      }
    }
  }

  const hasRtl = /[\u0590-\u05FF]/.test(slideXml)
  const hasTitle = titleInfo.normalizedTitle.length > 0
  const mixedContent = (textBoxCount + bodyPhCount >= 1 ? 1 : 0) + (pictureCount >= 1 ? 1 : 0) + (customShapeCount >= 1 ? 1 : 0)

  let score = 0
  if (content.length >= 3) score += 1
  if (content.length >= 5) score += 1
  if (content.length >= 8) score += 1
  if (textBoxCount + bodyPhCount >= 2) score += 1
  if (textBoxCount + bodyPhCount >= 1 && customShapeCount >= 1) score += 2
  if (pictureCount >= 1 && (textBoxCount + bodyPhCount) >= 1) score += 2
  if (pictureCount >= 2) score += 1
  if (customShapeCount >= 1 && hasTitle) score += 1
  if (groupCount >= 1 && content.length >= 4) score += 1
  if (mixedContent >= 2 && content.length >= 4) score += 2
  if (inversions >= 1) score += 1
  if (inversions >= 2) score += 2
  if (hasRtl && content.length >= 4) score += 1
  if (hasRtl && pictureCount >= 1 && textBoxCount >= 1) score += 1
  if (bodyPhCount >= 1 && textBoxCount >= 1) score += 2
  if (bodyPhCount >= 1 && customShapeCount >= 1) score += 1
  if (content.length >= 2 && textBoxCount + bodyPhCount >= 2 && !hasTitle) score += 1
  const paragraphCount = (slideXml.match(/<a:p[\s>]/gi) ?? []).length
  const textRunCount = (slideXml.match(/<a:t>/gi) ?? []).length
  if (hasTitle && bodyPhCount >= 1 && (paragraphCount >= 4 || textRunCount >= 8)) score += 3
  if (hasRtl && bodyPhCount >= 1 && (paragraphCount >= 4 || textRunCount >= 8)) score += 2
  if (bodyPhCount >= 1 && textRunCount >= 10) score += 2

  return {
    score,
    risk: score >= 5,
    objectCount: content.length,
    textBoxCount: textBoxCount + bodyPhCount,
    pictureCount,
    customShapeCount,
    groupCount,
    inversions,
  }
}

function inventoryLocation(slideNum: number, obj: InventoryObject): string {
  if (obj.tag === 'pic' || obj.objectKind === 'picture' || (obj.hasBlip && !obj.isCustomShape)) {
    return `שקופית ${slideNum} / תמונה ${obj.index}`
  }
  if (obj.objectKind === 'chart' || obj.objectKind === 'smartArt' || obj.objectKind === 'graphicFrame') {
    return `שקופית ${slideNum} / אובייקט גרפי ${obj.index}`
  }
  return `שקופית ${slideNum} / אובייקט חזותי ${obj.index}`
}

function detectFontIssues(xml: string): number {
  let small = 0
  for (const chunk of findAllTaggedChunks(xml, 'sp')) {
    const isTitlePh = /type="(?:title|ctrTitle)"/i.test(chunk)
    for (const run of chunk.matchAll(/<a:r[\s\S]*?<\/a:r>/gi)) {
      const sz = parseInt(run[0].match(/sz="(\d+)"/i)?.[1] ?? '0', 10)
      if (sz <= 0) continue
      const pt = ptFromSz(sz)
      if (isTitlePh && pt < 24) small++
      else if (!isTitlePh && pt < 18) small++
    }
  }
  return small
}

/** Disabled: theme/background inference is unreliable; PowerPoint passed contrast on SE-B-8. */
function detectContrastIssues(_xml: string): number {
  return 0
}

export function analyzePptx(
  filePath: string,
  fileSize: number,
): { drafts: IssueDraft[]; signals: ExtractedSignals } {
  console.log('[PPTX SCANNER IS RUNNING]', {
    build: PPTX_SCANNER_BUILD,
    filePath,
    fileSize,
  })

  const zip = readZip(filePath)
  const slidePaths = listZipEntries(zip, /^ppt\/slides\/slide\d+\.xml$/i)
  const signatureRepeat = buildDeckSignatureRepeat(slidePaths, zip)
  const drafts: IssueDraft[] = []
  const classificationSummaries: SlideClassificationSummary[] = []

  const missingTitleSlides: number[] = []
  const duplicateTitleSlides: number[] = []
  const readingOrderRiskSlides: number[] = []
  const missingAltLocations: string[] = []
  const decorativeCandidateLocations: string[] = []
  let missingAltCount = 0
  let decorativeCandidateCount = 0

  const slideTitles: SlideTitleInfo[] = []
  const inventorySummaries: SlideSlideSummary[] = []

  for (const slidePath of slidePaths) {
    const slideXml = readZipEntryText(zip, slidePath)
    if (!slideXml) continue
    const slideNum = parseInt(slidePath.match(/slide(\d+)\.xml/i)?.[1] ?? '0', 10)
    const layoutXml = getLayoutXml(zip, slidePath)
    const relMap = getSlideRelMap(zip, slidePath)

    const titleInfo = extractSlideTitle(slideNum, slideXml, layoutXml)
    slideTitles.push(titleInfo)

    const { inventory, groupChildrenSkipped } = buildSlideInventory(
      slideNum,
      slideXml,
      layoutXml,
      relMap,
      signatureRepeat,
    )
    logObjectInventory(slideNum, inventory)
    const readingOrder = computeReadingOrderScore(slideXml, inventory, titleInfo)
    const slideClass = summarizeSlideClassification(
      slideNum,
      inventory,
      groupChildrenSkipped,
      readingOrder.risk,
    )
    classificationSummaries.push(slideClass)
    console.log('[ACCESSIBILITY OBJECT CLASSIFICATION]', slideClass)
    console.log('[OBJECT CLASSIFICATION SUMMARY]', slideClass)

    let slideMissingAlt = 0
    let slideDecorative = 0

    if (titleInfo.isMissingTitle) {
      missingTitleSlides.push(slideNum)
      drafts.push({
        id: `slide-without-title-${slideNum}`,
        wcagKey: 'page-titled-headings',
        title: OFFICE_TITLES.slideWithoutTitle,
        severity: 'Medium',
        category: CAT_DOCUMENT_STRUCTURE,
        affectedUsers: ['משתמשי קורא מסך'],
        impact: 'בשקופית זו לא זוהתה כותרת שקופית תקינה לפי מבנה PowerPoint.',
        recommendation: 'הוסיפו כותרת ייחודית בשדה הכותרת של השקופית (Title placeholder).',
        location: `שקופית ${slideNum}`,
        confidence: 'High',
      })
    }

    for (const obj of inventory) {
      const loc = inventoryLocation(slideNum, obj)

      if (obj.accessibilityState === 'decorativeCandidate') {
        slideDecorative++
        decorativeCandidateCount++
        decorativeCandidateLocations.push(loc)
        drafts.push({
          id: `decorative-candidate-${slideNum}-${obj.cnvId}`,
          wcagKey: 'non-text-content',
          title: OFFICE_TITLES.decorativeCandidates,
          severity: 'Low',
          category: CAT_MEDIA,
          affectedUsers: ['משתמשי קורא מסך'],
          impact: 'צורה דקורטיבית ללא טקסט חלופי — ניתן לסמן כ־Decorative.',
          recommendation:
            'אם האובייקט משמש לקישוט בלבד, סמנו אותו כ־Decorative ב־PowerPoint. אם הוא מעביר מידע, הוסיפו טקסט חלופי.',
          location: loc,
          confidence: 'Medium',
          issueTier: 'quickFix',
        })
        continue
      }

      if (obj.accessibilityState === 'needsAltText') {
        slideMissingAlt++
        missingAltCount++
        missingAltLocations.push(loc)
        drafts.push({
          id: `missing-alt-${slideNum}-${obj.cnvId}`,
          wcagKey: 'non-text-content',
          title: OFFICE_TITLES.missingAltText,
          severity: 'High',
          category: CAT_MEDIA,
          affectedUsers: ['משתמשים עיוורים', 'משתמשי קורא מסך'],
          impact: 'לאובייקט חזותי זה חסר טקסט חלופי (title/descr) משמעותי.',
          recommendation:
            'ב-PowerPoint: לחצו ימני על האובייקט → Edit Alt Text והוסיפו תיאור משמעותי.',
          location: loc,
          confidence:
            obj.tag === 'pic' || obj.objectKind === 'chart' || obj.objectKind === 'media' || obj.hasGraphicData
              ? 'High'
              : 'Medium',
        })
      }
    }

    if (readingOrder.risk) {
      readingOrderRiskSlides.push(slideNum)
      drafts.push({
        id: `reading-order-${slideNum}`,
        wcagKey: 'meaningful-sequence',
        title: OFFICE_TITLES.readingOrderCheck,
        severity: 'Medium',
        category: CAT_DOCUMENT_STRUCTURE,
        affectedUsers: ['משתמשי קורא מסך'],
        impact: `רכיבים: ${readingOrder.objectCount}, תיבות טקסט: ${readingOrder.textBoxCount}, תמונות: ${readingOrder.pictureCount}, צורות מותאמות: ${readingOrder.customShapeCount}. מומלץ לבדוק סדר קריאה.`,
        recommendation:
          'בדקו את סדר הקריאה בחלונית Accessibility → Reading Order Pane וודאו שהתוכן נקרא בסדר הגיוני.',
        location: `שקופית ${slideNum}`,
        confidence: 'Medium',
      })
    }

    inventorySummaries.push({
      slide: slideNum,
      chosenTitle: titleInfo.chosenTitle,
      missingTitle: titleInfo.isMissingTitle,
      duplicateTitle: false,
      visualObjectsCount: inventory.filter(
        (o) => o.accessibilityState === 'needsAltText' || o.accessibilityState === 'alreadyAccessible',
      ).length,
      customShapesWithText: inventory.filter((o) => o.isCustomShape && o.hasText).length,
      imageLikeObjects: inventory.filter((o) => o.hasBlip || o.tag === 'pic').length,
      missingAltObjects: slideMissingAlt,
      decorativeCandidates: slideDecorative,
      readingOrderScore: readingOrder.score,
      readingOrderRisk: readingOrder.risk,
    })

    for (const m of slideXml.matchAll(/<a:t>([^<]*)<\/a:t>/g)) {
      if (isUnclearLink(m[1])) {
        drafts.push({
          id: `unclear-link-${slideNum}-${m[1].slice(0, 8)}`,
          wcagKey: 'link-purpose',
          title: OFFICE_TITLES.unclearHyperlink,
          severity: 'Medium',
          category: CAT_DOCUMENT_STRUCTURE,
          affectedUsers: ['משתמשי קורא מסך'],
          impact: `בשקופית ${slideNum}: "${snippet(m[1])}".`,
          recommendation: 'החליפו בטקסט קישור תיאורי.',
          location: `שקופית ${slideNum}`,
          confidence: 'High',
        })
      }
    }

    const smallFont = detectFontIssues(slideXml)
    if (smallFont >= 1) {
      drafts.push({
        id: `small-font-${slideNum}`,
        wcagKey: 'resize-text',
        title: OFFICE_TITLES.smallFont,
        severity: 'Low',
        category: CAT_COLOR_CONTRAST,
        affectedUsers: ['משתמשים עם לקות ראייה'],
        impact: `בשקופית ${slideNum}: גופן קטן מ-18pt (${smallFont} מופעים).`,
        recommendation: 'הגדילו גופן גוף ל-18pt לפחות.',
        location: `שקופית ${slideNum}`,
        confidence: 'High',
      })
    }

    detectContrastIssues(slideXml)
  }

  const titleGroups = new Map<string, number[]>()
  for (const st of slideTitles) {
    if (!st.chosenTitle || isBoilerplateTitleText(st.chosenTitle)) continue
    const key = normalizeTitleForDuplicate(st.chosenTitle)
    if (!key || key.length < 3) continue
    const list = titleGroups.get(key) ?? []
    list.push(st.slideNum)
    titleGroups.set(key, list)
  }

  for (const [norm, slideNums] of titleGroups) {
    if (slideNums.length < 2) continue
    const sorted = [...slideNums].sort((a, b) => a - b)
    for (const slideNum of sorted.slice(1)) {
      duplicateTitleSlides.push(slideNum)
      const otherSlides = sorted.filter((n) => n !== slideNum).join(', ')
      drafts.push({
        id: `duplicate-title-${slideNum}`,
        wcagKey: 'page-titled-headings',
        title: OFFICE_TITLES.duplicateSlideTitle,
        severity: 'Medium',
        category: CAT_DOCUMENT_STRUCTURE,
        affectedUsers: ['משתמשי קורא מסך'],
        impact: `כותרת זו מופיעה גם בשקופית ${otherSlides}.`,
        recommendation: 'העניקו לכל שקופית כותרת ייחודית.',
        location: `שקופית ${slideNum}`,
        confidence: 'High',
      })
      const sum = inventorySummaries.find((s) => s.slide === slideNum)
      if (sum) sum.duplicateTitle = true
    }
    console.log('[PPTX duplicate-title group]', { normalized: snippet(norm, 40), slides: sorted })
  }

  const comparison = {
    missingAltText: { count: missingAltCount, locations: [...missingAltLocations] },
    missingSlideTitle: {
      count: missingTitleSlides.length,
      locations: missingTitleSlides.map((n) => `שקופית ${n}`),
    },
    duplicateSlideTitle: {
      count: duplicateTitleSlides.length,
      locations: duplicateTitleSlides.map((n) => `שקופית ${n}`),
    },
    readingOrder: {
      count: readingOrderRiskSlides.length,
      locations: readingOrderRiskSlides.map((n) => `שקופית ${n}`),
    },
    decorativeCandidates: {
      count: decorativeCandidateCount,
      locations: [...decorativeCandidateLocations],
    },
  }

  const contrastIssues = 0
  const mainIssueTotal =
    missingAltCount +
    missingTitleSlides.length +
    duplicateTitleSlides.length +
    readingOrderRiskSlides.length +
    contrastIssues

  const officeLikeCounts = {
    missingAltText: missingAltCount,
    quickFixDecorative: decorativeCandidateCount,
    missingSlideTitle: missingTitleSlides.length,
    duplicateSlideTitle: duplicateTitleSlides.length,
    readingOrder: readingOrderRiskSlides.length,
    contrastIssues,
    tableIssues: 0,
    mainIssueTotal,
    quickFixTotal: decorativeCandidateCount,
  }

  console.log('[OFFICE-LIKE COUNTS]', officeLikeCounts)
  console.log('[POWERPOINT-LIKE COUNTS]', officeLikeCounts)
  console.log('[POWERPOINT COMPARISON SUMMARY]', comparison)
  console.log('[PPTX OBJECT INVENTORY SUMMARY]', inventorySummaries)
  console.log('[ACCESSIBILITY CLASSIFICATION DECK SUMMARY]', {
    mainIssueTotal,
    quickFixTotal: decorativeCandidateCount,
    groupChildrenSkippedTotal: classificationSummaries.reduce((s, c) => s + c.groupChildrenSkipped, 0),
    slides: classificationSummaries,
  })

  const signals: ExtractedSignals = {
    fileType: 'pptx',
    fileSize,
    slideCount: slidePaths.length,
    imageCount: inventorySummaries.reduce((s, i) => s + i.imageLikeObjects, 0),
    missingTitleSlides,
    titleCount: slideTitles.filter((t) => t.normalizedTitle).length,
    duplicateTitleSlides,
    imageLikeObjectCount: inventorySummaries.reduce((s, i) => s + i.imageLikeObjects, 0),
    missingAltCount,
    decorativeCandidateCount,
    readingOrderRiskSlides,
    mainIssueTotal,
    quickFixTotal: decorativeCandidateCount,
  }

  return { drafts, signals }
}
