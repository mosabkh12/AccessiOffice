/** Safe diagnostic logging for PPTX scans (no long text). */

export function snippet(text: string, max = 40): string {
  const t = text.replace(/\s+/g, ' ').trim()
  if (t.length <= max) return t
  return `${t.slice(0, max - 1)}…`
}

export type ObjectClassification =
  | 'missing-alt'
  | 'decorative-candidate'
  | 'decorative-marked'
  | 'skipped-title-placeholder'
  | 'skipped-background'
  | 'skipped-not-visual'
  | 'skipped-layout-chrome'
  | 'has-alt'
  | 'not-counted'

export interface PptxObjectDiagnostic {
  slide?: number
  index: number
  objectType: string
  xmlTag: string
  cNvPrId: string
  name?: string
  titleAttr?: string
  descrAttr?: string
  hasBlip: boolean
  relationshipTarget?: string
  inGroup: boolean
  isPictureLike: boolean
  isShapeLike: boolean
  isGraphicFrame: boolean
  isDecorative: boolean
  isGenericName: boolean
  hasMeaningfulAlt?: boolean
  needsAltText: boolean
  classifiedAs: ObjectClassification
  reason: string
  x: number
  y: number
}

export interface PptxSlideDiagnostic {
  slideNumber: number
  title: {
    hasTitlePlaceholder: boolean
    titlePlaceholdersFound?: number
    titleText: string
    titlePlaceholderEmpty: boolean
    titleSource: string
    titleIsHidden?: boolean
    missingTitle?: boolean
    isMissingTitle: boolean
  }
  objects: PptxObjectDiagnostic[]
  counts: {
    pictureLikeCount: number
    missingAltCount: number
    decorativeCandidateCount: number
    textBoxCount: number
    shapeCount: number
    groupCount: number
    graphicFrameCount: number
    readingOrderObjectCount: number
    readingOrderScore: number
    readingOrderRisk: boolean
  }
}

export interface PptxDeckDiagnostic {
  slideCount: number
  missingTitleSlides: number[]
  duplicateTitleSlides: number[]
  normalizedTitles: Array<{
    slideNumber: number
    rawTitle: string
    normalizedTitle: string
    titleSource: string
  }>
  missingAltLocations: string[]
  decorativeCandidateLocations: string[]
  readingOrderRiskSlides: number[]
  slides: PptxSlideDiagnostic[]
}

export function logPptxDiagnostics(deck: PptxDeckDiagnostic, fileName?: string): void {
  console.log('[AccessiOffice PPTX diagnostic]', {
    fileName,
    slideCount: deck.slideCount,
    missingTitleSlides: deck.missingTitleSlides,
    duplicateTitleSlides: deck.duplicateTitleSlides,
    missingAltLocations: deck.missingAltLocations,
    decorativeCandidateLocations: deck.decorativeCandidateLocations,
    readingOrderRiskSlides: deck.readingOrderRiskSlides,
    normalizedTitles: deck.normalizedTitles.map((t) => ({
      slide: t.slideNumber,
      raw: snippet(t.rawTitle),
      normalized: snippet(t.normalizedTitle),
      source: t.titleSource,
    })),
  })

  for (const slide of deck.slides) {
    console.log('[AccessiOffice PPTX slide]', JSON.stringify({
      slideNumber: slide.slideNumber,
      title: {
        ...slide.title,
        titleText: snippet(slide.title.titleText),
      },
      counts: slide.counts,
      objects: slide.objects.map((o) => ({
        slide: o.slide ?? slide.slideNumber,
        objectIndex: o.index,
        xmlTag: o.xmlTag,
        objectType: o.objectType,
        name: o.name ? snippet(o.name, 30) : undefined,
        title: o.titleAttr ? snippet(o.titleAttr, 30) : undefined,
        descr: o.descrAttr ? snippet(o.descrAttr, 30) : undefined,
        hasBlip: o.hasBlip,
        relationshipTarget: o.relationshipTarget,
        inGroup: o.inGroup,
        isGenericName: o.isGenericName,
        hasMeaningfulAlt: o.hasMeaningfulAlt,
        isDecorative: o.isDecorative,
        classifiedAs: o.classifiedAs,
        reason: o.reason,
      })),
    }))
  }
}
