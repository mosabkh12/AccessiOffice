import ExcelJS from 'exceljs'
import type { ExtractedSignals, IssueDraft } from './scannerShared.js'
import {
  CAT_COLOR_CONTRAST,
  CAT_DOCUMENT_STRUCTURE,
  CAT_TABLES,
  OFFICE_TITLES,
} from './officeCategories.js'
import { contrastRatio, isGenericSheetName } from './ooxml.utils.js'

function argbToHex(argb?: string): string | null {
  if (!argb || argb.length < 6) return null
  return argb.slice(-6)
}

function cellContrastRisk(cell: ExcelJS.Cell): boolean {
  const fg = cell.font?.color && 'argb' in cell.font.color ? argbToHex(cell.font.color.argb) : null
  const fill = cell.fill
  let bg: string | null = null
  if (fill && fill.type === 'pattern' && fill.fgColor && 'argb' in fill.fgColor) {
    bg = argbToHex(fill.fgColor.argb)
  }
  if (fg && bg) {
    const ratio = contrastRatio(fg, bg)
    return ratio != null && ratio < 4.5
  }
  return false
}

function rowLooksLikeHeader(row: ExcelJS.Row): boolean {
  const values: string[] = []
  row.eachCell({ includeEmpty: false }, (cell) => {
    const v = cell.text?.trim()
    if (v) values.push(v)
  })
  if (values.length < 2) return false
  const allNumeric = values.every((v) => /^[\d.,%-]+$/.test(v))
  if (allNumeric) return false
  return values.every((v) => v.length > 0 && v.length < 80)
}

export async function analyzeXlsx(
  filePath: string,
  fileSize: number,
): Promise<{ drafts: IssueDraft[]; signals: ExtractedSignals }> {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(filePath)

  const drafts: IssueDraft[] = []
  const genericSheetNames: string[] = []
  let mergedCellCount = 0

  const signals: ExtractedSignals = {
    fileType: 'xlsx',
    fileSize,
    sheetCount: workbook.worksheets.length,
    mergedCellCount: 0,
    genericSheetNames: [],
  }

  workbook.eachSheet((sheet) => {
    const name = sheet.name
    const loc = `גיליון "${name}"`

    if (isGenericSheetName(name)) {
      genericSheetNames.push(name)
      drafts.push({
        id: `sheet-name-${name}`,
        wcagKey: 'headings-labels',
        title: OFFICE_TITLES.genericSheetName,
        severity: 'Medium',
        category: CAT_DOCUMENT_STRUCTURE,
        affectedUsers: ['משתמשי קורא מסך'],
        impact: `שם "${name}" אינו מתאר את תוכן הגיליון.`,
        recommendation: 'שנו לשם תיאורי (למשל "ציוני סטודנטים").',
        location: loc,
        confidence: 'High',
      })
    }

    const merges = sheet.model.merges ?? []
    if (merges.length > 0) {
      mergedCellCount += merges.length
      drafts.push({
        id: `merged-${name}`,
        wcagKey: 'info-relationships',
        title: OFFICE_TITLES.mergedOrSplitCells,
        severity: 'High',
        category: CAT_TABLES,
        affectedUsers: ['משתמשי קורא מסך'],
        impact: `ב"${name}" זוהו ${merges.length} אזורי מיזוג.`,
        recommendation: 'בטלו מיזוג תאים והשתמשו ביישור/גבולות.',
        location: loc,
        confidence: 'High',
      })
    }

    const dim = sheet.dimensions
    if (!dim || dim.bottom < dim.top) return

    const startRow = dim.top
    const endRow = dim.bottom
    const startCol = dim.left
    const endCol = dim.right
    const firstRow = sheet.getRow(startRow)
    const hasDataRows = endRow > startRow

    if (hasDataRows && !rowLooksLikeHeader(firstRow)) {
      const filled = firstRow.cellCount
      if (filled >= 2) {
        drafts.push({
          id: `header-${name}`,
          wcagKey: 'info-relationships',
          title: OFFICE_TITLES.missingTableHeader,
          severity: 'High',
          category: CAT_TABLES,
          affectedUsers: ['משתמשי קורא מסך'],
          impact: `ב"${name}" השורה הראשונה אינה כותרת טקסטואלית ברורה.`,
          recommendation: 'הוסיפו שורת כותרת עם שמות עמודה תיאוריים.',
          location: `${loc} · שורה ${startRow}`,
          confidence: 'High',
        })
      }
    }

    const headerTexts: string[] = []
    if (hasDataRows) {
      firstRow.eachCell({ includeEmpty: true }, (cell, col) => {
        if (col <= endCol) headerTexts[col] = cell.text?.trim() ?? ''
      })
    }

    let emptyInside = 0
    let totalCells = 0
    let contrastIssues = 0

    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        const cell = sheet.getCell(r, c)
        const val = cell.text?.trim() ?? ''
        if (r > startRow || c > startCol) {
          totalCells++
          if (!val) emptyInside++
        }
        if (val && cellContrastRisk(cell)) contrastIssues++
      }
    }

    const columnsWithoutHeader: number[] = []
    if (hasDataRows) {
      for (let c = startCol; c <= endCol; c++) {
        const hasDataInCol = sheet.getColumn(c).values?.some(
          (v, i) => i > startRow && v != null && String(v).trim() !== '',
        )
        if (hasDataInCol && !headerTexts[c]) columnsWithoutHeader.push(c)
      }
    }

    if (columnsWithoutHeader.length > 0 && hasDataRows) {
      drafts.push({
        id: `columns-no-header-${name}`,
        wcagKey: 'info-relationships',
        title: OFFICE_TITLES.columnsWithoutHeader,
        severity: 'Medium',
        category: CAT_TABLES,
        affectedUsers: ['משתמשי קורא מסך'],
        impact: `ב"${name}" עמודות ${columnsWithoutHeader.join(', ')} ללא כותרת.`,
        recommendation: 'הוסיפו כותרות עמודה בשורה הראשונה.',
        location: loc,
        confidence: 'High',
      })
    }

    if (totalCells > 12 && emptyInside / totalCells > 0.35) {
      drafts.push({
        id: `empty-cells-${name}`,
        wcagKey: 'info-relationships',
        title: OFFICE_TITLES.emptyCellsInRange,
        severity: 'Medium',
        category: CAT_TABLES,
        affectedUsers: ['משתמשי קורא מסך'],
        impact: `ב"${name}" ${emptyInside} תאים ריקים בתוך טווח הנתונים.`,
        recommendation: 'צמצמו ריווח מלאכותי בתוך הטבלה.',
        location: loc,
        confidence: 'High',
      })
    }

    if (contrastIssues >= 2) {
      drafts.push({
        id: `contrast-${name}`,
        wcagKey: 'contrast-minimum',
        title: OFFICE_TITLES.hardTextContrast,
        severity: 'High',
        category: CAT_COLOR_CONTRAST,
        affectedUsers: ['משתמשים עם לקות ראייה'],
        impact: `ב"${name}" ${contrastIssues} תאים עם ניגודיות מתחת ל-4.5:1.`,
        recommendation: 'שפרו צבעי גופן ומילוי תאים.',
        location: loc,
        confidence: 'Low',
      })
    }
  })

  signals.mergedCellCount = mergedCellCount
  signals.genericSheetNames = genericSheetNames
  return { drafts, signals }
}
