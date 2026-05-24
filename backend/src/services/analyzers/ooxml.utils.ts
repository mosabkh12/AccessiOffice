import AdmZip from 'adm-zip'
import { XMLParser } from 'fast-xml-parser'

export const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  removeNSPrefix: true,
  isArray: (name) => ['p', 'r', 't', 'sp', 'pic', 'tbl', 'tr', 'tc', 'hyperlink'].includes(name),
})

export const UNCLEAR_LINK_PATTERNS = [
  /לחץ כאן/i,
  /^כאן$/i,
  /^here$/i,
  /click here/i,
  /read more/i,
  /more info/i,
]

export function readZip(filePath: string): AdmZip {
  return new AdmZip(filePath)
}

export function readZipEntryText(zip: AdmZip, entryPath: string): string | null {
  const entry = zip.getEntry(entryPath)
  if (!entry) return null
  return entry.getData().toString('utf8')
}

export function listZipEntries(zip: AdmZip, pattern: RegExp): string[] {
  return zip
    .getEntries()
    .map((e: { entryName: string }) => e.entryName)
    .filter((name: string) => pattern.test(name))
    .sort((a, b) => {
      const na = parseInt(a.match(/(\d+)/)?.[1] ?? '0', 10)
      const nb = parseInt(b.match(/(\d+)/)?.[1] ?? '0', 10)
      return na - nb
    })
}

export function parseXml<T = unknown>(xml: string): T {
  return xmlParser.parse(xml) as T
}

export function collectText(node: unknown): string {
  if (node == null) return ''
  if (typeof node === 'string' || typeof node === 'number') return String(node)
  if (Array.isArray(node)) return node.map(collectText).join('')
  if (typeof node === 'object') {
    const obj = node as Record<string, unknown>
    if ('t' in obj) return collectText(obj.t)
    // Skip XML attribute values (prefixed @_) — they are not text content
    return Object.entries(obj)
      .filter(([k]) => !k.startsWith('@_'))
      .map(([, v]) => collectText(v))
      .join('')
  }
  return ''
}

export function isUnclearLink(text: string): boolean {
  const t = text.trim()
  if (!t || t.length > 80) return false
  return UNCLEAR_LINK_PATTERNS.some((p) => p.test(t))
}

export function parseSrgbHex(color?: string): { r: number; g: number; b: number } | null {
  if (!color) return null
  const hex = color.replace(/^#/, '').trim()
  if (hex.length === 6) {
    return {
      r: parseInt(hex.slice(0, 2), 16),
      g: parseInt(hex.slice(2, 4), 16),
      b: parseInt(hex.slice(4, 6), 16),
    }
  }
  return null
}

function luminance(r: number, g: number, b: number): number {
  const lin = [r, g, b].map((c) => {
    const s = c / 255
    return s <= 0.03928 ? s / 12.92 : ((s + 0.055) / 1.055) ** 2.4
  })
  return 0.2126 * lin[0] + 0.7152 * lin[1] + 0.0722 * lin[2]
}

export function contrastRatio(fg: string, bg: string): number | null {
  const f = parseSrgbHex(fg)
  const b = parseSrgbHex(bg)
  if (!f || !b) return null
  const l1 = luminance(f.r, f.g, f.b)
  const l2 = luminance(b.r, b.g, b.b)
  return (Math.max(l1, l2) + 0.05) / (Math.min(l1, l2) + 0.05)
}

export function isGenericSheetName(name: string): boolean {
  return /^(sheet\d+|גיליון\d*|גליון\d*)$/i.test(name.trim())
}
