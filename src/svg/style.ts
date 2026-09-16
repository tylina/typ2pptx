import type { Element as XmlElement } from '@xmldom/xmldom'

import { drawingMlLineEnds, hasLineMarkers } from './markers'
import {
  applyMatrix,
  IDENTITY_MATRIX,
  parseTransformList,
  type AffineMatrix
} from './matrix'

const NAMED_COLORS: Readonly<Record<string, string>> = {
  black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF',
  yellow: 'FFFF00', cyan: '00FFFF', magenta: 'FF00FF', orange: 'FFA500',
  purple: '800080', pink: 'FFC0CB', brown: 'A52A2A', gray: '808080', grey: '808080',
  silver: 'C0C0C0', gold: 'FFD700', navy: '000080', teal: '008080', maroon: '800000',
  olive: '808000', lime: '00FF00', aqua: '00FFFF', fuchsia: 'FF00FF',
  transparent: 'FFFFFF'
}

export interface SvgStyleContext {
  fill: string
  stroke: string
  strokeWidth: number
  opacity: number
  fillOpacity: number
  strokeOpacity: number
  lineCap: string
  lineJoin: string
  dashArray: string
  dashOffset: number
}

interface ResolvedGradient {
  kind: 'linearGradient' | 'radialGradient'
  stops: XmlElement[]
  units: string
  spread: string
  transform: AffineMatrix
  hasTransform: boolean
  attribute(name: string): string | null
}

export const DEFAULT_SVG_STYLE: SvgStyleContext = {
  fill: '#000000',
  stroke: 'none',
  strokeWidth: 1,
  opacity: 1,
  fillOpacity: 1,
  strokeOpacity: 1,
  lineCap: 'butt',
  lineJoin: 'miter',
  dashArray: 'none',
  dashOffset: 0
}

export function resolveStyle(
  element: XmlElement,
  parent: SvgStyleContext
): SvgStyleContext {
  const inline = parseInlineStyle(element.getAttribute('style'))
  const value = (name: string): string | null => element.getAttribute(name) || inline.get(name) || null
  return {
    fill: value('fill') ?? parent.fill,
    stroke: value('stroke') ?? parent.stroke,
    strokeWidth: finiteNonNegative(value('stroke-width'), parent.strokeWidth),
    opacity: parent.opacity * finiteUnit(value('opacity'), 1),
    fillOpacity: parent.fillOpacity * finiteUnit(value('fill-opacity'), 1),
    strokeOpacity: parent.strokeOpacity * finiteUnit(value('stroke-opacity'), 1),
    lineCap: value('stroke-linecap') ?? parent.lineCap,
    lineJoin: value('stroke-linejoin') ?? parent.lineJoin,
    dashArray: value('stroke-dasharray') ?? parent.dashArray,
    dashOffset: finiteNumber(value('stroke-dashoffset'), parent.dashOffset)
  }
}

export function supportsDrawingMlStyle(
  element: XmlElement,
  style: SvgStyleContext,
  defs: ReadonlyMap<string, XmlElement>
): boolean {
  const inline = parseInlineStyle(element.getAttribute('style'))
  const value = (name: string): string | null => element.getAttribute(name) || inline.get(name) || null
  if (element.getAttribute('class')) return false
  if (!supportedKeyword(value('display'), ['inline'])) return false
  if (!supportedKeyword(value('visibility'), ['visible'])) return false
  if (!supportedKeyword(value('fill-rule'), ['nonzero'])) return false
  if (!supportedKeyword(value('clip-rule'), ['nonzero'])) return false
  if (parseDashArray(style.dashArray) === null) return false
  if (Math.abs(style.dashOffset) >= 1e-10) return false
  if (!['butt', 'round', 'square'].includes(style.lineCap)) return false
  if (!['miter', 'round', 'bevel'].includes(style.lineJoin)) return false
  if (!supportedKeyword(value('stroke-miterlimit'), ['4'])) return false
  if (
    hasLineMarkers(element) &&
    (localName(element) !== 'line' || drawingMlLineEnds(element, defs) === null)
  ) return false
  if (!supportedKeyword(value('vector-effect'), ['none'])) return false
  if (!supportedKeyword(value('paint-order'), ['normal'])) return false
  if (!supportedKeyword(value('mix-blend-mode'), ['normal'])) return false
  if (!supportedKeyword(value('isolation'), ['auto'])) return false
  if (!supportedPaint(style.fill, defs) || !supportedPaint(style.stroke, defs)) return false
  return ['opacity', 'fill-opacity', 'stroke-opacity'].every((name) =>
    supportedOpacity(value(name)))
}

export function drawingMlFill(
  style: SvgStyleContext,
  defs: ReadonlyMap<string, XmlElement>
): string {
  if (style.fill === 'none') return '<a:noFill/>'
  const reference = paintReference(style.fill)
  if (reference) {
    const gradient = defs.get(reference)
    if (gradient && isGradient(gradient)) {
      return drawingMlGradient(gradient, defs, style.opacity * style.fillOpacity)
    }
    return '<a:noFill/>'
  }
  const color = parseColor(style.fill)
  if (!color) return '<a:noFill/>'
  return `<a:solidFill><a:srgbClr val="${color.rgb}">${alphaXml(
    color.alpha * style.opacity * style.fillOpacity
  )}</a:srgbClr></a:solidFill>`
}

export function drawingMlStroke(
  element: XmlElement,
  style: SvgStyleContext,
  defs: ReadonlyMap<string, XmlElement>,
  emuPerUnit: number
): string {
  if (style.stroke === 'none' || style.strokeWidth <= 0) return '<a:ln><a:noFill/></a:ln>'
  const reference = paintReference(style.stroke)
  const paint = reference ? defs.get(reference) : undefined
  const fill = paint && isGradient(paint)
    ? drawingMlGradient(paint, defs, style.opacity * style.strokeOpacity)
    : (() => {
        const color = parseColor(style.stroke)
        return color
          ? `<a:solidFill><a:srgbClr val="${color.rgb}">${alphaXml(
              color.alpha * style.opacity * style.strokeOpacity
            )}</a:srgbClr></a:solidFill>`
          : '<a:noFill/>'
      })()
  const cap = style.lineCap === 'round' ? 'rnd' : style.lineCap === 'square' ? 'sq' : 'flat'
  const join = style.lineJoin === 'round'
    ? '<a:round/>'
    : style.lineJoin === 'bevel'
      ? '<a:bevel/>'
      : '<a:miter lim="800000"/>'
  const dash = drawingMlDash(style.dashArray, style.strokeWidth)
  const lineEnds = drawingMlLineEnds(element, defs) ??
    '<a:headEnd type="none"/><a:tailEnd type="none"/>'
  return `<a:ln w="${Math.max(1, Math.round(style.strokeWidth * emuPerUnit))}" cap="${cap}">` +
    `${fill}${dash}${join}${lineEnds}</a:ln>`
}

function drawingMlDash(value: string, strokeWidth: number): string {
  const dash = parseDashArray(value)
  if (!dash || dash.length === 0) return ''
  const normalized = dash.map((part) => part / Math.max(strokeWidth, 0.001))
  if (approximatelyDash(normalized, [4, 4])) return '<a:prstDash val="dash"/>'
  if (approximatelyDash(normalized, [2, 2])) return '<a:prstDash val="sysDot"/>'
  if (approximatelyDash(normalized, [8, 4])) return '<a:prstDash val="lgDash"/>'
  if (approximatelyDash(normalized, [8, 4, 2, 4])) {
    return '<a:prstDash val="lgDashDot"/>'
  }
  const even = dash.length % 2 === 0 ? dash : [...dash, ...dash]
  const entries: string[] = []
  for (let index = 0; index < even.length; index += 2) {
    const length = Math.max(1, Math.round(even[index]! / Math.max(strokeWidth, 0.001) * 100_000))
    const space = Math.max(
      1,
      Math.round(even[index + 1]! / Math.max(strokeWidth, 0.001) * 100_000)
    )
    entries.push(`<a:ds d="${length}" sp="${space}"/>`)
  }
  return `<a:custDash>${entries.join('')}</a:custDash>`
}

function parseDashArray(value: string): number[] | null {
  const normalized = value.trim().toLocaleLowerCase()
  if (!normalized || normalized === 'none') return []
  const parts = normalized.split(/[\s,]+/u).filter(Boolean)
  const numbers = parts.map(Number)
  return numbers.length > 0 && numbers.every((part) => Number.isFinite(part) && part > 0)
    ? numbers
    : null
}

function approximatelyDash(actual: readonly number[], expected: readonly number[]): boolean {
  return actual.length === expected.length && actual.every((value, index) =>
    Math.abs(value - expected[index]!) <= 0.01)
}

export function parseColor(value: string): { rgb: string; alpha: number } | null {
  const normalized = value.trim().toLocaleLowerCase()
  if (!normalized || normalized === 'none') return null
  if (normalized === 'transparent') return { rgb: 'FFFFFF', alpha: 0 }
  if (normalized.startsWith('#')) {
    const hex = normalized.slice(1)
    if (/^[0-9a-f]{3}$/u.test(hex)) {
      return { rgb: [...hex].map((part) => part + part).join('').toLocaleUpperCase(), alpha: 1 }
    }
    if (/^[0-9a-f]{6}$/u.test(hex)) return { rgb: hex.toLocaleUpperCase(), alpha: 1 }
    if (/^[0-9a-f]{8}$/u.test(hex)) {
      return {
        rgb: hex.slice(0, 6).toLocaleUpperCase(),
        alpha: Number.parseInt(hex.slice(6), 16) / 255
      }
    }
    return null
  }
  const rgb = normalized.match(/^rgba?\(\s*([\d.]+)%?[,\s]+([\d.]+)%?[,\s]+([\d.]+)%?(?:[,/\s]+([\d.]+)%?)?\s*\)$/u)
  if (rgb) {
    const percentage = normalized.slice(0, normalized.indexOf(')')).includes('%')
    const channel = (index: number) => Math.max(0, Math.min(255, Math.round(
      Number(rgb[index] ?? 0) * (percentage ? 2.55 : 1)
    )))
    const alphaSource = rgb[4]
    const alpha = alphaSource === undefined
      ? 1
      : Math.max(0, Math.min(1, Number(alphaSource) / (normalized.includes(`${alphaSource}%`) ? 100 : 1)))
    return {
      rgb: [channel(1), channel(2), channel(3)]
        .map((part) => part.toString(16).padStart(2, '0'))
        .join('')
        .toLocaleUpperCase(),
      alpha
    }
  }
  const named = NAMED_COLORS[normalized]
  return named ? { rgb: named, alpha: normalized === 'transparent' ? 0 : 1 } : null
}

function drawingMlGradient(
  element: XmlElement,
  defs: ReadonlyMap<string, XmlElement>,
  opacity: number
): string {
  const gradient = resolveGradient(element, defs)
  if (!gradient || gradient.stops.length === 0) return '<a:noFill/>'
  const entries = gradient.stops.map((stop, index) => {
    const inline = parseInlineStyle(stop.getAttribute('style'))
    const color = parseColor(stop.getAttribute('stop-color') || inline.get('stop-color') || '#000000')
      ?? { rgb: '000000', alpha: 1 }
    const stopOpacity = finiteUnit(
      stop.getAttribute('stop-opacity') || inline.get('stop-opacity') || null,
      1
    )
    const offset = gradientOffset(stop.getAttribute('offset'), index, gradient.stops.length)
    return `<a:gs pos="${offset}"><a:srgbClr val="${color.rgb}">${alphaXml(
      color.alpha * opacity * stopOpacity
    )}</a:srgbClr></a:gs>`
  }).join('')
  if (gradient.kind === 'radialGradient') {
    return `<a:gradFill rotWithShape="1"><a:gsLst>${entries}</a:gsLst>` +
      '<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/>' +
      '</a:path></a:gradFill>'
  }
  const x1 = coordinate(gradient.attribute('x1'), 0)
  const y1 = coordinate(gradient.attribute('y1'), 0)
  const x2 = coordinate(gradient.attribute('x2'), 1)
  const y2 = coordinate(gradient.attribute('y2'), 0)
  const start = applyMatrix(gradient.transform, x1, y1)
  const end = applyMatrix(gradient.transform, x2, y2)
  const angle = Math.round((
    (Math.atan2(end[1] - start[1], end[0] - start[0]) * 180 / Math.PI + 360) % 360
  ) * 60_000)
  return `<a:gradFill rotWithShape="1"><a:gsLst>${entries}</a:gsLst>` +
    `<a:lin ang="${angle}" scaled="1"/></a:gradFill>`
}

function parseInlineStyle(value: string | null): Map<string, string> {
  const result = new Map<string, string>()
  for (const declaration of value?.split(';') ?? []) {
    const separator = declaration.indexOf(':')
    if (separator <= 0) continue
    result.set(
      declaration.slice(0, separator).trim().toLocaleLowerCase(),
      declaration.slice(separator + 1).trim()
    )
  }
  return result
}

function paintReference(value: string): string | null {
  const match = value.trim().match(/^url\(\s*['"]?#([^'"\s)]+)['"]?\s*\)$/u)
  return match?.[1] ?? null
}

function supportedPaint(value: string, defs: ReadonlyMap<string, XmlElement>): boolean {
  if (value === 'none' || parseColor(value)) return true
  const reference = paintReference(value)
  if (!reference) return false
  const gradient = defs.get(reference)
  return Boolean(gradient && isSupportedGradient(gradient, defs))
}

function isSupportedGradient(
  element: XmlElement,
  defs: ReadonlyMap<string, XmlElement>
): boolean {
  const gradient = resolveGradient(element, defs)
  if (!gradient || gradient.stops.length === 0 || gradient.spread !== 'pad') return false
  if (gradient.kind === 'radialGradient') {
    return gradient.units === 'objectBoundingBox' &&
      !gradient.hasTransform &&
      !['fx', 'fy', 'fr'].some((name) => gradient.attribute(name))
  }
  if (gradient.units === 'objectBoundingBox') return !gradient.hasTransform
  if (gradient.units !== 'userSpaceOnUse' || !isPositiveAxisScale(gradient.transform)) {
    return false
  }
  return [
    coordinate(gradient.attribute('x1'), 0),
    coordinate(gradient.attribute('y1'), 0),
    coordinate(gradient.attribute('x2'), 1),
    coordinate(gradient.attribute('y2'), 0)
  ].every((value) => value >= 0 && value <= 1)
}

function resolveGradient(
  element: XmlElement,
  defs: ReadonlyMap<string, XmlElement>
): ResolvedGradient | null {
  if (!isGradient(element)) return null
  const kind = localName(element) as ResolvedGradient['kind']
  const chain: XmlElement[] = []
  const visited = new Set<string>()
  let current: XmlElement | undefined = element
  while (current) {
    if (localName(current) !== kind) return null
    chain.push(current)
    const href = current.getAttribute('href') || current.getAttribute('xlink:href')
    if (!href) break
    if (!href.startsWith('#') || href.length === 1 || visited.has(href)) return null
    visited.add(href)
    current = defs.get(href.slice(1))
    if (!current) return null
  }
  const attribute = (name: string): string | null => {
    for (const candidate of chain) {
      const value = candidate.getAttribute(name)
      if (value) return value
    }
    return null
  }
  const stops = chain
    .map((candidate) => childElements(candidate)
      .filter((child) => localName(child) === 'stop'))
    .find((candidate) => candidate.length > 0) ?? []
  const transformSource = attribute('gradientTransform')
  let transform = IDENTITY_MATRIX
  try {
    transform = parseTransformList(transformSource)
  } catch {
    return null
  }
  return {
    kind,
    stops,
    units: attribute('gradientUnits') ?? 'objectBoundingBox',
    spread: attribute('spreadMethod') ?? 'pad',
    transform,
    hasTransform: Boolean(transformSource),
    attribute
  }
}

function isPositiveAxisScale(matrix: AffineMatrix): boolean {
  const [scaleX, skewY, skewX, scaleY, translateX, translateY] = matrix
  return scaleX > 0 && scaleY > 0 &&
    Math.abs(skewX) < 1e-10 &&
    Math.abs(skewY) < 1e-10 &&
    Math.abs(translateX) < 1e-10 &&
    Math.abs(translateY) < 1e-10
}

function supportedKeyword(value: string | null, allowed: readonly string[]): boolean {
  return value === null || allowed.includes(value.trim().toLocaleLowerCase())
}

function supportedOpacity(value: string | null): boolean {
  if (value === null) return true
  const number = Number.parseFloat(value)
  if (!Number.isFinite(number)) return false
  const normalized = value.includes('%') ? number / 100 : number
  return normalized >= 0 && normalized <= 1
}

function alphaXml(value: number): string {
  const alpha = Math.max(0, Math.min(1, value))
  return alpha < 0.99999 ? `<a:alpha val="${Math.round(alpha * 100_000)}"/>` : ''
}

function gradientOffset(value: string | null, index: number, length: number): number {
  if (!value) return length <= 1 ? 0 : Math.round(index * 100_000 / (length - 1))
  const number = Number.parseFloat(value)
  if (!Number.isFinite(number)) return 0
  return Math.round(Math.max(0, Math.min(1, value.includes('%') ? number / 100 : number)) * 100_000)
}

function coordinate(value: string | null, fallback: number): number {
  if (!value) return fallback
  const number = Number.parseFloat(value)
  if (!Number.isFinite(number)) return fallback
  return value.includes('%') ? number / 100 : number
}

function finiteUnit(value: string | null, fallback: number): number {
  if (value === null) return fallback
  const number = Number.parseFloat(value)
  return Number.isFinite(number) ? Math.max(0, Math.min(1, number)) : fallback
}

function finiteNonNegative(value: string | null, fallback: number): number {
  if (value === null) return fallback
  const number = Number.parseFloat(value)
  return Number.isFinite(number) ? Math.max(0, number) : fallback
}

function finiteNumber(value: string | null, fallback: number): number {
  if (value === null) return fallback
  const number = Number.parseFloat(value)
  return Number.isFinite(number) ? number : Number.NaN
}

function isGradient(element: XmlElement): boolean {
  return localName(element) === 'linearGradient' || localName(element) === 'radialGradient'
}

function localName(element: XmlElement): string {
  return element.localName || element.tagName.replace(/^.*:/u, '')
}

function childElements(element: XmlElement): XmlElement[] {
  const result: XmlElement[] = []
  for (let index = 0; index < element.childNodes.length; index += 1) {
    const node = element.childNodes.item(index)
    if (node?.nodeType === 1) result.push(node as XmlElement)
  }
  return result
}
