import type { Element as XmlElement } from '@xmldom/xmldom'

import { parseColor } from './style'

const DRAWING_ML_EFFECT_SHAPES = new Set([
  'path', 'rect', 'circle', 'ellipse', 'line', 'polygon', 'polyline'
])

interface FilterParameters {
  standardDeviation: number
  dx: number
  dy: number
  opacity: number
  color: string
}

/** Return an effect list, an empty string for no effect, or null for an unsafe filter. */
export function drawingMlEffect(
  element: XmlElement,
  defs: ReadonlyMap<string, XmlElement>,
  emuPerUnit: number
): string | null {
  const source = attributeValue(element, 'filter')
  if (!source || source.trim().toLocaleLowerCase() === 'none') return ''
  if (!DRAWING_ML_EFFECT_SHAPES.has(localName(element))) return null
  const reference = paintReference(source)
  const filter = reference ? defs.get(reference) : undefined
  if (!filter || localName(filter) !== 'filter') return null
  if (['x', 'y', 'width', 'height', 'filterUnits', 'primitiveUnits'].some((name) =>
    filter.getAttribute(name))) return null
  const primitives = childElements(filter)
    .filter((child) => localName(child).startsWith('fe'))
  if (primitives.length !== 1 || localName(primitives[0]!) !== 'feDropShadow') return null
  const parameters = dropShadowParameters(primitives[0]!)
  if (!parameters) return null
  return shadowXml(parameters, emuPerUnit)
}

function dropShadowParameters(element: XmlElement): FilterParameters | null {
  const input = element.getAttribute('in')
  if (input && input !== 'SourceGraphic') return null
  const standardDeviation = deviation(element.getAttribute('stdDeviation'), 4)
  const dx = finiteAttribute(element, 'dx', 0)
  const dy = finiteAttribute(element, 'dy', 0)
  const opacity = unitAttribute(element, 'flood-opacity', 0.3)
  const color = effectColor(element.getAttribute('flood-color'))
  if (
    standardDeviation === null ||
    dx === null ||
    dy === null ||
    opacity === null ||
    !color
  ) {
    return null
  }
  return { standardDeviation, dx, dy, opacity, color }
}

function shadowXml(parameters: FilterParameters, emuPerUnit: number): string {
  const blur = toEmu(parameters.standardDeviation * 2, emuPerUnit)
  const distance = toEmu(Math.hypot(parameters.dx, parameters.dy), emuPerUnit)
  const direction = Math.round(
    ((Math.atan2(parameters.dy, parameters.dx) * 180 / Math.PI + 360) % 360) * 60_000
  )
  const alpha = Math.round(parameters.opacity * 100_000)
  return '<a:effectLst>' +
    `<a:outerShdw blurRad="${blur}" dist="${distance}" dir="${direction}" ` +
    `algn="${shadowAlignment(parameters.dx, parameters.dy)}" rotWithShape="0">` +
    `<a:srgbClr val="${parameters.color}"><a:alpha val="${alpha}"/>` +
    '</a:srgbClr></a:outerShdw></a:effectLst>'
}

function shadowAlignment(dx: number, dy: number): string {
  const threshold = 0.5
  if (Math.abs(dx) < threshold) return 'ctr'
  if (Math.abs(dy) < threshold) return dx > 0 ? 'l' : 'r'
  if (dx > 0 && dy > 0) return 'tl'
  if (dx < 0 && dy > 0) return 'tr'
  if (dx > 0 && dy < 0) return 'bl'
  return 'br'
}

function deviation(value: string | null, fallback: number): number | null {
  if (!value) return fallback
  const parts = value.trim().split(/[\s,]+/u).map(Number)
  return parts.length > 0 && parts.every((part) => Number.isFinite(part) && part >= 0)
    ? Math.max(...parts)
    : null
}

function finiteAttribute(
  element: XmlElement | undefined,
  name: string,
  fallback: number
): number | null {
  const value = element?.getAttribute(name)
  if (!value) return fallback
  const number = Number(value)
  return Number.isFinite(number) ? number : null
}

function unitAttribute(
  element: XmlElement | undefined,
  name: string,
  fallback: number
): number | null {
  const value = finiteAttribute(element, name, fallback)
  return value !== null && value >= 0 && value <= 1 ? value : null
}

function effectColor(value: string | null | undefined): string | null {
  return parseColor(value || '#000000')?.rgb ?? null
}

function attributeValue(element: XmlElement, name: string): string | null {
  const direct = element.getAttribute(name)
  if (direct) return direct
  for (const declaration of element.getAttribute('style')?.split(';') ?? []) {
    const separator = declaration.indexOf(':')
    if (separator <= 0) continue
    if (declaration.slice(0, separator).trim().toLocaleLowerCase() === name) {
      return declaration.slice(separator + 1).trim()
    }
  }
  return null
}

function paintReference(value: string): string | null {
  const match = value.trim().match(/^url\(\s*['"]?#([^'"\s)]+)['"]?\s*\)$/u)
  return match?.[1] ?? null
}

function childElements(element: XmlElement): XmlElement[] {
  const result: XmlElement[] = []
  for (let index = 0; index < element.childNodes.length; index += 1) {
    const node = element.childNodes.item(index)
    if (node?.nodeType !== 1) continue
    result.push(node as XmlElement)
  }
  return result
}

function localName(element: XmlElement): string {
  return element.localName || element.tagName.replace(/^.*:/u, '')
}

function toEmu(value: number, emuPerUnit: number): number {
  return Math.max(0, Math.round(value * emuPerUnit))
}
