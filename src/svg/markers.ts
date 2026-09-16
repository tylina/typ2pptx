import type { Element as XmlElement } from '@xmldom/xmldom'

import { normalizeSvgPath, parseSvgPath } from './path'

type MarkerSize = 'sm' | 'med' | 'lg'

interface MarkerClassification {
  width: MarkerSize
  length: MarkerSize
}

export function hasLineMarkers(element: XmlElement): boolean {
  return ['marker-start', 'marker-mid', 'marker-end'].some((name) => {
    const value = attributeValue(element, name)
    return Boolean(value && value.trim().toLocaleLowerCase() !== 'none')
  })
}

/** Return native line ends, or null when a requested SVG marker is not provable. */
export function drawingMlLineEnds(
  element: XmlElement,
  defs: ReadonlyMap<string, XmlElement>
): string | null {
  const middle = attributeValue(element, 'marker-mid')
  if (middle && middle.trim().toLocaleLowerCase() !== 'none') return null
  const head = lineEnd(element, defs, 'marker-start', 'headEnd', 'auto-start-reverse')
  const tail = lineEnd(element, defs, 'marker-end', 'tailEnd', 'auto')
  if (head === null || tail === null) return null
  return (head || '<a:headEnd type="none"/>') + (tail || '<a:tailEnd type="none"/>')
}

function lineEnd(
  element: XmlElement,
  defs: ReadonlyMap<string, XmlElement>,
  attribute: string,
  tag: 'headEnd' | 'tailEnd',
  requiredOrientation: 'auto' | 'auto-start-reverse'
): string | null {
  const source = attributeValue(element, attribute)
  if (!source || source.trim().toLocaleLowerCase() === 'none') return ''
  const reference = paintReference(source)
  const marker = reference ? defs.get(reference) : undefined
  if (!marker || localName(marker) !== 'marker') return null
  const classification = classifyMarker(marker, requiredOrientation)
  if (!classification) return null
  return `<a:${tag} type="triangle" w="${classification.width}" ` +
    `len="${classification.length}"/>`
}

function classifyMarker(
  marker: XmlElement,
  requiredOrientation: 'auto' | 'auto-start-reverse'
): MarkerClassification | null {
  if ((attributeValue(marker, 'orient') ?? '').trim() !== requiredOrientation) return null
  if ((marker.getAttribute('markerUnits') || 'strokeWidth') !== 'strokeWidth') return null
  if (marker.getAttribute('transform')) return null
  const viewBox = numberList(marker.getAttribute('viewBox'))
  if (viewBox.length !== 4 || viewBox[2]! <= 0 || viewBox[3]! <= 0) return null
  const refX = finiteNumber(marker.getAttribute('refX'))
  const refY = finiteNumber(marker.getAttribute('refY'))
  if (refX === null || refY === null) return null
  const children = childElements(marker)
  if (children.length !== 1) return null
  const child = children[0]!
  if (child.getAttribute('transform')) return null
  const fill = attributeValue(child, 'fill') || attributeValue(marker, 'fill')
  const stroke = attributeValue(child, 'stroke') || attributeValue(marker, 'stroke')
  if (fill?.trim().toLocaleLowerCase() !== 'context-stroke') return null
  if (stroke && stroke.trim().toLocaleLowerCase() !== 'none') return null
  const points = markerPoints(child)
  if (!points || !isAnchoredRightPointingTriangle(points, viewBox, refX, refY)) return null
  const [width, length] = markerSizes(marker)
  return { width, length }
}

function markerPoints(element: XmlElement): Array<readonly [number, number]> | null {
  const tag = localName(element)
  if (tag === 'polygon') {
    const values = numberList(element.getAttribute('points'))
    if (values.length !== 6) return null
    return [
      [values[0]!, values[1]!],
      [values[2]!, values[3]!],
      [values[4]!, values[5]!]
    ]
  }
  if (tag !== 'path') return null
  const value = element.getAttribute('d')
  if (!value) return null
  try {
    const commands = normalizeSvgPath(parseSvgPath(value))
    if (
      commands.length !== 4 ||
      commands[0]?.command !== 'M' ||
      commands[1]?.command !== 'L' ||
      commands[2]?.command !== 'L' ||
      commands[3]?.command !== 'Z'
    ) return null
    return commands.slice(0, 3).map((command) => [command.values[0]!, command.values[1]!])
  } catch {
    return null
  }
}

function isAnchoredRightPointingTriangle(
  points: ReadonlyArray<readonly [number, number]>,
  viewBox: readonly number[],
  refX: number,
  refY: number
): boolean {
  const [viewX, viewY, viewWidth, viewHeight] = viewBox as [number, number, number, number]
  const tolerance = Math.max(viewWidth, viewHeight) * 1e-6 + 1e-9
  if (points.some(([x, y]) =>
    x < viewX - tolerance ||
    x > viewX + viewWidth + tolerance ||
    y < viewY - tolerance ||
    y > viewY + viewHeight + tolerance)) return false
  const maximumX = Math.max(...points.map(([x]) => x))
  const tips = points.filter(([x]) => approximatelyEqual(x, maximumX, tolerance))
  if (tips.length !== 1) return false
  const tip = tips[0]!
  const base = points.filter((point) => point !== tip)
  if (
    base.length !== 2 ||
    !approximatelyEqual(base[0]![0], base[1]![0], tolerance) ||
    !(base[0]![0] < tip[0] - tolerance)
  ) return false
  const centerY = (base[0]![1] + base[1]![1]) / 2
  return approximatelyEqual(tip[0], refX, tolerance) &&
    approximatelyEqual(tip[1], refY, tolerance) &&
    approximatelyEqual(centerY, refY, tolerance)
}

function approximatelyEqual(left: number, right: number, tolerance: number): boolean {
  return Math.abs(left - right) <= tolerance
}

function markerSizes(marker: XmlElement): readonly [MarkerSize, MarkerSize] {
  const width = finitePositive(marker.getAttribute('markerWidth'), 3)
  const height = finitePositive(marker.getAttribute('markerHeight'), 3)
  return [relativeSize(height), relativeSize(width)]
}

function relativeSize(value: number): MarkerSize {
  return value <= 2 ? 'sm' : value >= 3.5 ? 'lg' : 'med'
}

function finitePositive(value: string | null, fallback: number): number {
  if (!value) return fallback
  const number = Number(value)
  return Number.isFinite(number) && number > 0 ? number : fallback
}

function finiteNumber(value: string | null): number | null {
  if (!value) return null
  const number = Number(value)
  return Number.isFinite(number) ? number : null
}

function numberList(value: string | null): number[] {
  return [...(value ?? '').matchAll(/[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?/giu)]
    .map((match) => Number(match[0]))
    .filter(Number.isFinite)
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
    if (node?.nodeType === 1) result.push(node as XmlElement)
  }
  return result
}

function localName(element: XmlElement): string {
  return element.localName || element.tagName.replace(/^.*:/u, '')
}
