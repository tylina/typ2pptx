import {
  DOMParser,
  XMLSerializer,
  type Document as XmlDocument,
  type Element as XmlElement
} from '@xmldom/xmldom'

import {
  applyMatrix,
  IDENTITY_MATRIX,
  multiplyMatrices,
  parseTransformList,
  type AffineMatrix
} from './matrix'
import { drawingMlEffect } from './effects'
import { hasLineMarkers } from './markers'
import { pathToDrawingMlGeometry, toEmu } from './path'
import {
  DEFAULT_SVG_STYLE,
  drawingMlFill,
  drawingMlStroke,
  resolveStyle,
  supportsDrawingMlStyle,
  type SvgStyleContext
} from './style'

const SVG_NAMESPACE = 'http://www.w3.org/2000/svg'
const XLINK_NAMESPACE = 'http://www.w3.org/1999/xlink'
const MAX_SVG_BYTES = 16 * 1024 * 1024
const MAX_NATIVE_SHAPES = 100_000
const MAX_PAINT_LAYERS = 512
const MAX_RESIDUAL_LAYER_BYTES = 64 * 1024 * 1024
const RESIDUAL_LAYER_ATTRIBUTE = 'data-typ2pptx-residual-layer'
// PowerPoint ignores a fully transparent fill during edit-mode hit testing.
// 100 / 100_000 is 0.1% opacity: visually inert, but still selectable.
const GROUP_HIT_AREA_ALPHA = 100
const NON_VISUAL_TAGS = new Set([
  'defs', 'title', 'desc', 'metadata', 'style', 'linearGradient', 'radialGradient',
  'stop', 'clipPath', 'mask', 'filter', 'symbol'
])

export interface SvgDrawingMlOptions {
  emuPerUnit?: number
  firstShapeId?: number
  vectorGroups?: readonly SvgVectorGroup[]
}

export interface SvgVectorGroup {
  id: string
  kind: 'math'
  x: number
  y: number
  width: number
  height: number
}

export interface SvgDrawingMlResult {
  drawingMl: string
  residualSvg: string
  residualHasVisualContent: boolean
  layers: SvgDrawingMlLayer[]
  nativeShapeCount: number
  unsupportedElementCount: number
  warnings: string[]
}

export type SvgDrawingMlLayer =
  | { kind: 'drawingMl'; drawingMl: string; nativeShapeCount: number }
  | { kind: 'residualSvg'; svg: string; unsupportedElementCount: number }

interface ConversionState {
  document: XmlDocument
  defs: Map<string, XmlElement>
  emuPerUnit: number
  nextShapeId: number
  shapes: ConvertedShape[]
  paintLayers: PaintLayer[]
  nextResidualLayerId: number
  unsupportedElementCount: number
  warnings: Set<string>
}

type PaintLayer =
  | { kind: 'drawingMl'; shapes: ConvertedShape[] }
  | { kind: 'residualSvg'; marker: string; unsupportedElementCount: number }

interface ConversionContext {
  matrix: AffineMatrix
  style: SvgStyleContext
  fromUse: boolean
}

interface ShapeGeometry {
  x: number
  y: number
  width: number
  height: number
  xml: string
  name: string
  transformAttributes?: string
}

interface ConvertedShape {
  geometry: ShapeGeometry
  xml: string
  fromUse: boolean
}

export function convertSvgToDrawingMl(
  svg: string,
  options: SvgDrawingMlOptions = {}
): SvgDrawingMlResult {
  if (new TextEncoder().encode(svg).byteLength > MAX_SVG_BYTES) {
    throw new Error('SVG exceeds the 16 MB conversion budget')
  }
  if (/<!DOCTYPE|<!ENTITY/iu.test(svg)) throw new Error('SVG declarations are not allowed')
  const errors: string[] = []
  let document: XmlDocument
  try {
    document = new DOMParser({
      onError: (level, message) => {
        if (level !== 'warning') errors.push(message)
      }
    }).parseFromString(svg, 'image/svg+xml')
  } catch {
    throw new Error('Cannot convert invalid SVG XML')
  }
  if (errors.length > 0 || document.getElementsByTagName('parsererror').length > 0) {
    throw new Error('Cannot convert invalid SVG XML')
  }
  const root = document.documentElement
  if (!root || localName(root) !== 'svg') throw new Error('Cannot convert invalid SVG root')
  if (hasReservedLayerAttribute(document)) {
    throw new Error(`SVG contains reserved ${RESIDUAL_LAYER_ATTRIBUTE} metadata`)
  }
  const state: ConversionState = {
    document,
    defs: collectDefs(root),
    emuPerUnit: positiveFinite(options.emuPerUnit ?? 12_700, 'EMU scale'),
    nextShapeId: options.firstShapeId ?? 2,
    shapes: [],
    paintLayers: [],
    nextResidualLayerId: 1,
    unsupportedElementCount: 0,
    warnings: new Set()
  }
  processElement(root, {
    matrix: IDENTITY_MATRIX,
    style: DEFAULT_SVG_STYLE,
    fromUse: false
  }, state, true)
  const residualHasVisualContent = hasVisualContent(root)
  const layers = materializePaintLayers(document, state, options.vectorGroups ?? [])
  stripResidualLayerAttributes(document)
  return {
    drawingMl: layers
      .filter((layer) => layer.kind === 'drawingMl')
      .map((layer) => layer.drawingMl)
      .join(''),
    residualSvg: new XMLSerializer().serializeToString(document),
    residualHasVisualContent,
    layers,
    nativeShapeCount: state.shapes.length,
    unsupportedElementCount: state.unsupportedElementCount,
    warnings: [...state.warnings]
  }
}

function processChildren(
  parent: XmlElement,
  context: ConversionContext,
  state: ConversionState,
  mutate: boolean
): void {
  for (const child of childElements(parent)) {
    const converted = processElement(child, context, state, mutate)
    if (converted && mutate) parent.removeChild(child)
  }
}

function processElement(
  element: XmlElement,
  parent: ConversionContext,
  state: ConversionState,
  mutate: boolean
): boolean {
  const tag = localName(element)
  if (NON_VISUAL_TAGS.has(tag)) return false
  const style = resolveStyle(element, parent.style)
  const ownMatrix = parseTransformList(element.getAttribute('transform'))
  const matrix = multiplyMatrices(parent.matrix, ownMatrix)
  const context = { matrix, style, fromUse: parent.fromUse }
  const effect = drawingMlEffect(element, state.defs, state.emuPerUnit)
  if (hasClipOrMask(element) || effect === null) {
    return retainResidual(
      element,
      countVisualLeaves(element),
      'Some clipped, masked, or filtered SVG content remains in the fallback layer.',
      state,
      mutate
    )
  }
  if (!supportsDrawingMlStyle(element, style, state.defs)) {
    return retainResidual(
      element,
      countVisualLeaves(element),
      'Some SVG paint or stroke features remain in the fallback layer.',
      state,
      mutate
    )
  }
  if (tag === 'g' || tag === 'svg' || tag === 'a') {
    processChildren(element, context, state, mutate)
    return !hasVisualContent(element)
  }
  if (tag === 'use') return convertUse(element, context, state, mutate)

  const geometry = shapeGeometry(element, context, state)
  if (!geometry) {
    return retainResidual(
      element,
      1,
      `Unsupported SVG <${tag}> content remains in the fallback layer.`,
      state,
      mutate
    )
  }
  if (state.shapes.length >= MAX_NATIVE_SHAPES) {
    throw new Error(`SVG conversion exceeds the ${MAX_NATIVE_SHAPES}-shape budget`)
  }
  appendNativeShape(state, {
    geometry,
    xml: wrapShape(element, geometry, context.style, effect, state),
    fromUse: context.fromUse
  })
  return true
}

function convertUse(
  element: XmlElement,
  context: ConversionContext,
  state: ConversionState,
  mutate: boolean
): boolean {
  const href = element.getAttribute('href') || element.getAttributeNS(XLINK_NAMESPACE, 'href') ||
    element.getAttribute('xlink:href')
  if (!href?.startsWith('#')) throw new Error('external SVG references are not supported')
  const referenced = state.defs.get(href.slice(1))
  if (!referenced) {
    return retainResidual(
      element,
      1,
      'An unresolved SVG reference remains in the fallback layer.',
      state,
      mutate
    )
  }
  const x = finiteAttribute(element, 'x', 0)
  const y = finiteAttribute(element, 'y', 0)
  const translated: ConversionContext = {
    ...context,
    matrix: multiplyMatrices(context.matrix, [1, 0, 0, 1, x, y]),
    fromUse: true
  }
  const before = state.shapes.length
  const unsupportedBefore = state.unsupportedElementCount
  const shapeIdBefore = state.nextShapeId
  const layersBefore = clonePaintLayers(state.paintLayers)
  if (localName(referenced) === 'symbol' || localName(referenced) === 'g') {
    processChildren(referenced, translated, state, false)
  } else {
    processElement(referenced, translated, state, false)
  }
  if (state.unsupportedElementCount > unsupportedBefore) {
    state.shapes.splice(before)
    state.paintLayers = layersBefore
    state.nextShapeId = shapeIdBefore
    if (mutate) {
      appendResidualLayer(
        element,
        state.unsupportedElementCount - unsupportedBefore,
        state
      )
    }
    state.warnings.add('A partially supported SVG reference remains intact in the fallback layer.')
    return false
  }
  return state.shapes.length > before
}

function appendNativeShape(state: ConversionState, shape: ConvertedShape): void {
  state.shapes.push(shape)
  const last = state.paintLayers.at(-1)
  if (last?.kind === 'drawingMl') last.shapes.push(shape)
  else {
    assertPaintLayerBudget(state)
    state.paintLayers.push({ kind: 'drawingMl', shapes: [shape] })
  }
}

function retainResidual(
  element: XmlElement,
  count: number,
  warning: string,
  state: ConversionState,
  mutate: boolean
): false {
  state.unsupportedElementCount += count
  state.warnings.add(warning)
  if (mutate) appendResidualLayer(element, count, state)
  return false
}

function appendResidualLayer(
  element: XmlElement,
  count: number,
  state: ConversionState
): void {
  const last = state.paintLayers.at(-1)
  if (last?.kind === 'residualSvg') {
    element.setAttribute(RESIDUAL_LAYER_ATTRIBUTE, last.marker)
    last.unsupportedElementCount += count
    return
  }
  assertPaintLayerBudget(state)
  const marker = String(state.nextResidualLayerId++)
  element.setAttribute(RESIDUAL_LAYER_ATTRIBUTE, marker)
  state.paintLayers.push({ kind: 'residualSvg', marker, unsupportedElementCount: count })
}

function assertPaintLayerBudget(state: ConversionState): void {
  if (state.paintLayers.length >= MAX_PAINT_LAYERS) {
    throw new Error(`SVG conversion exceeds the ${MAX_PAINT_LAYERS}-paint-layer budget`)
  }
}

function clonePaintLayers(layers: readonly PaintLayer[]): PaintLayer[] {
  return layers.map((layer) => layer.kind === 'drawingMl'
    ? { kind: 'drawingMl', shapes: [...layer.shapes] }
    : { ...layer })
}

function materializePaintLayers(
  document: XmlDocument,
  state: ConversionState,
  groups: readonly SvgVectorGroup[]
): SvgDrawingMlLayer[] {
  let residualBytes = 0
  return state.paintLayers.map((layer) => {
    if (layer.kind === 'drawingMl') {
      return {
        kind: 'drawingMl',
        drawingMl: groupedDrawingMl(layer.shapes, groups, state),
        nativeShapeCount: layer.shapes.length
      }
    }
    const svg = serializeResidualLayer(document, layer.marker)
    residualBytes += new TextEncoder().encode(svg).byteLength
    if (residualBytes > MAX_RESIDUAL_LAYER_BYTES) {
      throw new Error('SVG residual layers exceed the 64 MB conversion budget')
    }
    return {
      kind: 'residualSvg',
      svg,
      unsupportedElementCount: layer.unsupportedElementCount
    }
  })
}

function serializeResidualLayer(document: XmlDocument, marker: string): string {
  const clone = document.cloneNode(true) as XmlDocument
  const marked = elementsWithResidualMarker(clone)
  for (const element of marked) {
    if (element.getAttribute(RESIDUAL_LAYER_ATTRIBUTE) === marker) {
      element.removeAttribute(RESIDUAL_LAYER_ATTRIBUTE)
    } else {
      element.parentNode?.removeChild(element)
    }
  }
  return new XMLSerializer().serializeToString(clone)
}

function stripResidualLayerAttributes(document: XmlDocument): void {
  for (const element of elementsWithResidualMarker(document)) {
    element.removeAttribute(RESIDUAL_LAYER_ATTRIBUTE)
  }
}

function elementsWithResidualMarker(document: XmlDocument): XmlElement[] {
  const result: XmlElement[] = []
  const elements = document.getElementsByTagName('*')
  for (let index = 0; index < elements.length; index += 1) {
    const element = elements.item(index) as XmlElement | null
    if (element?.hasAttribute(RESIDUAL_LAYER_ATTRIBUTE)) result.push(element)
  }
  return result
}

function hasReservedLayerAttribute(document: XmlDocument): boolean {
  return elementsWithResidualMarker(document).length > 0
}

function shapeGeometry(
  element: XmlElement,
  context: ConversionContext,
  state: ConversionState
): ShapeGeometry | null {
  const tag = localName(element)
  if (tag === 'path') {
    const path = element.getAttribute('d')
    if (!path) return null
    return customPath(path, context.matrix, state.emuPerUnit, 'SVG path')
  }
  if (tag === 'rect') return rectangleGeometry(element, context.matrix, state.emuPerUnit)
  if (tag === 'circle') {
    const radius = finiteAttribute(element, 'r', 0)
    return radius > 0
      ? ellipseGeometry(
          finiteAttribute(element, 'cx', 0) - radius,
          finiteAttribute(element, 'cy', 0) - radius,
          radius * 2,
          radius * 2,
          context.matrix,
          state.emuPerUnit
        )
      : null
  }
  if (tag === 'ellipse') {
    const rx = finiteAttribute(element, 'rx', 0)
    const ry = finiteAttribute(element, 'ry', 0)
    return rx > 0 && ry > 0
      ? ellipseGeometry(
          finiteAttribute(element, 'cx', 0) - rx,
          finiteAttribute(element, 'cy', 0) - ry,
          rx * 2,
          ry * 2,
          context.matrix,
          state.emuPerUnit
        )
      : null
  }
  if (tag === 'line') {
    if (hasLineMarkers(element)) {
      return presentationLineGeometry(element, context.matrix, state.emuPerUnit)
    }
    return customPath(
      `M ${finiteAttribute(element, 'x1', 0)} ${finiteAttribute(element, 'y1', 0)} ` +
        `L ${finiteAttribute(element, 'x2', 0)} ${finiteAttribute(element, 'y2', 0)}`,
      context.matrix,
      state.emuPerUnit,
      'SVG line'
    )
  }
  if (tag === 'polygon' || tag === 'polyline') {
    const numbers = parseNumberList(element.getAttribute('points'))
    if (numbers.length < 4 || numbers.length % 2 !== 0) return null
    let path = `M ${numbers[0]} ${numbers[1]}`
    for (let index = 2; index < numbers.length; index += 2) {
      path += ` L ${numbers[index]} ${numbers[index + 1]}`
    }
    if (tag === 'polygon') path += ' Z'
    return customPath(path, context.matrix, state.emuPerUnit, `SVG ${tag}`)
  }
  return null
}

function presentationLineGeometry(
  element: XmlElement,
  matrix: AffineMatrix,
  emuPerUnit: number
): ShapeGeometry {
  const first = applyMatrix(
    matrix,
    finiteAttribute(element, 'x1', 0),
    finiteAttribute(element, 'y1', 0)
  )
  const second = applyMatrix(
    matrix,
    finiteAttribute(element, 'x2', 0),
    finiteAttribute(element, 'y2', 0)
  )
  const minimum = 1 / emuPerUnit
  const flipH = first[0] > second[0]
  const flipV = first[1] > second[1]
  return {
    x: Math.min(first[0], second[0]),
    y: Math.min(first[1], second[1]),
    width: Math.max(Math.abs(second[0] - first[0]), minimum),
    height: Math.max(Math.abs(second[1] - first[1]), minimum),
    xml: '<a:prstGeom prst="line"><a:avLst/></a:prstGeom>',
    name: 'SVG arrow line',
    transformAttributes: `${flipH ? ' flipH="1"' : ''}${flipV ? ' flipV="1"' : ''}`
  }
}

function rectangleGeometry(
  element: XmlElement,
  matrix: AffineMatrix,
  emuPerUnit: number
): ShapeGeometry | null {
  const x = finiteAttribute(element, 'x', 0)
  const y = finiteAttribute(element, 'y', 0)
  const width = finiteAttribute(element, 'width', 0)
  const height = finiteAttribute(element, 'height', 0)
  if (width <= 0 || height <= 0) return null
  const rxSource = element.getAttribute('rx')
  const rySource = element.getAttribute('ry')
  const rx = rxSource ? finiteAttribute(element, 'rx', 0) : rySource ? finiteAttribute(element, 'ry', 0) : 0
  const ry = rySource ? finiteAttribute(element, 'ry', 0) : rx
  if (isAxisAligned(matrix)) {
    const [left, top] = applyMatrix(matrix, x, y)
    const [right, bottom] = applyMatrix(matrix, x + width, y + height)
    const shapeWidth = Math.abs(right - left)
    const shapeHeight = Math.abs(bottom - top)
    const preset = rx > 0 && ry > 0 ? 'roundRect' : 'rect'
    return {
      x: Math.min(left, right),
      y: Math.min(top, bottom),
      width: shapeWidth,
      height: shapeHeight,
      xml: `<a:prstGeom prst="${preset}"><a:avLst/></a:prstGeom>`,
      name: 'SVG rectangle'
    }
  }
  return customPath(
    `M ${x} ${y} L ${x + width} ${y} L ${x + width} ${y + height} ` +
      `L ${x} ${y + height} Z`,
    matrix,
    emuPerUnit,
    'SVG rectangle'
  )
}

function ellipseGeometry(
  x: number,
  y: number,
  width: number,
  height: number,
  matrix: AffineMatrix,
  emuPerUnit: number
): ShapeGeometry {
  if (isAxisAligned(matrix)) {
    const [left, top] = applyMatrix(matrix, x, y)
    const [right, bottom] = applyMatrix(matrix, x + width, y + height)
    return {
      x: Math.min(left, right),
      y: Math.min(top, bottom),
      width: Math.abs(right - left),
      height: Math.abs(bottom - top),
      xml: '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>',
      name: 'SVG ellipse'
    }
  }
  const cx = x + width / 2
  const cy = y + height / 2
  const rx = width / 2
  const ry = height / 2
  return customPath(
    `M ${cx - rx} ${cy} A ${rx} ${ry} 0 1 0 ${cx + rx} ${cy} ` +
      `A ${rx} ${ry} 0 1 0 ${cx - rx} ${cy} Z`,
    matrix,
    emuPerUnit,
    'SVG ellipse'
  )
}

function customPath(
  value: string,
  matrix: AffineMatrix,
  emuPerUnit: number,
  name: string
): ShapeGeometry {
  const geometry = pathToDrawingMlGeometry(value, { transform: matrix, emuPerUnit })
  return {
    ...geometry.bounds,
    xml: '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>' +
      '<a:rect l="l" t="t" r="r" b="b"/><a:pathLst>' +
      `${geometry.xml}</a:pathLst></a:custGeom>`,
    name
  }
}

function wrapShape(
  element: XmlElement,
  geometry: ShapeGeometry,
  style: SvgStyleContext,
  effect: string,
  state: ConversionState
): string {
  const id = state.nextShapeId++
  const fill = drawingMlFill(style, state.defs)
  const stroke = drawingMlStroke(element, style, state.defs, state.emuPerUnit)
  return '<p:sp><p:nvSpPr>' +
    `<p:cNvPr id="${id}" name="${escapeXml(`${geometry.name} ${id}`)}"/>` +
    `<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm${geometry.transformAttributes ?? ''}>` +
    `<a:off x="${toEmu(geometry.x, state.emuPerUnit)}" y="${toEmu(
      geometry.y,
      state.emuPerUnit
    )}"/><a:ext cx="${toEmu(geometry.width, state.emuPerUnit)}" cy="${toEmu(
      geometry.height,
      state.emuPerUnit
    )}"/></a:xfrm>${geometry.xml}${fill}${stroke}${effect}</p:spPr></p:sp>`
}

function groupedDrawingMl(
  shapes: readonly ConvertedShape[],
  groups: readonly SvgVectorGroup[],
  state: ConversionState
): string {
  const membership = new Map<number, number>()
  const membersByGroup = new Map<number, number[]>()
  for (const [groupIndex, group] of groups.entries()) {
    validateVectorGroup(group)
    const members: number[] = []
    for (const [shapeIndex, shape] of shapes.entries()) {
      if (membership.has(shapeIndex) || !shape.fromUse) continue
      const centerX = shape.geometry.x + shape.geometry.width / 2
      const centerY = shape.geometry.y + shape.geometry.height / 2
      if (
        centerX >= group.x && centerX <= group.x + group.width &&
        centerY >= group.y && centerY <= group.y + group.height
      ) {
        membership.set(shapeIndex, groupIndex)
        members.push(shapeIndex)
      }
    }
    if (members.length > 0) membersByGroup.set(groupIndex, members)
  }

  const emittedGroups = new Set<number>()
  const output: string[] = []
  for (const [shapeIndex, shape] of shapes.entries()) {
    const groupIndex = membership.get(shapeIndex)
    if (groupIndex === undefined) {
      output.push(shape.xml)
      continue
    }
    if (emittedGroups.has(groupIndex)) continue
    emittedGroups.add(groupIndex)
    const group = groups[groupIndex]!
    const members = membersByGroup.get(groupIndex)!.map((index) => shapes[index]!)
    output.push(wrapGroup(group, members, state))
  }
  return output.join('')
}

function wrapGroup(
  group: SvgVectorGroup,
  shapes: readonly ConvertedShape[],
  state: ConversionState
): string {
  const left = Math.min(group.x, ...shapes.map((shape) => shape.geometry.x))
  const top = Math.min(group.y, ...shapes.map((shape) => shape.geometry.y))
  const right = Math.max(
    group.x + group.width,
    ...shapes.map((shape) => shape.geometry.x + shape.geometry.width)
  )
  const bottom = Math.max(
    group.y + group.height,
    ...shapes.map((shape) => shape.geometry.y + shape.geometry.height)
  )
  const x = toEmu(left, state.emuPerUnit)
  const y = toEmu(top, state.emuPerUnit)
  const width = Math.max(1, toEmu(right - left, state.emuPerUnit))
  const height = Math.max(1, toEmu(bottom - top, state.emuPerUnit))
  const id = state.nextShapeId++
  return '<p:grpSp><p:nvGrpSpPr>' +
    `<p:cNvPr id="${id}" name="${escapeXml(`Typst ${group.kind} ${group.id}`)}"/>` +
    '<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm>' +
    `<a:off x="${x}" y="${y}"/><a:ext cx="${width}" cy="${height}"/>` +
    `<a:chOff x="${x}" y="${y}"/><a:chExt cx="${width}" cy="${height}"/>` +
    `</a:xfrm></p:grpSpPr>${groupHitArea(group, state)}` +
    `${shapes.map((shape) => shape.xml).join('')}</p:grpSp>`
}

function groupHitArea(group: SvgVectorGroup, state: ConversionState): string {
  const id = state.nextShapeId++
  return '<p:sp><p:nvSpPr>' +
    `<p:cNvPr id="${id}" name="${escapeXml(`Typst ${group.kind} hit area ${group.id}`)}"/>` +
    '<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm>' +
    `<a:off x="${toEmu(group.x, state.emuPerUnit)}" ` +
    `y="${toEmu(group.y, state.emuPerUnit)}"/>` +
    `<a:ext cx="${toEmu(group.width, state.emuPerUnit)}" ` +
    `cy="${toEmu(group.height, state.emuPerUnit)}"/>` +
    '</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
    `<a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="${GROUP_HIT_AREA_ALPHA}"/>` +
    '</a:srgbClr></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>'
}

function validateVectorGroup(group: SvgVectorGroup): void {
  if (!group.id || group.id.length > 256 || group.kind !== 'math') {
    throw new Error('SVG vector group has invalid identity')
  }
  for (const value of [group.x, group.y, group.width, group.height]) {
    if (!Number.isFinite(value)) throw new Error('SVG vector group has invalid bounds')
  }
  if (group.width <= 0 || group.height <= 0) {
    throw new Error('SVG vector group has invalid bounds')
  }
}

function collectDefs(root: XmlElement): Map<string, XmlElement> {
  const result = new Map<string, XmlElement>()
  const visit = (element: XmlElement) => {
    const id = element.getAttribute('id')
    if (id) result.set(id, element)
    for (const child of childElements(element)) visit(child)
  }
  visit(root)
  return result
}

function hasClipOrMask(element: XmlElement): boolean {
  return ['clip-path', 'mask'].some((name) => Boolean(element.getAttribute(name)))
}

function countVisualLeaves(element: XmlElement): number {
  const children = childElements(element).filter((child) => !NON_VISUAL_TAGS.has(localName(child)))
  if (children.length === 0) return NON_VISUAL_TAGS.has(localName(element)) ? 0 : 1
  return children.reduce((sum, child) => sum + countVisualLeaves(child), 0)
}

function hasVisualContent(element: XmlElement): boolean {
  for (const child of childElements(element)) {
    const tag = localName(child)
    if (NON_VISUAL_TAGS.has(tag)) continue
    if (tag === 'g' || tag === 'svg' || tag === 'a') {
      if (hasVisualContent(child)) return true
      continue
    }
    return true
  }
  return false
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

function parseNumberList(value: string | null): number[] {
  return [...(value ?? '').matchAll(/[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?/giu)]
    .map((match) => Number(match[0]))
    .filter(Number.isFinite)
}

function finiteAttribute(element: XmlElement, name: string, fallback: number): number {
  const value = element.getAttribute(name)
  if (!value) return fallback
  const number = Number.parseFloat(value)
  if (!Number.isFinite(number)) throw new Error(`SVG <${localName(element)}> has invalid ${name}`)
  return number
}

function isAxisAligned(matrix: AffineMatrix): boolean {
  return Math.abs(matrix[1]) < 1e-10 && Math.abs(matrix[2]) < 1e-10
}

function positiveFinite(value: number, label: string): number {
  if (!Number.isFinite(value) || value <= 0) throw new Error(`Invalid ${label}`)
  return value
}

function escapeXml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
}
